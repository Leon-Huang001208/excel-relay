"""macOS Excel automation for a staged workbook copy."""

from __future__ import annotations

import json
import logging
import math
import os
import plistlib
import posixpath
import re
import subprocess
import sys
import tempfile
import time
import zipfile
from contextlib import contextmanager
from pathlib import Path
from typing import TYPE_CHECKING, Any, Iterator
from urllib.parse import unquote, urlsplit
from xml.etree import ElementTree as ET

if TYPE_CHECKING:
    from excel_relay.core.models import ParameterDef

_WIND_ADDIN = Path("/Library/WindFunc/windfunc.xlam")
_WIND_LOG_DIR = Path("/Library/OfficeWind/WindAddin/log")
_SHEET_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_REL_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
_WIND_FUNCTION = re.compile(
    r"(?i)(?:\bwindfunc\.(?:xlam|xla|xll)['\]]?!|"
    r"\b(?:wsd|wss|wsq|wset|wsi|wst|wtd|wpf|tdays|tdaysoffset|tdayscount|edb|s_wq_[a-z_]+|s_info_name)\s*\()"
)
_IFIND_FUNCTION = re.compile(r"(?i)\b(?:ths(?:_\w+)?\s*\(|thsxls)")

_AUTOMATION = Path(__file__).with_suffix(".jxa")
logger = logging.getLogger("excel-relay.macos_excel")


class ExcelBusyError(RuntimeError):
    """Another background Excel job is running or awaiting safe cleanup."""


class WorkerCleanupError(RuntimeError):
    """The task workbook may still be open, so its staging file must be kept."""


def probe_environment() -> dict[str, Any]:
    """Check installation without opening Excel or starting a Wind session."""
    excel_version = None
    excel_app = Path("/Applications/Microsoft Excel.app")
    if not excel_app.is_dir():
        excel_app = Path.home() / "Applications" / "Microsoft Excel.app"
    if excel_app.is_dir():
        try:
            with (excel_app / "Contents/Info.plist").open("rb") as stream:
                excel_version = plistlib.load(stream).get("CFBundleShortVersionString")
        except (OSError, plistlib.InvalidFileException):
            pass

    return {
        "platform": sys.platform,
        "python_version": ".".join(str(part) for part in sys.version_info[:3]),
        "excel_ok": excel_app.is_dir(),
        "excel_version": excel_version,
        "excel_error": None if excel_app.is_dir() else "未找到 macOS 版 Microsoft Excel",
        "plugins": ["Wind (已安装)"] if _WIND_ADDIN.is_file() else [],
        "backend": "macos-excel",
        "automation_verified": False,
    }


def _lock_path() -> Path:
    directory = subprocess.check_output(
        ["/usr/bin/getconf", "DARWIN_USER_TEMP_DIR"], text=True, timeout=3,
    ).strip()
    if not directory:
        raise RuntimeError("无法定位 macOS 用户临时目录。")
    return Path(directory) / f"excel-relay-excel-{os.getuid()}.lock"


def _run_automation(mode: str, data: dict[str, Any] | None = None) -> Any:
    args = ["/usr/bin/osascript", "-l", "JavaScript", str(_AUTOMATION), mode]
    with tempfile.TemporaryDirectory(prefix="excel-relay-event-") as directory:
        if data is not None:
            payload = Path(directory) / "request.json"
            payload.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
            args.append(str(payload))
        timeout = data["seconds"] + 2 if data else 4
        result = subprocess.run(args, capture_output=True, text=True, timeout=timeout)
    if result.returncode:
        raise RuntimeError(result.stderr.strip() or "Excel 后台操作失败。")
    return json.loads(result.stdout)


def _worker_marker() -> Path:
    return _lock_path().with_suffix(".worker.json")


def _valid_worker(value: Any) -> bool:
    return (
        isinstance(value, dict)
        and set(value) == {"pid", "started"}
        and type(value["pid"]) is int
        and value["pid"] > 0
        and type(value["started"]) in (float, int)
        and math.isfinite(value["started"])
        and value["started"] > 0
    )


def _check_previous_worker(marker: Path) -> None:
    if not marker.exists():
        return
    try:
        record = json.loads(marker.read_text(encoding="utf-8"))
        if not isinstance(record["path"], str) or not Path(record["path"]).is_absolute():
            raise ValueError("Invalid recovery path")
        worker = record["worker"]
        if not _valid_worker(worker):
            raise ValueError("Unconfirmed worker identity")
        active = _run_automation("snapshot")
        if not isinstance(active, list) or not all(_valid_worker(app) for app in active):
            raise ValueError("Invalid Excel process snapshot")
    except (KeyError, TypeError, ValueError, OverflowError, OSError, RuntimeError, subprocess.TimeoutExpired) as error:
        raise ExcelBusyError(f"无法确认上次 Excel 刷新已清理。恢复记录：{marker}") from error
    if worker in active:
        raise ExcelBusyError(f"上次刷新进程仍在运行。待恢复副本：{record['path']}")
    marker.unlink()


def _hold_lock_until_worker_exits(marker: Path, lock_fd: int) -> None:
    """Continue holding the inherited legacy lock after incomplete cleanup."""
    try:
        while marker.exists():
            try:
                _check_previous_worker(marker)
            except ExcelBusyError:
                time.sleep(1)
    finally:
        os.close(lock_fd)


def _retain_worker_lock(marker: Path, lock_fd: int) -> subprocess.Popen:
    command = [
        sys.executable,
        "-B",
        "-c",
        "from pathlib import Path; import sys; "
        "from excel_relay.core.macos_excel import _hold_lock_until_worker_exits; "
        "_hold_lock_until_worker_exits(Path(sys.argv[1]), int(sys.argv[2]))",
        str(marker),
        str(lock_fd),
    ]
    return subprocess.Popen(
        command, pass_fds=(lock_fd,), start_new_session=True,
        stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )


@contextmanager
def excel_refresh_lock() -> Iterator[None]:
    """Serialize Mac refreshes using the standard per-user directory."""
    import fcntl

    lock_path = _lock_path()
    fd = os.open(lock_path, os.O_CREAT | os.O_WRONLY | os.O_NOFOLLOW, 0o600)
    with os.fdopen(fd, "wb") as lock:
        try:
            fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError as error:
            raise ExcelBusyError("Excel Relay 正在刷新另一份工作簿，请在完成后重试。") from error
        marker = lock_path.with_suffix(".worker.json")
        _check_previous_worker(marker)
        try:
            yield
        finally:
            if marker.exists():
                try:
                    _retain_worker_lock(marker, lock.fileno())
                except OSError as error:
                    logger.error("无法保持兼容锁，可能仍需恢复 Excel 实例：%s", error)


def _relationships(archive: zipfile.ZipFile, path: str) -> dict[str, str]:
    return {
        relation.attrib["Id"]: relation.attrib["Target"]
        for relation in ET.fromstring(archive.read(path))
    }


def _member_path(directory: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(directory, target))


def _is_installed_wind_link(workbook: Path, target: str) -> bool:
    link = urlsplit(unquote(target))
    if link.scheme not in ("", "file") or link.netloc or link.query or link.fragment:
        return False
    target_path = workbook.parent / link.path
    return target_path.resolve() == _WIND_ADDIN.resolve()


def _wind_cells(path: Path) -> tuple[list[str], dict[tuple[str, str], tuple[str | None, str | None]]]:
    with zipfile.ZipFile(path) as archive:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        wb_rels = _relationships(archive, "xl/_rels/workbook.xml.rels")
        wind_links: set[int] = set()
        for i, reference in enumerate(workbook.findall(f".//{_SHEET_NS}externalReference"), 1):
            link = _member_path("xl", wb_rels[reference.attrib[_REL_ID]])
            base, name = posixpath.split(link)
            rels = _relationships(archive, posixpath.join(base, "_rels", f"{name}.rels"))
            external_book = ET.fromstring(archive.read(link)).find(f"{_SHEET_NS}externalBook")
            if external_book is None or not _is_installed_wind_link(path, rels[external_book.attrib[_REL_ID]]):
                raise RuntimeError("请先在 Excel 中将底稿外链绑定到本机 Wind 加载项。")
            for target in rels.values():
                if _is_installed_wind_link(path, target):
                    continue
                # Excel stores a relative alternate beside the absolute add-in link.
                # A copied workbook can be deeper than its original location.
                relative = unquote(target)
                alternate = path.parent / relative
                if (
                    re.fullmatch(r"(?:\.\./)+Library/WindFunc/windfunc\.xlam", relative)
                    and not alternate.exists()
                    and not alternate.is_symlink()
                ):
                    continue
                raise RuntimeError("macOS 刷新仅支持本机 Wind 工作簿外链。")
            wind_links.add(i)
        names = archive.namelist()
        if "xl/connections.xml" in names:
            raise RuntimeError("这份工作簿包含 Excel 外部数据连接，当前 macOS 刷新未覆盖此连接。")
        if any(name.endswith("/vbaProject.bin") or name.startswith("xl/macrosheets/") for name in names):
            raise RuntimeError("macOS 刷新需要不包含工作簿宏的 .xlsx 副本。")

        sheet_names: list[str] = []
        cells: dict[tuple[str, str], tuple[str | None, str | None]] = {}
        for sheet in workbook.findall(f".//{_SHEET_NS}sheet"):
            name = sheet.attrib["name"]
            member = _member_path("xl", wb_rels[sheet.attrib[_REL_ID]])
            if not member.startswith("xl/worksheets/"):
                continue
            sheet_names.append(name)
            if name.casefold() == "_excel_relay":
                continue
            for cell in ET.fromstring(archive.read(member)).iter(f"{_SHEET_NS}c"):
                formula = cell.findtext(f"{_SHEET_NS}f")
                if not formula:
                    continue
                if _IFIND_FUNCTION.search(formula):
                    raise RuntimeError("macOS 刷新暂不支持包含 iFinD 公式的工作簿。")
                used_links = {int(i) for i in re.findall(r"\[(\d+)\]", formula)}
                if used_links & wind_links or _WIND_FUNCTION.search(formula):
                    cells[(name, cell.attrib["r"])] = (
                        cell.get("t"),
                        cell.findtext(f"{_SHEET_NS}v"),
                    )
        return sheet_names, cells


def _parameter_args(
    params: list[tuple[ParameterDef, Any]], sheet_names: list[str]
) -> tuple[list[tuple[str, str, str, str]], set[tuple[str, str]]]:
    by_name = {name.casefold(): name for name in sheet_names}
    data_sheets = [name for name in sheet_names if name.casefold() != "_excel_relay"]
    args: list[tuple[str, str, str, str]] = []
    targets: set[tuple[str, str]] = set()
    for param, value in params:
        ref = param.cell.strip()
        sheet, sep, cell = ref.rpartition("!")
        if sep:
            sheet = sheet.strip()
            if sheet.startswith("'") and sheet.endswith("'"):
                sheet = sheet[1:-1].replace("''", "'")
            actual_sheet = by_name.get(sheet.casefold())
        else:
            actual_sheet = data_sheets[0] if data_sheets else None
        cell = cell.replace("$", "").strip().upper()
        if (
            actual_sheet is None
            or actual_sheet.casefold() == "_excel_relay"
            or not re.fullmatch(r"[A-Z]{1,3}[1-9]\d{0,6}", cell)
        ):
            raise RuntimeError(f"无法定位已登记的参数坐标 {param.cell}。")

        if value is None:
            scalar_type, serialized = "empty", ""
        elif isinstance(value, bool):
            scalar_type, serialized = "boolean", str(value).lower()
        elif isinstance(value, (int, float)) and math.isfinite(value):
            scalar_type, serialized = "number", str(value)
        elif isinstance(value, str) and "\x00" not in value:
            scalar_type, serialized = "text", value
        else:
            raise RuntimeError(f"参数 {param.name} 需要是 Excel 可写入的标量。")
        args.append((actual_sheet, cell, scalar_type, serialized))
        targets.add((actual_sheet, cell))
    return args, targets


def _wind_logs_for_pid(pid: int) -> set[Path]:
    try:
        result = subprocess.run(
            ["/usr/sbin/lsof", "-nP", "-b", "-a", "-p", str(pid), "-Fn"],
            capture_output=True, text=True, timeout=10,
        )
    except subprocess.TimeoutExpired as error:
        raise RuntimeError(f"检查 Excel 进程 {pid} 的 Wind 日志归属时超时。") from error
    if result.returncode:
        raise RuntimeError(f"无法检查 Excel 进程 {pid} 的 Wind 日志归属。")
    log_dir = _WIND_LOG_DIR.resolve()
    return {
        log_dir / path.name for line in result.stdout.splitlines()
        if line.startswith("n/") and re.fullmatch(r"excel_\d{8}_\d{6}\.log", (path := Path(line[1:])).name)
        if path.parent.samefile(log_dir)
    }


def _owned_wind_logs(pid: int, instances: list[dict]) -> set[Path]:
    owned = _wind_logs_for_pid(pid)
    if not owned or any(
        owned & _wind_logs_for_pid(app["pid"])
        for app in instances if app["pid"] != pid
    ):
        raise RuntimeError("无法把 Wind 响应日志独立归属到本次后台 Excel。")
    return owned


def _wind_log_offsets(paths: set[Path]) -> dict[Path, tuple[int, int]]:
    offsets = {}
    for path in paths:
        stat = path.stat()
        offsets[path] = (stat.st_ino, stat.st_size)
    return offsets


def _wind_response_count(offsets: dict[Path, tuple[int, int]]) -> int:
    count = 0
    for path, (inode, offset) in offsets.items():
        try:
            with path.open("rb") as log:
                stat = os.fstat(log.fileno())
                if stat.st_ino != inode or stat.st_size < offset:
                    continue
                log.seek(offset)
                for line in log:
                    if not line.endswith(b"\n"):
                        break
                    line = line.lower()
                    if b"post_matrix_request" in line and b"response" in line and b"request:" in line:
                        count += 1
        except FileNotFoundError:
            continue
    return count


def _wait_for_wind_wave(offsets: dict[Path, tuple[int, int]], previous: int, deadline: float) -> int:
    """Wait for a quiet window after fresh responses from the worker."""
    last_count = previous
    changed_at = time.monotonic()
    while (now := time.monotonic()) < deadline:
        count = _wind_response_count(offsets)
        if count != last_count:
            changed_at, last_count = now, count
        if count > previous and now - changed_at >= 0.8:
            return count
        time.sleep(min(0.2, deadline - now))
    raise RuntimeError("等待本次后台 Excel 的新 Wind 响应超时。")


def _remember_worker(marker: Path, record: dict[str, Any]) -> None:
    with tempfile.NamedTemporaryFile(dir=marker.parent, prefix=marker.name, delete=False) as stream:
        temporary = Path(stream.name)
        try:
            stream.write(json.dumps(record, ensure_ascii=False).encode())
            stream.flush()
            os.fsync(stream.fileno())
            temporary.replace(marker)
        finally:
            temporary.unlink(missing_ok=True)


def _new_excel_worker(before: list[dict], deadline: float) -> dict:
    while True:
        candidates = [app for app in _run_automation("snapshot") if app not in before]
        if len(candidates) == 1:
            return candidates[0]
        if candidates:
            raise RuntimeError("启动期间出现多个 Excel 实例，无法判定本次任务进程。")
        if time.monotonic() >= deadline:
            raise RuntimeError("未观察到独立的后台 Excel 进程。")
        time.sleep(0.2)


def refresh_xlsx(
    path: Path, resolved_params: list[tuple[ParameterDef, Any]], timeout_sec: int
) -> None:
    """Recalculate a staged copy in a hidden, PID-addressed Excel instance."""
    if sys.platform != "darwin" or path.suffix.casefold() != ".xlsx":
        raise RuntimeError("此刷新入口需要 macOS 版 Microsoft Excel 和 .xlsx 副本。")
    path = path.resolve(strict=True)
    try:
        sheet_names, cells = _wind_cells(path)
        params, targets = _parameter_args(resolved_params, sheet_names)
    except (ET.ParseError, KeyError, zipfile.BadZipFile) as error:
        raise RuntimeError("无法读取 Excel 副本及其公式外链。") from error

    watches = cells.keys() - targets
    wind_workbook = bool(watches)
    if wind_workbook and not _WIND_ADDIN.is_file():
        raise RuntimeError("未找到 macOS 版 Wind Excel 加载项。")

    before = _run_automation("snapshot")
    marker = _worker_marker()
    record = {"path": str(path), "before": before, "worker": None}
    _remember_worker(marker, record)
    started_at = time.monotonic()
    deadline = started_at + max(5, int(timeout_sec))
    worker = None
    owned = False
    data = {"path": str(path), "wind": wind_workbook, "watches": sorted(watches), "params": params}

    def stage(mode: str) -> dict:
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            raise RuntimeError("Excel 刷新超过时间预算。")
        try:
            return _run_automation(mode, {**data, **worker, "seconds": remaining})
        except (RuntimeError, subprocess.TimeoutExpired) as error:
            raise RuntimeError(f"Excel 后台阶段 {mode} 失败：{error}") from error

    try:
        launch = subprocess.run(
            ["/usr/bin/open", "-n", "-g", "-j", "-b", "com.microsoft.Excel", str(path)],
            capture_output=True, text=True, timeout=min(10, max(5, int(timeout_sec))),
        )
        if launch.returncode:
            raise RuntimeError(f"无法在 Excel 打开计算副本：{launch.stderr.strip()}")
        worker = _new_excel_worker(before, deadline)
        record["worker"] = worker
        _remember_worker(marker, record)
        stage("open")
        owned = True
        opened_at = time.monotonic()
        prepared = stage("prepare")
        data["calculation"] = prepared["calculation"]
        instances = _run_automation("snapshot")
        offsets = _wind_log_offsets(_owned_wind_logs(worker["pid"], instances)) if wind_workbook else {}

        refresh = stage("refresh")
        wind_wait_ms = 0
        settle_ms = 0
        response_count = 0
        while True:
            if wind_workbook:
                wait_started = time.monotonic()
                response_count = _wait_for_wind_wave(offsets, response_count, deadline)
                wind_wait_ms += (time.monotonic() - wait_started) * 1000
                logs = _owned_wind_logs(worker["pid"], _run_automation("snapshot"))
                if not set(offsets) <= logs:
                    raise RuntimeError("无法确认本次后台 Excel 的 Wind 响应归属。")
            settled = stage("settle")
            settle_ms += settled["settle_ms"]
            if settled["status"] == "saved":
                break
            if settled["status"] != "pending" or not wind_workbook:
                raise RuntimeError("Excel 未确认保存计算副本。")

        if wind_workbook:
            _, saved = _wind_cells(path)
            if any(address not in saved or saved[address][0] == "e" or not saved[address][1] for address in watches):
                raise RuntimeError("保存后的 Wind 公式缺少有效缓存结果。")
            logs = _owned_wind_logs(worker["pid"], _run_automation("snapshot"))
            if not set(offsets) <= logs or _wind_response_count(offsets) < response_count:
                raise RuntimeError("未观察到本次后台 Excel 的新 Wind 请求响应。")
        logger.info(
            "Excel worker pid=%d open=%.2fs calculate=%dms wind_wait=%.0fms settle=%dms save=%dms",
            worker["pid"], opened_at - started_at, refresh["calculate_ms"], wind_wait_ms,
            settle_ms, settled["save_ms"],
        )
    except subprocess.TimeoutExpired as error:
        raise RuntimeError("Excel 自动化超时，请检查 Excel 或 Wind 的登录与授权提示。") from error
    finally:
        original = sys.exc_info()[1]
        try:
            if worker is not None:
                cleanup_error = None
                if worker in _run_automation("snapshot"):
                    try:
                        _run_automation("cleanup", {**data, **worker, "seconds": 5, "owned": owned})
                    except (RuntimeError, subprocess.TimeoutExpired) as error:
                        cleanup_error = error
                quit_by = time.monotonic() + 4
                while worker in _run_automation("snapshot"):
                    if time.monotonic() >= quit_by:
                        raise RuntimeError("任务 Excel 进程仍在运行。") from cleanup_error
                    time.sleep(0.2)
            else:
                raise RuntimeError("无法确认 LaunchServices 是否已完成启动请求。")
            marker.unlink()
        except (OSError, ValueError, RuntimeError, subprocess.TimeoutExpired) as error:
            detail = f" 原始错误：{original}" if original is not None else ""
            raise WorkerCleanupError(
                f"无法确认后台 Excel 已退出。恢复记录：{marker}。{error}{detail}"
            ) from error
