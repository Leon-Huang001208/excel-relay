"""macOS Excel automation for a staged workbook copy."""

from __future__ import annotations

import math
import os
import plistlib
import posixpath
import re
import subprocess
import sys
import tempfile
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

_OPEN_SCRIPT = r"""
on run argv
    set targetPath to item 1 of argv
    set workbookName to item 2 of argv
    set secondsLimit to (item 3 of argv) as integer
    set windWorkbook to item 4 of argv is "wind"
    set deadline to (current date) + secondsLimit
    with timeout of secondsLimit seconds
        tell application id "com.microsoft.Excel"
            repeat until exists workbook workbookName
                if (current date) is greater than or equal to deadline then error "Excel 未打开计算副本。"
                delay 0.2
            end repeat
            if full name of workbook workbookName is not targetPath then error "Excel 打开的路径与计算副本不一致。"
            if windWorkbook then
                if (count of workbooks) is not 1 then error "Wind 刷新会清理 Excel 会话的共享缓存。请先关闭其他工作簿。"
                set installedNames to name of every add in whose installed is true
                if "windfunc.xlam" is not in installedNames then error "Excel 未启用 Wind 加载项。"
            end if
            -- Finish requests issued by Excel while opening the saved workbook.
            repeat
                set openingPending to false
                repeat with watchIndex from 5 to (count of argv) by 2
                    set watchSheet to worksheet (item watchIndex of argv) of workbook workbookName
                    set watchCell to range (item (watchIndex + 1) of argv) of watchSheet
                    set openingText to string value of watchCell
                    ignoring case
                        if openingText contains "fetch" or openingText contains "请求中" or openingText contains "获取中" then set openingPending to true
                    end ignoring
                end repeat
                if not openingPending then exit repeat
                if (current date) is greater than or equal to deadline then error "等待打开时触发的 Wind 请求结束超时。"
                delay 0.5
            end repeat
            return "opened"
        end tell
    end timeout
end run
"""

_REFRESH_SCRIPT = r"""
on run argv
    set targetPath to item 1 of argv
    set workbookName to item 2 of argv
    set secondsLimit to (item 3 of argv) as integer
    set windWorkbook to item 4 of argv is "wind"
    set paramCount to (item 5 of argv) as integer
    set deadline to (current date) + secondsLimit
    set checkedCopy to missing value
    with timeout of secondsLimit seconds
        tell application id "com.microsoft.Excel"
            try
                set openedBook to workbook workbookName
                if full name of openedBook is not targetPath then error "Excel 打开的路径与计算副本不一致。"
                set checkedCopy to openedBook
                if read only of checkedCopy then error "Excel 以只读方式打开计算副本。"
                if windWorkbook and (count of workbooks) is not 1 then error "Wind 刷新会清理当前 Excel 会话的共享缓存。请先关闭其他工作簿。"

                set argIndex to 6
                repeat paramCount times
                    set targetSheet to worksheet (item argIndex of argv) of checkedCopy
                    set targetCell to range (item (argIndex + 1) of argv) of targetSheet
                    set scalarType to item (argIndex + 2) of argv
                    set rawValue to item (argIndex + 3) of argv
                    if scalarType is "number" then
                        set value of targetCell to rawValue as real
                    else if scalarType is "boolean" then
                        set value of targetCell to (rawValue is "true")
                    else if scalarType is "empty" then
                        clear contents targetCell
                    else
                        set value of targetCell to ("'" & rawValue)
                    end if
                    set argIndex to argIndex + 4
                end repeat

                if windWorkbook then
                    run VB macro "'windfunc.xlam'!windexcel_ui.clear_cache_data"
                end if
                set targetSheetNames to name of every worksheet of checkedCopy
                repeat with sheetName in targetSheetNames
                    set targetSheet to worksheet (contents of sheetName) of checkedCopy
                    if (contents of sheetName) is not "_excel_relay" then
                        set wasEnabled to enable calculation of targetSheet
                        set enable calculation of targetSheet to false
                        set enable calculation of targetSheet to true
                        calculate targetSheet
                        if not wasEnabled then set enable calculation of targetSheet to false
                    end if
                end repeat

                set stableReadings to 0
                set previousValues to missing value
                repeat while (current date) is less than deadline
                    set foundPending to false
                    set currentValues to {}
                    repeat with watchIndex from argIndex to (count of argv) by 2
                        set watchSheet to worksheet (item watchIndex of argv) of checkedCopy
                        set watchCell to range (item (watchIndex + 1) of argv) of watchSheet
                        set rawValue to value of watchCell
                        if rawValue is missing value then
                            set foundPending to true
                        else
                            set valueText to rawValue as text
                            ignoring case
                                if valueText is "" or valueText starts with "#" or valueText contains "fetch" or valueText contains "请求中" or valueText contains "获取中" then set foundPending to true
                            end ignoring
                            set end of currentValues to valueText
                        end if
                    end repeat
                    if not foundPending and currentValues = previousValues then
                        set stableReadings to stableReadings + 1
                        if stableReadings is 2 then exit repeat
                    else
                        set stableReadings to 0
                    end if
                    set previousValues to currentValues
                    delay 1
                end repeat
                if stableReadings is not 2 then error "Wind 结果存在错误、空值或异步取数超时。"

                save checkedCopy
                if not (saved of checkedCopy) then error "Excel 未确认保存计算副本。"
                close checkedCopy saving no
                return "saved"
            on error errText number errNumber
                if checkedCopy is not missing value then
                    try
                        close checkedCopy saving no
                    end try
                end if
                error errText number errNumber
            end try
        end tell
    end timeout
end run
"""

_CLOSE_SCRIPT = """
on run argv
    if application id "com.microsoft.Excel" is not running then return "closed"
    with timeout of 4 seconds
        tell application id "com.microsoft.Excel"
            set workbookName to item 2 of argv
            if exists workbook workbookName then
                set candidate to workbook workbookName
                if full name of candidate is not (item 1 of argv) then error "Excel 副本路径不一致。"
                close candidate saving no
            end if
            return "closed"
        end tell
    end timeout
end run
"""


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


@contextmanager
def excel_refresh_lock() -> Iterator[None]:
    """Serialize access to Excel and the process-wide Wind cache across invocations."""
    import fcntl

    lock_path = Path(tempfile.gettempdir()) / f"excel-relay-excel-{os.getuid()}.lock"
    fd = os.open(lock_path, os.O_CREAT | os.O_WRONLY | os.O_NOFOLLOW, 0o600)
    with os.fdopen(fd, "wb") as lock:
        try:
            fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError as error:
            raise RuntimeError("Excel Relay 正在刷新另一份工作簿，请在完成后重试。") from error
        yield


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
) -> tuple[list[str], set[tuple[str, str]]]:
    by_name = {name.casefold(): name for name in sheet_names}
    data_sheets = [name for name in sheet_names if name.casefold() != "_excel_relay"]
    args: list[str] = []
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
        args.extend((actual_sheet, cell, scalar_type, serialized))
        targets.add((actual_sheet, cell))
    return args, targets


def _wind_log_offsets() -> dict[Path, tuple[int, int]]:
    offsets = {}
    for path in _WIND_LOG_DIR.glob("excel_*.log"):
        try:
            stat = path.stat()
            offsets[path] = (stat.st_ino, stat.st_size)
        except FileNotFoundError:
            continue
    return offsets


def _new_wind_response(offsets: dict[Path, tuple[int, int]]) -> bool:
    for path in _WIND_LOG_DIR.glob("excel_*.log"):
        try:
            with path.open("rb") as log:
                stat = os.fstat(log.fileno())
                inode, offset = offsets.get(path, (0, 0))
                if stat.st_ino == inode and offset <= stat.st_size:
                    log.seek(offset)
                for line in log:
                    line = line.lower()
                    if b"post_matrix_request" in line and b"response" in line and b"request:" in line:
                        return True
        except FileNotFoundError:
            continue
    return False


def refresh_xlsx(
    path: Path, resolved_params: list[tuple[ParameterDef, Any]], timeout_sec: int
) -> None:
    """Recalculate the staged file and require a new Wind response for Wind formulas."""
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
    seconds = max(5, int(timeout_sec))
    open_args = [str(path), path.name, str(seconds), "wind" if wind_workbook else "offline"]
    watch_args = [part for target in sorted(watches) for part in target]

    def run_script(script: str, args: list[str], wait: int, expected: str) -> None:
        result = subprocess.run(
            ["/usr/bin/osascript", "-", *args],
            input=script,
            capture_output=True,
            text=True,
            check=False,
            timeout=wait,
        )
        if result.returncode or result.stdout.strip() != expected:
            raise RuntimeError(f"macOS Excel 刷新失败：{result.stderr.strip() or 'Excel 未确认操作'}")

    try:
        launch = subprocess.run(
            ["/usr/bin/open", "-g", "-b", "com.microsoft.Excel", str(path)],
            capture_output=True,
            text=True,
            check=False,
            timeout=10,
        )
        if launch.returncode:
            raise RuntimeError(f"无法在 Excel 打开计算副本：{launch.stderr.strip()}")
        run_script(_OPEN_SCRIPT, [*open_args, *watch_args], seconds + 5, "opened")
        offsets = _wind_log_offsets() if wind_workbook else {}
        script_args = [*open_args, str(len(resolved_params)), *params, *watch_args]
        run_script(_REFRESH_SCRIPT, script_args, seconds + 8, "saved")
        if wind_workbook:
            _, saved = _wind_cells(path)
            if any(
                address not in saved or saved[address][0] == "e" or not saved[address][1]
                for address in watches
            ):
                raise RuntimeError("保存后的 Wind 公式缺少有效缓存结果。")
            if not _new_wind_response(offsets):
                raise RuntimeError("未观察到新的 Wind 请求响应，无法确认取数已刷新。")
    except (OSError, ValueError, RuntimeError, subprocess.TimeoutExpired) as error:
        try:
            run_script(_CLOSE_SCRIPT, [str(path), path.name], 6, "closed")
        except (OSError, RuntimeError, subprocess.TimeoutExpired) as close_error:
            raise RuntimeError(f"{error} 此外无法确认 Excel 副本已关闭：{close_error}") from error
        if isinstance(error, subprocess.TimeoutExpired):
            raise RuntimeError("Excel 自动化超时，请检查 Excel 或 Wind 的登录与授权提示。") from error
        raise RuntimeError(str(error)) from error
