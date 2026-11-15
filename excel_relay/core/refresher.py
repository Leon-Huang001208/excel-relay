"""Excel 刷新引擎与计算产物管理。"""

from __future__ import annotations

import datetime as dt
import logging
import os
import shutil
import sys
import tempfile
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from excel_relay.core.catalog import CascadingCatalog, CatalogManager
from excel_relay.core.context import get_cascading_catalog
from excel_relay.core.inspector import (
    ExcelInspector,
    build_canonical_params,
    coerce_numeric_scalar,
    normalize_cell_ref,
)
from excel_relay.core.models import (
    ParameterDef,
    RunArtifact,
    SourceEntry,
    compute_artifact_id,
    compute_params_hash,
)

logger = logging.getLogger("excel-relay.refresher")


class RefreshError(Exception):
    """Excel 刷新与参数校验异常."""
    pass


class ExcelComSession:
    """Windows Win32 COM Excel 独立进程会话上下文管理器.

    提供端到端的 COM 生命周期管理：
    1. 进程安全拉起与隐藏模式预设 (Visible=False, DisplayAlerts=False, RTD=200ms)；
    2. 打开/创建工作簿自动跟踪注册，会话退出时批量无痕关闭；
    3. 进程 PID 实时嗅探与活动跟踪；
    4. 退出时执行强制垃圾回收 (gc.collect())，彻底回收 COM 接口包装器与 IUnknown 引用，
       杜绝 pythoncom.CoUninitialize() 触发 RPC_E_DISCONNECTED (0x80010108) 崩溃；
    5. 独立线程 Apartment 安全注销与孤儿进程强杀兜底。
    """

    def __init__(
        self,
        visible: bool = False,
        display_alerts: bool = False,
        screen_updating: bool = False,
        ask_to_update_links: bool = False,
        rtd_throttle_interval: Optional[int] = 200,
    ):
        self.visible = visible
        self.display_alerts = display_alerts
        self.screen_updating = screen_updating
        self.ask_to_update_links = ask_to_update_links
        self.rtd_throttle_interval = rtd_throttle_interval
        self.app = None
        self.pid: Optional[int] = None
        self._workbooks: List[Any] = []
        self._co_initialized: bool = False

    def __enter__(self) -> "ExcelComSession":
        if sys.platform != "win32":
            raise RuntimeError("非 Windows 平台，不支持 Win32 COM")

        import pythoncom
        import win32com.client as win32

        pythoncom.CoInitialize()
        self._co_initialized = True
        try:
            self.app = win32.DispatchEx("Excel.Application")
            self.pid = HeadlessExcelRefresher._get_excel_pid(self.app)
            self.app.Visible = self.visible
            self.app.DisplayAlerts = self.display_alerts
            self.app.ScreenUpdating = self.screen_updating
            self.app.AskToUpdateLinks = self.ask_to_update_links
            if self.rtd_throttle_interval is not None:
                try:
                    self.app.RTD.ThrottleInterval = self.rtd_throttle_interval
                except Exception as e:
                    logger.debug("设置 RTD.ThrottleInterval 异常: %s", e)
            return self
        except Exception:
            self._cleanup()
            raise

    def open_workbook(
        self,
        file_path: str | Path,
        update_links: int = 0,
        read_only: bool = False,
    ) -> Any:
        """安全打开工作簿并登记到会话跟踪列表，退出时统一安全关闭."""
        if not self.app:
            raise RuntimeError("Excel COM 会话尚未初始化或已关闭")
        if read_only:
            wb = self.app.Workbooks.Open(str(file_path), UpdateLinks=update_links, ReadOnly=True)
        else:
            wb = self.app.Workbooks.Open(str(file_path), UpdateLinks=update_links)
        self.track_workbook(wb)
        return wb

    def track_workbook(self, wb: Any) -> Any:
        """登记外部创建或获取的工作簿对象，以便退出时统一清理."""
        if wb is not None and wb not in self._workbooks:
            self._workbooks.append(wb)
        return wb

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        self._cleanup()

    def _cleanup(self) -> None:
        import gc
        import faulthandler

        # 1. 安全关闭并释放所有登记的工作簿
        for wb in self._workbooks:
            try:
                wb.Close(SaveChanges=False)
            except Exception as e:
                logger.debug("关闭工作簿异常: %s", e)
        self._workbooks.clear()

        # 2. 退出 Excel 应用程序
        pid = self.pid
        if self.app is not None:
            try:
                self.app.Quit()
            except Exception as e:
                logger.debug("Excel app.Quit() 异常: %s", e)

        # 3. 释放已退出的 Excel COM 接口与垃圾回收期间，暂时屏蔽 faulthandler，
        #    避免 Windows 向量化异常处理器 (VEH) 捕获良性的 RPC_E_DISCONNECTED (0x80010108)
        #    或 RPC_S_CALL_FAILED (0x800706be) 并向 stderr 输出虚假的 fatal exception 堆栈
        was_fh_enabled = False
        try:
            was_fh_enabled = faulthandler.is_enabled()
            if was_fh_enabled:
                faulthandler.disable()
        except Exception:
            pass

        try:
            self.app = None
            gc.collect()
            try:
                import pythoncom
                pythoncom.PumpWaitingMessages()
            except Exception:
                pass
        finally:
            if was_fh_enabled:
                try:
                    faulthandler.enable()
                except Exception:
                    pass

        # 4. 优先等待 Excel 进程自然退出 (若超时则强杀)，确保不会留下孤儿后台进程
        if pid:
            HeadlessExcelRefresher._ensure_process_terminated(pid)
            self.pid = None


class HeadlessExcelRefresher:
    """Excel 后台无声刷新与沙箱重算器."""

    KNOWN_ADDINS = [
        "WDF.Addin",              # Wind (万得)
        "ThsXlsSlug.ExcelAddin",  # iFinD (同花顺)
    ]

    KNOWN_VENDORS = {
        "WDF.Addin": "Wind (万得)",
        "ThsXlsSlug.ExcelAddin": "iFinD (同花顺)",
    }

    PENDING_KEYWORDS = [
        "fetching",
        "retrieving",
        "loading",
        "waiting",
        "requesting",
        "connecting",
        "#getting_data",
        "#wait!",
        "正在检索",
        "正在查询",
        "正在获取",
        "正在请求",
        "正在加载",
        "正在连接",
        "正在登录",
        "等待数据",
        "数据请求中",
        "排队中",
        "请稍候",
    ]

    @staticmethod
    def _get_excel_pid(app) -> Optional[int]:
        """通过窗口句柄精确获取 Excel 独立进程的操作系统 PID."""
        try:
            import ctypes

            hwnd = getattr(app, "Hwnd", None)
            if hwnd:
                pid = ctypes.c_ulong()
                ctypes.windll.user32.GetWindowThreadProcessId(int(hwnd), ctypes.byref(pid))
                if pid.value > 0:
                    return int(pid.value)
        except Exception as e:
            logger.debug("获取 Excel PID 异常: %s", e)
        return None

    @staticmethod
    def _is_process_alive(pid: int) -> bool:
        """检查指定 PID 进程当前是否仍在活动运行."""
        if not pid or pid <= 0:
            return False
        try:
            import ctypes

            handle = ctypes.windll.kernel32.OpenProcess(0x1000, False, pid)
            if not handle:
                return False
            exit_code = ctypes.c_ulong()
            ctypes.windll.kernel32.GetExitCodeProcess(handle, ctypes.byref(exit_code))
            ctypes.windll.kernel32.CloseHandle(handle)
            return exit_code.value == 259  # STILL_ACTIVE
        except Exception as e:
            logger.debug("检查进程存活状态异常 (PID=%s): %s", pid, e)
            return False

    @classmethod
    def _ensure_process_terminated(cls, pid: Optional[int], wait_sec: float = 2.0) -> None:
        """强力防孤儿进程兜底：若 Excel 在退出后仍处于活动状态，精准强制终结该独立进程."""
        if not pid or pid <= 0:
            return

        # 轮询等待 Excel 进程正常自然退出，避免过早强杀触发 RPC_S_CALL_FAILED (0x800706be)
        if wait_sec > 0:
            deadline = time.time() + wait_sec
            while time.time() < deadline:
                if not cls._is_process_alive(pid):
                    return
                time.sleep(0.05)

        if not cls._is_process_alive(pid):
            return

        import os
        import signal

        try:
            os.kill(pid, signal.SIGTERM)
        except Exception as e_kill:
            logger.debug("os.kill (PID=%s) 失败: %s", pid, e_kill)
            try:
                import subprocess

                subprocess.run(
                    ["taskkill", "/F", "/PID", str(pid)],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    check=False,
                )
            except Exception as e_taskkill:
                logger.debug("taskkill (PID=%s) 失败: %s", pid, e_taskkill)

    @classmethod
    def probe_environment(cls) -> Dict[str, Any]:
        """探测本地 Excel 自动化环境与金融插件。"""
        if sys.platform == "darwin":
            from excel_relay.core.macos_excel import probe_environment

            return probe_environment()

        info: Dict[str, Any] = {
            "python_version": f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}",
            "platform": sys.platform,
            "excel_ok": False,
            "excel_version": None,
            "excel_error": None,
            "plugins": [],
        }

        if sys.platform != "win32":
            info["excel_error"] = "非 Windows 平台，不支持 Win32 COM"
            return info

        plugins_found: List[str] = []
        try:
            with ExcelComSession() as session:
                app = session.app
                info["excel_ok"] = True
                info["excel_version"] = str(app.Version)
                plugins_found.extend(cls._probe_excel_plugins(app))
                del app
        except Exception as e:
            info["excel_error"] = str(e)
            logger.debug("探测本地 Excel COM 环境异常: %s", e)

        if not plugins_found:
            plugins_found.extend(cls._probe_registry_plugins())

        info["plugins"] = list(dict.fromkeys(plugins_found))
        return info

    @classmethod
    def _probe_excel_plugins(cls, xl) -> List[str]:
        """从活跃 Excel 实例中探测已连接的 COM 加载项."""
        plugins: List[str] = []
        try:
            com_addins = xl.COMAddIns
            count = int(com_addins.Count)
            for i in range(1, count + 1):
                addin = com_addins(i)
                prog_id = str(getattr(addin, "ProgId", "")).lower()
                is_connected = bool(getattr(addin, "Connect", False))
                del addin
                for known_id, vendor_name in cls.KNOWN_VENDORS.items():
                    if prog_id == known_id.lower():
                        status_text = "[就绪]" if is_connected else "[未连接]"
                        plugins.append(f"{vendor_name} {status_text}")
            del com_addins
        except Exception as e:
            logger.debug("读取 Excel COMAddIns 异常: %s", e)
        return plugins

    @classmethod
    def _probe_registry_plugins(cls) -> List[str]:
        """从 Windows 注册表嗅探已安装的金融插件 (万得/同花顺)."""
        import winreg

        plugins: List[str] = []
        for root in [winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE]:
            for key_path in [
                r"Software\Microsoft\Office\Excel\Addins",
                r"Software\WOW6432Node\Microsoft\Office\Excel\Addins",
            ]:
                try:
                    with winreg.OpenKey(root, key_path) as k:
                        count = winreg.QueryInfoKey(k)[0]
                        for idx in range(count):
                            try:
                                sub = winreg.EnumKey(k, idx)
                                if "WDF" in sub or "Wind" in sub:
                                    plugins.append("Wind (万得) [注册表]")
                                elif "THS" in sub or "iFinD" in sub or "ThsXlsSlug" in sub:
                                    plugins.append("iFinD (同花顺) [注册表]")
                            except OSError:
                                break
                except Exception as e:
                    logger.debug("探测注册表加载项异常 (%s): %s", key_path, e)
        return plugins

    @classmethod
    def validate_and_resolve_params(
        cls,
        params: Dict[str, Any],
        allowed_inputs: Optional[List[ParameterDef]],
    ) -> List[Tuple[ParameterDef, Any]]:
        """严格校验传入参数是否在数据源已登记的 inputs 白名单中.

        规则:
        1. 若数据源未登记任何 inputs (allowed_inputs 为空)，严禁传参修改单元格；
        2. 传入的每个参数 key 必须匹配某个已登记 ParameterDef 的 name 或 cell；
        3. 返回解析后的 [(ParameterDef, target_value), ...] 供 COM 精准写入。
        """
        if not params:
            return []

        if not allowed_inputs:
            raise RefreshError(
                "安全拦截: 该数据源未在 Catalog 中登记任何可调参数 (inputs 为空)，"
                "属于无参模型 f(t)，仅支持无参刷新，不允许传入参数修改单元格。"
                "严禁私自调用工具篡改参数白名单；若需配置可调参数，必须先向用户确认具体的参数单元格坐标。"
            )

        resolved_list: List[Tuple[ParameterDef, Any]] = []
        seen_param_cells: set[str] = set()
        for raw_key, val in params.items():
            key_clean = str(raw_key).strip()
            key_lower = key_clean.lower()
            key_norm_cell = normalize_cell_ref(key_clean).lower()

            matched_def: Optional[ParameterDef] = None
            for inp in allowed_inputs:
                if inp.name.strip().lower() == key_lower or normalize_cell_ref(inp.cell).lower() == key_norm_cell:
                    matched_def = inp
                    break

            if not matched_def:
                valid_desc = ", ".join([f"{i.name}({i.cell})" for i in allowed_inputs])
                raise RefreshError(
                    f"安全拦截: 参数 '{raw_key}' 未在该数据源的可调参数白名单中登记！"
                    f"当前允许的参数: [{valid_desc}]。"
                    "严禁私自修改参数白名单；请向用户报告当前允许的参数范围，或由用户明确指定后再配置新参数。"
                )

            cell_id = normalize_cell_ref(matched_def.cell).lower()
            if cell_id in seen_param_cells:
                raise RefreshError(
                    f"参数冲突: 参数 [{matched_def.name} ({matched_def.cell})] 在单次调用中被重复指定！"
                )
            seen_param_cells.add(cell_id)

            # 数值类型安全提升：若传入的是纯数字字符串 (如 "15", "0.05")，自动提升为 int 或 float，
            # 避免以文本格式写入 Excel 单元格导致 SUM、AVERAGE 等公式计算被忽略或比较运算失效
            coerced_val = coerce_numeric_scalar(val)
            resolved_list.append((matched_def, coerced_val))

        return resolved_list

    @classmethod
    def build_full_params(
        cls,
        params: Optional[Dict[str, Any]],
        allowed_inputs: Optional[List[ParameterDef]],
    ) -> Optional[Dict[str, Any]]:
        """构建带参模型 f(t, X) 的全量规范化参数配置字典 X."""
        return build_canonical_params(allowed_inputs, params=params, strict=False)

    @classmethod
    def refresh(
        cls,
        source_id_or_path: str | Path,
        params: Optional[Dict[str, Any]] = None,
        output_path: Optional[str | Path] = None,
        timeout_sec: int = 120,
        keep_formulas: bool = False,
        catalog_mgr: Optional[CatalogManager | CascadingCatalog] = None,
    ) -> Tuple[str, float, Optional[RunArtifact]]:
        """刷新已注册的 Excel 数据源，并更新目录或记录计算产物。

        macOS 使用桌面 Excel 刷新 Wind .xlsx 底稿，每次生成保留公式的结果副本。
        Windows 分流与持久化契约：
        1. 强制从 catalog.json 解析已注册的 SourceEntry (禁止未注册文件刷新，禁止外部伪造 inputs)；
        2. 无参模型 f(t) (entry.profile.inputs 为空)：
           - 默认原地刷新保存母版，并在 core 层直接调用 mgr.record_in_place_refresh() 更新母版刷新时间戳；
           - 仅当显式指定 output_path 时另存为副本并调用 mgr.record_artifact() 记录产物。
        3. 带参模型 f(t, X) (entry.profile.inputs 非空)：
           - 无论是否传入 params (未传即取默认基准参数向量 X_0)，一律另存为沙箱副本，母版享有绝对只读保护；
           - 沙箱副本文件名按 artifact_id ({source_hash6}_{params_hash6}_{stem}) 确定性命名并原地覆盖更新最新时间 t 的结果；
           - 在 core 层直接调用 mgr.record_artifact() 完成 catalog.json 的唯一覆盖更新 (Upsert)。

        Returns:
            Tuple[落盘文件绝对路径, 耗时秒数, RunArtifact对象(若原地刷新则为None)]
        """
        if sys.platform == "win32":
            # 与 Windows 原流程一致，在创建副本前验证 COM 依赖。
            import pythoncom  # noqa: F401
            import win32com.client  # noqa: F401
        elif sys.platform != "darwin":
            raise RefreshError("Excel 刷新支持 Windows 和 macOS（Excel + Wind .xlsx）。")

        if catalog_mgr is not None:
            mgr = catalog_mgr
        else:
            probe_path = None
            if isinstance(source_id_or_path, (str, Path)):
                try:
                    p = Path(source_id_or_path)
                    if p.exists():
                        probe_path = p
                except Exception:
                    pass
            mgr = get_cascading_catalog(workspace_or_cwd=probe_path)

        entry = mgr.get_source(source_id_or_path)
        if entry is None:
            raise RefreshError(
                f"未在 Catalog 中找到对应数据源: '{source_id_or_path}'。"
                "所有刷新与改参操作均严格基于 Catalog 契约管理，请先注册数据源。"
            )

        src = mgr.resolve_absolute_path(entry)
        if not src.is_file():
            raise FileNotFoundError(f"源文件不存在: {src}")
        vendors = {vendor.strip() for vendor in entry.physical.vendor.split(",")}
        if sys.platform == "darwin" and (src.suffix.lower() != ".xlsx" or "ifind" in vendors):
            raise RefreshError("macOS 刷新支持 Excel 与 Wind 加载项使用的 .xlsx 底稿。")

        allowed_inputs = entry.inputs

        # 核心防御：若存在可调参数，在计算全量参数向量与哈希前，先离线从物理文件单元格中读取绝对最新的基准数值，
        # 彻底杜绝分析师在外部修改母版单元格后 catalog.json 缓存滞后导致的“账面记录与实际计算脱节”
        if allowed_inputs:
            ExcelInspector.refresh_input_values(src, allowed_inputs)

        # 0. 在拉起 Excel 进程前优先执行严格参数白名单校验
        resolved_params: List[Tuple[ParameterDef, Any]] = []
        if params:
            resolved_params = cls.validate_and_resolve_params(params, allowed_inputs)

        if sys.platform == "darwin":
            return cls._refresh_macos(mgr, entry, src, params, resolved_params, output_path, timeout_sec)

        # 决定运行模式：带参模型 f(t, X) 或显式指定 output_path 时一律走沙箱副本
        has_declared_inputs = bool(allowed_inputs)
        is_sandbox = has_declared_inputs or bool(resolved_params) or (output_path is not None)

        backup_file: Optional[Path] = None
        # 引入进程 PID 隔离备份文件名，彻底避免并发刷新时备份冲突碰撞
        pid_tag = os.getpid()
        if output_path:
            dst = Path(output_path).resolve()
            dst.parent.mkdir(parents=True, exist_ok=True)
            if dst.is_file():
                backup_file = dst.with_name(f".~excel_relay_bak_{pid_tag}_{dst.name}")
                shutil.copy2(dst, backup_file)
            if dst != src:
                shutil.copy2(src, dst)
            work_file = dst
        elif is_sandbox:
            artifacts_dir = mgr.get_artifacts_dir(entry.source_id)
            full_params = cls.build_full_params(params, allowed_inputs)
            params_hash = compute_params_hash(full_params)
            artifact_id = compute_artifact_id(entry.source_id, params_hash)
            deterministic_file = artifacts_dir / f"{artifact_id}_{src.stem}{src.suffix}"
            try:
                if deterministic_file.is_file():
                    backup_file = deterministic_file.with_name(f".~excel_relay_bak_{pid_tag}_{deterministic_file.name}")
                    shutil.copy2(deterministic_file, backup_file)
                shutil.copy2(src, deterministic_file)
                work_file = deterministic_file
            except PermissionError:
                if backup_file and backup_file.is_file():
                    backup_file.unlink(missing_ok=True)
                backup_file = None
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                work_file = artifacts_dir / f"{artifact_id}_{src.stem}_{timestamp}{src.suffix}"
                shutil.copy2(src, work_file)
        else:
            work_file = src
            backup_file = src.with_name(f".~excel_relay_bak_{pid_tag}_{src.name}")
            shutil.copy2(src, backup_file)

        start_time = time.time()
        elapsed: float = 0.0
        refresh_succeeded = False

        try:
            with ExcelComSession(
                visible=False,
                display_alerts=False,
                screen_updating=False,
                ask_to_update_links=False,
                rtd_throttle_interval=200,
            ) as session:
                app = session.app

                # 1. 尝试激活已知的数据插件
                for addin_name in cls.KNOWN_ADDINS:
                    try:
                        addin = app.COMAddIns(addin_name)
                        if not addin.Connect:
                            addin.Connect = True
                        del addin
                    except Exception as e:
                        logger.debug("激活加载项 %s 异常: %s", addin_name, e)

                # 2. 打开工作簿前预检：清除可能由 Git、模板或受限目录继承的只读文件属性，确保重算后正常存盘
                try:
                    import stat

                    file_mode = os.stat(work_file).st_mode
                    if not (file_mode & stat.S_IWRITE):
                        os.chmod(work_file, file_mode | stat.S_IWRITE)
                except Exception as e:
                    logger.debug("清除文件只读属性异常 (%s): %s", work_file, e)

                wb = session.open_workbook(str(work_file), update_links=0, read_only=False)

                # 3. 若有校验通过的参数，精准写入目标单元格
                if resolved_params:
                    cls._apply_validated_parameters(wb, resolved_params)

                # 4. 触发重算 (优先使用深度重构重算，重建 XLL/插件依赖树)
                try:
                    app.CalculateFullRebuild()
                except Exception:
                    app.CalculateFull()

                # 5. 等待异步数据返回与全表数值收敛 (若超时未完成将直接抛出 RefreshError 拒绝存盘)
                cls._wait_calculation(app, wb, timeout_sec)

                # 6. 若生成了副本且未要求保留公式，固化为纯数值（彻底解决插件脱机 #NAME? 陷阱）
                if is_sandbox and not keep_formulas:
                    cls._strip_formulas(wb)

                # 7. 存盘退出
                wb.Save()
                refresh_succeeded = True
                elapsed = round(time.time() - start_time, 2)

        except RefreshError:
            raise
        except Exception as e:
            raise RefreshError(f"后台刷新 Excel 失败: {e}") from e
        finally:
            # 事务性文件保护：无论原地刷新母版还是沙箱副本，成功则清理备份，超时或异常则回滚至刷新前状态或清理新增半成品
            if refresh_succeeded:
                if backup_file and backup_file.is_file():
                    try:
                        backup_file.unlink(missing_ok=True)
                    except Exception as e:
                        logger.debug("清理备份文件失败 (%s): %s", backup_file, e)
            else:
                try:
                    if backup_file and backup_file.is_file():
                        # 在 Windows 上短暂重试原子回滚，避免 Excel COM 退出瞬间文件锁延迟释放导致 shutil.move 报错
                        replaced = False
                        for _attempt in range(5):
                            try:
                                os.replace(backup_file, work_file)
                                replaced = True
                                break
                            except OSError:
                                time.sleep(0.05)
                        if not replaced:
                            shutil.copy2(backup_file, work_file)
                            backup_file.unlink(missing_ok=True)
                    elif work_file != src and work_file.is_file():
                        work_file.unlink(missing_ok=True)
                except Exception as e:
                    logger.debug("事务回滚文件异常 (%s): %s", work_file, e)

        # 9. 在 core 层内聚完成 catalog.json 状态闭环更新 (在 Excel COM 释放之后安全落盘)
        artifact: Optional[RunArtifact] = None
        if is_sandbox:
            artifact = mgr.record_artifact(
                source_id_or_path=entry.source_id,
                artifact_path=work_file,
                params=params,
                elapsed_seconds=elapsed,
                keep_formulas=keep_formulas,
                refresh_inputs=False,
            )
        else:
            mgr.record_in_place_refresh(
                source_id_or_path=entry.source_id,
            )

        return str(work_file), elapsed, artifact

    @classmethod
    def _refresh_macos(
        cls,
        mgr: CatalogManager | CascadingCatalog,
        entry: SourceEntry,
        src: Path,
        params: Optional[Dict[str, Any]],
        resolved_params: List[Tuple[ParameterDef, Any]],
        output_path: Optional[str | Path],
        timeout_sec: int,
    ) -> Tuple[str, float, RunArtifact]:
        from excel_relay.core.macos_excel import excel_refresh_lock, refresh_xlsx

        try:
            with excel_refresh_lock():
                if output_path is None:
                    full_params = cls.build_full_params(params, entry.inputs)
                    artifact_id = compute_artifact_id(entry.source_id, compute_params_hash(full_params))
                    dst = mgr.get_artifacts_dir(entry.source_id) / f"{artifact_id}_{src.stem}.xlsx"
                else:
                    dst = Path(output_path).expanduser()

                dst = dst.resolve()
                if dst == src or (dst.is_file() and dst.samefile(src)):
                    raise RefreshError("macOS 刷新需要单独的输出文件，不能覆盖母版。")
                if dst.suffix.lower() != ".xlsx" or (dst.exists() and not dst.is_file()):
                    raise RefreshError("请输入有效的 .xlsx 输出路径。")
                if mgr.get_source(dst) is not None:
                    raise RefreshError("macOS 刷新不能覆盖已注册底稿，请选择单独的输出路径。")

                dst.parent.mkdir(parents=True, exist_ok=True)
                with tempfile.NamedTemporaryFile(prefix="excel-relay-", suffix=".xlsx", dir=dst.parent, delete=False) as stage:
                    stage_path = Path(stage.name)
                backup = stage_path.with_suffix(".previous")
                try:
                    shutil.copyfile(src, stage_path)
                    start_time = time.monotonic()
                    refresh_xlsx(stage_path, resolved_params, timeout_sec)
                    elapsed = round(time.monotonic() - start_time, 2)

                    try:
                        if dst.is_file():
                            shutil.copy2(dst, backup)
                        os.replace(stage_path, dst)
                    except OSError:
                        backup.unlink(missing_ok=True)
                        raise
                    try:
                        artifact = mgr.record_artifact(
                            source_id_or_path=entry.source_id,
                            artifact_path=dst,
                            params=params,
                            elapsed_seconds=elapsed,
                            keep_formulas=True,
                            refresh_inputs=False,
                        )
                        if artifact is None:
                            raise RefreshError("数据源已从目录移除，无法登记刷新结果。")
                    except Exception:
                        if backup.is_file():
                            try:
                                os.replace(backup, dst)
                            except OSError as exc:
                                raise RefreshError(f"登记失败，上一份结果保存在 {backup}：{exc}") from exc
                        else:
                            dst.unlink(missing_ok=True)
                        raise
                    else:
                        backup.unlink(missing_ok=True)
                finally:
                    stage_path.unlink(missing_ok=True)
        except (OSError, RuntimeError) as exc:
            raise RefreshError(f"macOS Excel 刷新失败: {exc}") from exc

        return str(dst), elapsed, artifact

    @staticmethod
    def _strip_formulas(wb) -> None:
        """将工作簿所有工作表的公式批量固化为纯数值 (完好保留 _excel_relay 元数据规范表)."""
        for ws in wb.Worksheets:
            try:
                if str(ws.Name).strip().lower() == "_excel_relay":
                    continue
                used = ws.UsedRange
                if used is not None:
                    used.Value = used.Value
            except Exception as e:
                logger.debug("脱水公式异常 (Sheet=%s): %s", getattr(ws, "Name", ""), e)

    @classmethod
    def _locate_com_range(cls, wb, cell_ref: str):
        """在 COM Workbook 中根据 'Sheet!A1' 单元格坐标定位 Range 对象 (不支持命名区域)."""
        clean = normalize_cell_ref(cell_ref)
        if "!" in clean:
            s_name, c_addr = clean.rsplit("!", 1)
            # 大小写与首尾空格容错比对：分析师制作底稿常有工作表名尾随空格（如 "假设表 "），
            # normalize_cell_ref 裁剪为 "假设表" 后若直接调 wb.Worksheets("假设表") 会抛出 COM 异常。
            target_sh_lower = s_name.strip().lower()
            for i in range(1, wb.Worksheets.Count + 1):
                ws = wb.Worksheets(i)
                if str(ws.Name).strip().lower() == target_sh_lower:
                    return ws.Range(c_addr)
            return wb.Worksheets(s_name).Range(c_addr)

        # 默认在第一个非 _excel_relay 的业务工作表中查找
        for i in range(1, wb.Worksheets.Count + 1):
            ws = wb.Worksheets(i)
            if str(ws.Name).strip().lower() != "_excel_relay":
                return ws.Range(clean)
        return wb.ActiveSheet.Range(clean)

    @classmethod
    def _apply_validated_parameters(
        cls,
        wb,
        resolved_params: List[Tuple[ParameterDef, Any]],
    ) -> None:
        """将已通过白名单校验的参数精准写入对应的 ParameterDef.cell."""
        for param_def, val in resolved_params:
            try:
                rng = cls._locate_com_range(wb, param_def.cell)
            except Exception as e:
                raise RefreshError(
                    f"无法在工作簿中定位已登记的参数 [{param_def.name}] (位置: {param_def.cell}): {e}"
                ) from e
            if rng is None:
                raise RefreshError(
                    f"无法在工作簿中定位已登记的参数 [{param_def.name}] (位置: {param_def.cell})"
                )
            rng.Value = val

    @classmethod
    def _inspect_workbook_state(cls, wb) -> Tuple[int, str, List[str]]:
        """扫描工作簿所有工作表的当前值状态，返回 (未就绪单元格数, 全表数值快照Hash, 未就绪示例列表)."""
        import hashlib

        hasher = hashlib.blake2b(digest_size=16)
        pending_count = 0
        pending_samples: List[str] = []

        for ws in wb.Worksheets:
            try:
                ws_name = str(ws.Name)
                if ws_name.strip().lower() == "_excel_relay":
                    continue
                hasher.update(ws_name.encode("utf-8", errors="ignore"))
                used_vals = ws.UsedRange.Value
                if used_vals is None:
                    hasher.update(b"\x00")
                    continue
                if not isinstance(used_vals, tuple):
                    used_vals = ((used_vals,),)

                sheet_pending_count = 0
                sheet_sample_val = ""
                for row in used_vals:
                    if not isinstance(row, tuple):
                        row = (row,)
                    for c in row:
                        if c is None:
                            hasher.update(b"\x01")
                        elif isinstance(c, str):
                            hasher.update(b"\x02")
                            hasher.update(c.encode("utf-8", errors="ignore"))
                            c_strip = c.strip()
                            if 0 < len(c_strip) <= 40:
                                c_lower = c_strip.lower()
                                if any(k in c_lower for k in cls.PENDING_KEYWORDS):
                                    pending_count += 1
                                    sheet_pending_count += 1
                                    if not sheet_sample_val:
                                        sheet_sample_val = c_strip
                        else:
                            hasher.update(b"\x03")
                            hasher.update(repr(c).encode("utf-8", errors="ignore"))

                if sheet_pending_count > 0 and len(pending_samples) < 3:
                    pending_samples.append(f"{ws_name}({sheet_pending_count}处, 如 '{sheet_sample_val}')")
            except Exception as e:
                # 异常单表隔离保护：对图表工作表 (ChartSheet)、被锁定的只读保护表或 COM 偶发访问异常安全跳过，
                # 避免单张非关键表的读取异常导致整表状态探测崩溃，防止外层收敛计数器被错误清零而引发假超时
                logger.debug("扫描工作表状态异常: %s", e)
                continue

        return pending_count, hasher.hexdigest(), pending_samples

    @classmethod
    def _wait_calculation(
        cls,
        app,
        wb,
        timeout_sec: int = 120,
        poll_interval: float = 0.5,
        required_streak: int = 6,
    ) -> None:
        """基于 RTD 主动拉取、全表数值快照收敛与活性滑动超时的计算完成判定器.

        判定准则 (必须连续 required_streak 次轮询同时满足，默认 6 * 0.5s = 3.0s 稳定窗口):
        1. Excel 原生计算引擎状态就绪 (app.CalculationState == 0)；
        2. 全工作簿无任何处于异步请求中的占位符单元格 (pending_count == 0)；
        3. 全工作簿数值快照哈希与上一轮完全一致 (snapshot_digest == last_digest，即数据停止跳动)。

        活性滑动超时机制 (Activity-based Dynamic Timeout):
        - 基础空闲超时为 timeout_sec；最大硬超时上限为 max(timeout_sec * 3, 300) 秒。
        - 若检测到全表数值哈希仍在变化，或未完成单元格数正在减少，说明插件正在持续回填数据，
          自动重置空闲超时窗口，避免大型工作表被过早截断。
        - 若超时仍未收敛，直接抛出 RefreshError 拒绝存盘，防止污染母版或固化半成品脏数据。
        """
        import pythoncom

        base_timeout = max(int(timeout_sec), 5)
        hard_limit = max(base_timeout * 3, 300)
        start_wait = time.time()
        deadline = start_wait + base_timeout
        hard_deadline = start_wait + hard_limit

        # 初始缓冲窗口，让插件异步请求完成发起登记
        time.sleep(1.0)
        pythoncom.PumpWaitingMessages()

        stable_success_streak = 0
        last_digest: Optional[str] = None
        last_pending_count: Optional[int] = None
        last_pending_samples: List[str] = []
        last_engine_done: bool = False

        while True:
            now = time.time()
            if now >= deadline or now >= hard_deadline:
                break

            pythoncom.PumpWaitingMessages()
            try:
                app.RTD.RefreshData()
            except Exception as e:
                logger.debug("RTD.RefreshData 异常: %s", e)
            pythoncom.PumpWaitingMessages()

            engine_done = False
            try:
                engine_done = (app.CalculationState == 0)
            except Exception:
                engine_done = False

            inspect_ok = False
            pending_count = 0
            snapshot_digest = ""
            pending_samples: List[str] = []
            try:
                pending_count, snapshot_digest, pending_samples = cls._inspect_workbook_state(wb)
                inspect_ok = True
            except Exception:
                inspect_ok = False

            last_engine_done = engine_done
            if inspect_ok:
                # 活性检测：若全表数值仍在更新，或未完成单元格数量正在下降，说明插件正在正常回推数据
                is_making_progress = (
                    (last_digest is not None and snapshot_digest != last_digest)
                    or (last_pending_count is not None and pending_count < last_pending_count)
                )
                if is_making_progress:
                    deadline = min(hard_deadline, max(deadline, time.time() + base_timeout))

                # 严格收敛判定：引擎就绪 + 零占位符 + 全表快照哈希与上一轮完全一致 (停止跳动)
                is_converged_step = (
                    engine_done
                    and pending_count == 0
                    and last_digest is not None
                    and snapshot_digest == last_digest
                )

                last_digest = snapshot_digest
                last_pending_count = pending_count
                last_pending_samples = pending_samples

                if is_converged_step:
                    stable_success_streak += 1
                    if stable_success_streak >= required_streak:
                        return
                else:
                    stable_success_streak = 0
            else:
                stable_success_streak = 0

            time.sleep(poll_interval)

        # 超时未完成：构建精准诊断信息并抛出 RefreshError，严禁静默存盘
        waited = round(time.time() - start_wait, 1)
        reasons: List[str] = []
        if not last_engine_done:
            reasons.append("Excel 公式引擎仍处于计算中 (CalculationState != 0)")
        if last_pending_count:
            sample_str = "、".join(last_pending_samples)
            reasons.append(f"仍有 {last_pending_count} 个单元格处于插件异步拉取状态 [{sample_str}]")
        if not reasons:
            reasons.append("全表单元格数值仍在持续跳动，尚未达到收敛稳定态")

        reason_desc = "；".join(reasons)
        raise RefreshError(
            f"刷新重算超时 (已等待 {waited} 秒): {reason_desc}。"
            "为防止半成品脏数据覆盖母版或固化到产物中，本次刷新已中止并放弃保存。"
            "请检查万得(Wind)/同花顺(iFinD)插件是否已正常登录连接，或调大 timeout 超时时间后重试。"
        )
