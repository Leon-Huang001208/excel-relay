"""数据源注册中心管理器 (Catalog Store Manager).

核心职责：
1. 统一维护 ~/.excel-relay/catalog.json 持久化仓储；
2. 基于文件绝对路径确定性生成 source_id ({hash6}_{stem})，支持 6 位短 Hash 快捷定位；
3. 重复注册时执行无损合并 (保留 ENRICHED 语义与 artifacts 历史)；
4. 严格区分“原地刷新” (仅更新母版时间戳与输出当前值) 与“沙箱副本” (记录 RunArtifact)；
5. 提供失效数据源自动清理 (clean_stale_sources) 与过期沙箱副本清理 (clean_expired_artifacts)；
6. 注销数据源时级联删除其临时沙箱副本，并提供孤儿产物目录兜底清理 (clean_orphan_artifact_dirs)。
"""

from __future__ import annotations

from contextlib import contextmanager
import json
import logging
import os
import re
import shutil
import stat
import sys
import threading
import time
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterator, List, NamedTuple, Optional, Tuple

from excel_relay.core.inspector import ExcelInspector, build_canonical_params
from excel_relay.core.models import (
    CatalogScope,
    CatalogStore,
    DescriptionOrigin,
    InspectionResult,
    MetadataStatus,
    RunArtifact,
    SourceEntry,
    compute_artifact_id,
    compute_params_hash,
    compute_source_id,
)

logger = logging.getLogger("excel-relay.catalog")

# 合法沙箱产物子目录名 (与 compute_source_id 生成的 {proj|glob}_{hash6}_{stem} 结构严格一致)
_ARTIFACT_SID_DIR_PATTERN = re.compile(r"^(proj|glob)_[0-9a-f]{6}_.+$")


class UnregisterResult(NamedTuple):
    """数据源注销结果明细 (含级联清理的临时沙箱副本统计)."""

    source_id: str
    scope: CatalogScope
    removed_artifacts_count: int
    failed_artifact_paths: List[str]


class CatalogManager:
    """数据源目录管理类 (单一资产空间引擎).

    设计不变式 (Invariants):
    1. 唯一核心入参 base_dir:
       - 若 base_dir 为 None 或等于用户主目录 (~)，一律归一化为【用户全局模式】(GLOBAL)；
       - 若 base_dir 为具体工程路径，一律归一化为【项目专属模式】(PROJECT)。
    2. 严格同构空间结构:
       - 无论全局还是项目，资产注册表与沙箱产物永远保存在 base_dir / ".excel-relay" 下；
       - 全局固定为 ~/.excel-relay，项目固定为 <project_root>/.excel-relay。
    3. 纯派生推导:
       - root_dir: 项目模式为 base_dir，全局模式强制为 None；
       - scope: 项目模式为 PROJECT，全局模式为 GLOBAL；
       - project_name: 项目模式优先继承已存名称，次选显式参数，再兜底为 base_dir.name；全局模式为 None。
    """

    def __init__(
        self,
        base_dir: Optional[Path | str] = None,
        project_name: Optional[str] = None,
    ):
        home_base = Path.home().resolve()

        if base_dir is None:
            resolved_base = home_base
        else:
            resolved_base = Path(base_dir).resolve()

        self.base_dir: Path = resolved_base
        self.is_global: bool = (self.base_dir == home_base)

        # 1. 物理目录结构 (严格同构: base_dir / ".excel-relay"，惰性按需创建，构造时零磁盘副作用)
        self.catalog_dir: Path = self.base_dir / ".excel-relay"
        self.catalog_file: Path = self.catalog_dir / "catalog.json"
        self.lock_file: Path = self.catalog_dir / "catalog.lock"
        self.artifacts_dir: Path = self.catalog_dir / "artifacts"
        self._thread_lock = threading.RLock()
        self._lock_depth: int = 0
        self._explicit_project_name: Optional[str] = project_name

        # 2. 边界锚点 (严格单向派生)
        if self.is_global:
            self.root_dir: Optional[Path] = None
        else:
            self.root_dir: Optional[Path] = self.base_dir

        self._last_loaded_sig: Tuple[int, int] = self._get_disk_signature()
        self.store = self._load()

    def _sync_store_scope(self, store: CatalogStore) -> CatalogStore:
        """同步 CatalogStore 的 scope 与 project_name 属性."""
        if self.is_global:
            store.scope = CatalogScope.GLOBAL
            store.project_name = None
        else:
            store.scope = CatalogScope.PROJECT
            if self._explicit_project_name is not None:
                store.project_name = self._explicit_project_name
            elif not store.project_name:
                store.project_name = self.base_dir.name
        return store

    @property
    def scope(self) -> CatalogScope:
        return self.store.scope

    def is_initialized(self) -> bool:
        """判断当前资产空间在磁盘上是否已持久化初始化 (catalog.json 是否存在)."""
        return self.catalog_file.is_file()

    def initialize(self) -> bool:
        """显式在磁盘上初始化当前资产空间目录结构 (.excel-relay/catalog.json 与 artifacts/).

        Returns:
            True 表示本次新建初始化成功；False 表示此前已初始化过 (幂等跳过)。
        """
        already_exists = self.is_initialized()
        self.catalog_dir.mkdir(parents=True, exist_ok=True)
        self.artifacts_dir.mkdir(parents=True, exist_ok=True)
        if not already_exists:
            self.save()
            return True
        return False

    @classmethod
    def init_project(
        cls,
        path: Optional[Path | str] = None,
        project_name: Optional[str] = None,
        allow_sensitive: bool = False,
    ) -> Tuple[CatalogManager, bool]:
        """在指定路径显式初始化一个项目级资产空间 (内聚敏感目录校验与幂等落盘).

        Returns:
            (CatalogManager 实例, created_new: bool 是否为首次新建)
        """
        from excel_relay.core.context import is_sensitive_dir

        target_dir = Path(path).resolve() if path else Path.cwd().resolve()
        if target_dir.is_file():
            target_dir = target_dir.parent

        if not allow_sensitive and is_sensitive_dir(target_dir):
            raise ValueError(
                f"目标路径 '{target_dir}' 为系统公共目录或根目录，不能作为独立项目资产空间。"
            )

        mgr = cls(base_dir=target_dir, project_name=project_name)
        created = mgr.initialize()
        return mgr, created

    def resolve_absolute_path(self, file_path_or_entry: str | Path | Any) -> Path:
        """运行时动态反解：将条目中记录的文件路径转换为实际物理绝对路径."""
        if hasattr(file_path_or_entry, "file_path"):
            file_path_str = file_path_or_entry.file_path
        else:
            file_path_str = file_path_or_entry
        p = Path(file_path_str)
        if p.is_absolute():
            return p.resolve()
        if self.root_dir:
            return (self.root_dir / p).resolve()
        return p.resolve()

    def to_storable_path(self, file_path: Path) -> Tuple[str, bool]:
        """将物理文件绝对路径转换为适合持久化在当前 catalog 中的路径格式."""
        resolved = file_path.resolve()
        # 严格限定：只有项目级且 root_dir 存在时，才允许相对化落盘
        if self.store.scope == CatalogScope.PROJECT and self.root_dir:
            try:
                rel = resolved.relative_to(self.root_dir)
                rel_posix = rel.as_posix()
                return (f"./{rel_posix}", True)
            except ValueError:
                pass
        return (str(resolved), False)

    def _get_disk_signature(self) -> Tuple[int, int]:
        """获取磁盘 catalog.json 的 (纳秒级修改时间, 字节大小) 签名，若不存在则返回 (0, 0)."""
        try:
            if self.catalog_file.is_file():
                st = self.catalog_file.stat()
                return (st.st_mtime_ns, st.st_size)
        except Exception:
            pass
        return (0, 0)

    def _reload_if_changed(self) -> bool:
        """检测磁盘 catalog.json 是否被外部进程修改，若是则自动重新加载最新状态."""
        current_sig = self._get_disk_signature()
        if current_sig != self._last_loaded_sig:
            self.store = self._load()
            self._last_loaded_sig = current_sig
            return True
        return False

    def reload(self) -> CatalogStore:
        """强制从本地磁盘重新加载 catalog.json 并同步内存状态."""
        self.store = self._load()
        self._last_loaded_sig = self._get_disk_signature()
        return self.store

    @contextmanager
    def _exclusive_lock(self, timeout: float = 10.0) -> Iterator[None]:
        """跨线程与跨进程排他文件锁 (基于 .excel-relay/catalog.lock).

        进入最外层锁时自动强制 reload 最新磁盘状态，彻底杜绝多实例并发写入时的更新丢失 (Lost Update)。
        """
        with self._thread_lock:
            if self._lock_depth > 0:
                self._lock_depth += 1
                try:
                    yield
                finally:
                    self._lock_depth -= 1
                return

            self.catalog_dir.mkdir(parents=True, exist_ok=True)
            fd = os.open(str(self.lock_file), os.O_CREAT | os.O_RDWR)
            locked = False
            deadline = time.time() + timeout
            try:
                if os.name == "nt":
                    import msvcrt

                    while True:
                        try:
                            os.lseek(fd, 0, os.SEEK_SET)
                            msvcrt.locking(fd, msvcrt.LK_NBLCK, 1)
                            locked = True
                            break
                        except OSError:
                            if time.time() >= deadline:
                                raise TimeoutError(f"等待 Catalog 排他锁超时: {self.lock_file}")
                            time.sleep(0.02)
                else:
                    import fcntl

                    while True:
                        try:
                            fcntl.flock(fd, fcntl.LOCK_EX | fcntl.LOCK_NB)
                            locked = True
                            break
                        except OSError:
                            if time.time() >= deadline:
                                raise TimeoutError(f"等待 Catalog 排他锁超时: {self.lock_file}")
                            time.sleep(0.02)

                self._lock_depth = 1
                if self.catalog_file.is_file():
                    self.reload()
                yield
            finally:
                self._lock_depth = 0
                if locked:
                    try:
                        if os.name == "nt":
                            import msvcrt

                            os.lseek(fd, 0, os.SEEK_SET)
                            msvcrt.locking(fd, msvcrt.LK_UNLCK, 1)
                        else:
                            import fcntl

                            fcntl.flock(fd, fcntl.LOCK_UN)
                    except Exception:
                        pass
                try:
                    os.close(fd)
                except Exception:
                    pass

    def get_artifacts_dir(self, source_id: Optional[str | Path] = None) -> Path:
        """获取沙箱测算副本存储目录 (~/.excel-relay/artifacts/[{source_id}])."""
        target = self.artifacts_dir / str(source_id) if source_id else self.artifacts_dir
        target.mkdir(parents=True, exist_ok=True)
        return target

    def _load(self) -> CatalogStore:
        """从本地磁盘读取 catalog.json 反序列化为 CatalogStore (若文件损坏则自动备份防丢失)."""
        if not self.catalog_file.exists():
            return self._sync_store_scope(CatalogStore())

        # 若文件存在但为空文件 (0 字节，可能由中断或异常触发创建)，直接初始化新实例，避免无效重试和误生成 .corrupt_* 备份
        try:
            if self.catalog_file.stat().st_size == 0:
                return self._sync_store_scope(CatalogStore())
        except OSError:
            pass

        last_err: Optional[Exception] = None
        for attempt in range(5):
            try:
                with open(self.catalog_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                return self._sync_store_scope(CatalogStore.model_validate(data))
            except PermissionError as pe:
                last_err = pe
                time.sleep(0.02 * (attempt + 1))
            except Exception as e:
                last_err = e
                break

        # 仅在非 PermissionError 引起的解析损坏 (如 JSONDecodeError, ValidationError) 时执行自动备份，
        # 避免因杀毒软件或外部并发读写导致的操作系统暂态文件锁而被误判为物理文件损坏
        if last_err is not None and not isinstance(last_err, PermissionError):
            try:
                if self.catalog_file.is_file() and self.catalog_file.stat().st_size > 0:
                    ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
                    corrupt_bak = self.catalog_file.with_name(f"catalog.json.corrupt_{ts}")
                    shutil.copy2(self.catalog_file, corrupt_bak)
            except Exception:
                pass
        return self._sync_store_scope(CatalogStore())

    def save(self) -> None:
        """通过跨进程锁 + 进程唯一临时文件 + os.replace 原子持久化保存到本地磁盘."""
        with self._exclusive_lock():
            self.catalog_dir.mkdir(parents=True, exist_ok=True)
            self.store.updated_at = datetime.now().isoformat()
            tmp_file = self.catalog_dir / f"catalog.json.{os.getpid()}.{uuid.uuid4().hex[:6]}.tmp"
            try:
                with open(tmp_file, "w", encoding="utf-8") as f:
                    json.dump(self.store.model_dump(), f, ensure_ascii=False, indent=2)
                    f.flush()
                    os.fsync(f.fileno())

                for attempt in range(5):
                    try:
                        os.replace(tmp_file, self.catalog_file)
                        break
                    except PermissionError:
                        if attempt == 4:
                            raise
                        time.sleep(0.02 * (attempt + 1))
            finally:
                if tmp_file.exists():
                    try:
                        tmp_file.unlink(missing_ok=True)
                    except Exception:
                        pass
            self._last_loaded_sig = self._get_disk_signature()

    # -------------------------------------------------------------------------
    # 1. 注册、无损合并与注销 (Register / Merge / Unregister)
    # -------------------------------------------------------------------------
    def register_path(
        self,
        path_str: str | Path,
        custom_tags: Optional[List[str]] = None,
    ) -> List[SourceEntry]:
        """注册单个文件或递归扫描目录下的所有 Excel 文件 (.xlsx/.xlsm/.xls).

        并发优化：先在排他锁外完成耗时的离线文件探测 (inspect_file)，再进入跨进程排他锁执行最新状态重载、无损合并与原子写盘。
        """
        p = Path(path_str).resolve()
        if not p.exists():
            raise FileNotFoundError(f"路径不存在: {p}")

        supported_exts = [".xlsx", ".xlsm", ".xls"]
        ignored_dir_names = {".excel-relay", ".git", ".venv", "venv", "env", "node_modules", "__pycache__"}
        inspected_items: List[Tuple[Path, InspectionResult]] = []

        if p.is_file():
            if any(part in ignored_dir_names for part in p.parts):
                raise ValueError(f"禁止将内部目录 ({p.parent.name}) 下的文件注册为数据源: {p}")
            if p.suffix.lower() in supported_exts:
                self._validate_project_enclosure(p)
                inspected_items.append((p, ExcelInspector.inspect_file(p)))
        elif p.is_dir():
            for item in p.rglob("*"):
                if item.is_file() and item.suffix.lower() in supported_exts:
                    if item.name.startswith(("~$", ".~excel_relay_bak_")):
                        continue
                    if any(part in ignored_dir_names for part in item.parts):
                        continue
                    try:
                        resolved_item = item.resolve()
                        self._validate_project_enclosure(resolved_item)
                        inspected_items.append((resolved_item, ExcelInspector.inspect_file(resolved_item)))
                    except Exception:
                        continue

        results: List[SourceEntry] = []
        with self._exclusive_lock():
            for resolved, inspected in inspected_items:
                entry = self._merge_and_store_inspected(resolved, inspected, custom_tags)
                results.append(entry)
            self.save()
        return results

    def _validate_project_enclosure(self, resolved: Path) -> None:
        """严格项目封闭原则：项目级数据源物理位置必须位于项目根目录内部."""
        if self.store.scope == CatalogScope.PROJECT and self.root_dir:
            try:
                resolved.relative_to(self.root_dir)
            except ValueError:
                raise ValueError(
                    f"项目专属数据源物理位置必须位于项目根目录 ({self.root_dir}) 内部。\n"
                    f"无法将外部文件注册为项目专属数据源: {resolved}\n"
                    f"提示: 若需跨项目复用该底稿，请使用 --global 注册为用户全局数据源。"
                )

    def _register_single_file(
        self,
        file_path: Path,
        custom_tags: Optional[List[str]] = None,
    ) -> SourceEntry:
        """探测并注册单个文件，若已存在则执行无损状态合并."""
        resolved = file_path.resolve()
        self._validate_project_enclosure(resolved)
        inspected = ExcelInspector.inspect_file(resolved)
        with self._exclusive_lock():
            # 补齐原子写盘契约：将探测结果与已有记录做无损合并后持久化保存到磁盘
            entry = self._merge_and_store_inspected(resolved, inspected, custom_tags)
            self.save()
            return entry

    def _merge_and_store_inspected(
        self,
        resolved: Path,
        inspected: InspectionResult,
        custom_tags: Optional[List[str]] = None,
    ) -> SourceEntry:
        """在持有排他锁与最新 store 的前提下，将探测结果与已有记录做无损合并并更新内存字典."""
        # 1. 由 CatalogManager 统一计算持久化路径与唯一 source_id
        storable_path, is_rel = self.to_storable_path(resolved)
        sid = compute_source_id(resolved, root_dir=self.root_dir if is_rel else None)

        physical = inspected.physical
        profile = inspected.profile
        inputs = inspected.inputs
        registered_at = datetime.now().isoformat()
        last_refreshed_at = None
        artifacts: List[RunArtifact] = []

        # 2. 若该文件之前已在 catalog 中注册，按四大组件各自规则执行无损状态合并
        old_entry = self.store.sources.get(sid)
        if old_entry:
            registered_at = old_entry.registered_at
            last_refreshed_at = old_entry.last_refreshed_at
            artifacts = old_entry.artifacts

            # 如果新探测结果的语义状态仅为 RAW (未在 _excel_relay 表中填写【一、模型基本信息】)，且旧条目已经过人工/Agent 赋能 (ENRICHED)
            if (
                profile.metadata_status == MetadataStatus.RAW
                and old_entry.profile.metadata_status == MetadataStatus.ENRICHED
            ):
                extracted_tags = list(profile.tags)
                profile = old_entry.profile.model_copy(deep=True)
                for t in extracted_tags:
                    if t not in profile.tags:
                        profile.tags.append(t)

            # 若源文件未在内部显式声明 inputs，且旧条目已在 Catalog 中配置过 inputs，直接无损保留并刷新当前基准数值
            if not inputs and old_entry.inputs:
                inputs = [inp.model_copy() for inp in old_entry.inputs]
                ExcelInspector.refresh_input_values(resolved, inputs)

        if custom_tags:
            for t in custom_tags:
                if t not in profile.tags:
                    profile.tags.append(t)
            # 用户显式指定自定义标签时，若当前状态非 DECLARED，自愈升级为 ENRICHED 并标记来源为 USER_MANUAL
            if profile.metadata_status != MetadataStatus.DECLARED:
                profile.metadata_status = MetadataStatus.ENRICHED
                profile.description_origin = DescriptionOrigin.USER_MANUAL

        new_entry = SourceEntry(
            source_id=sid,
            file_path=storable_path,
            scope=self.store.scope,
            is_relative=is_rel,
            project_name=self.store.project_name,
            registered_at=registered_at,
            last_refreshed_at=last_refreshed_at,
            physical=physical,
            profile=profile,
            inputs=inputs,
            artifacts=artifacts,
        )
        self.store.sources[sid] = new_entry
        return new_entry

    def unregister(self, source_id_or_path: str | Path) -> bool:
        """根据 source_id、短 Hash 或文件路径注销数据源 (并级联删除其临时沙箱副本)."""
        return self.unregister_detailed(source_id_or_path) is not None

    def unregister_detailed(self, source_id_or_path: str | Path) -> Optional[UnregisterResult]:
        """注销数据源并级联删除本仓名下的临时沙箱副本，返回清理明细.

        执行顺序为“先注销、后删除”：catalog 先落盘保证账目正确；若物理删除部分失败
        (如文件被 Excel 占用)，残留目录自然退化为孤儿目录，由 clean_orphan_artifact_dirs 兜底清理。
        仅删除位于本仓 artifacts_dir 内的临时副本，绝不触碰用户指定位置的产物或其它仓储的文件。

        Returns:
            未找到对应数据源时返回 None，否则返回 UnregisterResult。
        """
        self._reload_if_changed()
        if not self.get_source(source_id_or_path):
            return None
        with self._exclusive_lock():
            entry = self.get_source(source_id_or_path)
            if entry and entry.source_id in self.store.sources:
                sid = entry.source_id
                artifacts_snapshot = list(entry.artifacts)
                del self.store.sources[sid]
                self.save()
                removed_count, failed_paths = self._purge_source_artifacts(sid, artifacts_snapshot)
                return UnregisterResult(
                    source_id=sid,
                    scope=self.store.scope,
                    removed_artifacts_count=removed_count,
                    failed_artifact_paths=failed_paths,
                )
        return None

    # -------------------------------------------------------------------------
    # 沙箱产物安全删除工具 (仅作用于本仓 artifacts_dir，防越界误删)
    # -------------------------------------------------------------------------
    @staticmethod
    def _is_link_or_junction(p: Path) -> bool:
        """判断路径是否为符号链接或 Windows Junction / 其它重解析点 (rmtree 顶层不防护此类目录)."""
        try:
            if os.path.islink(p):
                return True
            st = os.lstat(p)
        except OSError:
            return False
        reparse_flag = getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0x400)
        attrs = getattr(st, "st_file_attributes", 0)
        return bool(attrs & reparse_flag)

    def _is_within_artifacts_root(self, p: Path) -> bool:
        """判断物理路径是否严格位于本仓 artifacts_dir 之内 (基于路径组件比较，而非字符串前缀)."""
        try:
            root = self.artifacts_dir.resolve()
            target = Path(p).resolve()
        except Exception:
            return False
        return target != root and target.is_relative_to(root)

    def _is_safe_artifact_dir(self, sid_dir: Path) -> bool:
        """判断目录是否允许整体删除：必须是本仓 artifacts_dir 的直接子目录、名称符合 source_id 格式、且非链接."""
        if not _ARTIFACT_SID_DIR_PATTERN.match(sid_dir.name):
            return False
        if self._is_link_or_junction(sid_dir):
            return False
        try:
            if sid_dir.resolve().parent != self.artifacts_dir.resolve():
                return False
        except Exception:
            return False
        return sid_dir.is_dir()

    def _purge_source_artifacts(
        self,
        sid: str,
        artifacts: List[RunArtifact],
    ) -> Tuple[int, List[str]]:
        """物理删除指定数据源在本仓名下的临时沙箱副本及 artifacts/{sid}/ 目录.

        Returns:
            (已删除的临时副本文件数, 删除失败的路径列表)
        """
        removed_count = 0
        failed_paths: List[str] = []

        # 1. 逐个删除已入账的临时副本 (仅限位于本仓 artifacts_dir 内的文件)
        for art in artifacts:
            if not art.is_temporary:
                continue
            try:
                ap = self.resolve_absolute_path(art.file_path)
            except Exception:
                continue
            if not self._is_within_artifacts_root(ap) or self._is_link_or_junction(ap):
                continue
            if not ap.is_file():
                continue
            try:
                ap.unlink()
                removed_count += 1
            except Exception:
                failed_paths.append(str(ap))

        # 2. 整体删除 artifacts/{sid}/ 目录 (覆盖备份文件、降级生成的带时间戳副本等未入账文件)
        #    注意：不调用 get_artifacts_dir()，避免其 mkdir 副作用
        sid_dir = self.artifacts_dir / sid
        if sid_dir.is_dir() and self._is_safe_artifact_dir(sid_dir):
            leftover = self._rmtree_collect_failures(sid_dir)
            for fp in leftover:
                if fp not in failed_paths:
                    failed_paths.append(fp)

        return removed_count, failed_paths

    @staticmethod
    def _rmtree_collect_failures(target: Path) -> List[str]:
        """删除目录树并收集删除失败的路径 (不抛异常)."""
        failures: List[str] = []

        def _on_error(_func: Any, path: str, _exc: Any) -> None:
            failures.append(str(path))

        try:
            if sys.version_info >= (3, 12):
                shutil.rmtree(target, onexc=_on_error)
            else:
                shutil.rmtree(target, onerror=_on_error)
        except Exception:
            failures.append(str(target))
        return failures

    def list_sources(
        self,
        tag_filter: Optional[str] = None,
        scope_filter: Optional[CatalogScope] = None,
    ) -> List[SourceEntry]:
        """列出所有已注册数据源."""
        self._reload_if_changed()
        if scope_filter is not None and self.store.scope != scope_filter:
            return []
        items = list(self.store.sources.values())
        if tag_filter:
            tf = tag_filter.lower()
            items = [item for item in items if any(tf in t.lower() for t in item.profile.tags)]
        return items

    def get_source(self, source_id_or_path: str | Path) -> Optional[SourceEntry]:
        """根据完整 source_id、短 Hash 或文件物理路径查找数据源."""
        self._reload_if_changed()
        key = str(source_id_or_path).strip() if source_id_or_path is not None else ""
        if not key:
            return None

        # 1. 精确匹配完整 source_id
        if key in self.store.sources:
            return self.store.sources[key]

        # 2. 尝试按文件物理绝对路径匹配
        try:
            # 无论调用方传入的是相对路径还是绝对路径，均锚定到 self.root_dir 进行绝对路径解析，与当前进程工作目录 (cwd) 解耦
            target_path = str(self.resolve_absolute_path(key)).lower()
            for entry in self.store.sources.values():
                abs_entry = str(self.resolve_absolute_path(entry.file_path)).lower()
                if abs_entry == target_path:
                    return entry
        except Exception:
            pass

        # 3. 支持带作用域前缀 (proj_a1f9c2 / glob_a1f9c2) 或纯 6 位短 Hash (a1f9c2) 前缀匹配
        # 由于 proj_ 与 glob_ 严格同构 (均为 5 字符前缀)，sid[5:] 恒为 {hash6}_{stem}
        key_lower = key.lower()
        prefix_matches = [
            entry
            for sid, entry in self.store.sources.items()
            if sid.lower().startswith(key_lower) or sid.lower()[5:].startswith(key_lower)
        ]
        if len(prefix_matches) == 1:
            return prefix_matches[0]

        return None

    # -------------------------------------------------------------------------
    # 2. 多维检索与语义更新 (Search & Enrich)
    # -------------------------------------------------------------------------
    def search(
        self,
        query: str,
        top_k: int = 5,
        tag_filter: Optional[str] = None,
        scope_filter: Optional[CatalogScope] = None,
    ) -> List[Tuple[SourceEntry, float, str]]:
        """基于关键词或意图打分搜索最匹配的数据源.

        综合匹配: ID、名称、摘要、标签、推荐问法、可调参数名、工作表横向表头与纵向指标名。
        """
        self._reload_if_changed()
        if scope_filter is not None and self.store.scope != scope_filter:
            return []
        candidates = self.list_sources(tag_filter=tag_filter, scope_filter=scope_filter)
        q = query.lower().strip()
        if not q:
            return [(s, 1.0, "默认排序") for s in candidates[:top_k]]

        # 正则多重标点切分：覆盖中英文逗号、顿号、分号、冒号、斜杠、感叹号与中英文括号，大幅提高金融复合词意图召回率
        keywords = [k for k in re.split(r"[\s,，、;；?？!！。.:：\(\)（）\[\]【】/]+", q) if k]
        results: List[Tuple[SourceEntry, float, str]] = []

        for entry in candidates:
            score = 0.0
            reasons = []

            # 1. ID / 名称强匹配
            if q in entry.source_id.lower() or any(k in entry.source_id.lower() for k in keywords):
                score += 4.0
                reasons.append("ID匹配")
            if q in entry.profile.name.lower() or any(k in entry.profile.name.lower() for k in keywords):
                score += 5.0
                reasons.append("名称匹配")

            # 2. 摘要匹配
            if q in entry.profile.summary.lower() or any(k in entry.profile.summary.lower() for k in keywords):
                score += 3.0
                reasons.append("摘要匹配")

            # 3. 标签匹配
            matched_tags = [t for t in entry.profile.tags if any(k in t.lower() or t.lower() in k for k in keywords)]
            if matched_tags:
                score += len(matched_tags) * 2.0
                reasons.append(f"标签命中: {','.join(matched_tags)}")

            # 4. 推荐问法匹配
            for qst in entry.profile.questions:
                if any(k in qst.lower() for k in keywords):
                    score += 2.5
                    reasons.append(f"问法命中: {qst}")
                    break

            # 5. 可调输入参数匹配
            matched_in = []
            for inp in entry.inputs:
                if any(k in inp.name.lower() or (inp.description and k in inp.description.lower()) for k in keywords):
                    matched_in.append(inp.name)
            if matched_in:
                score += 2.5
                reasons.append(f"参数命中: {','.join(matched_in[:3])}")

            # 6. 工作表名、横向列字段 (columns) 与纵向指标名 (row_headers) 匹配
            for s in entry.physical.sheets:
                if any(k in s.name.lower() for k in keywords):
                    score += 1.5
                    reasons.append(f"工作表名: {s.name}")
                matched_cols = [c for c in s.columns if any(k in c.lower() for k in keywords)]
                if matched_cols:
                    score += 1.2
                    reasons.append(f"横向字段: {','.join(matched_cols[:3])}")
                matched_rows = [r for r in s.row_headers if any(k in r.lower() for k in keywords)]
                if matched_rows:
                    score += 1.2
                    reasons.append(f"纵向指标: {','.join(matched_rows[:3])}")

            if score > 0:
                results.append((entry, score, "; ".join(reasons)))

        results.sort(key=lambda x: x[1], reverse=True)
        return results[:top_k]

    def update_source(
        self,
        source_id: str | Path,
        name: Optional[str] = None,
        summary: Optional[str] = None,
        tags: Optional[List[str]] = None,
        questions: Optional[List[str]] = None,
        notes: Optional[str] = None,
        origin: DescriptionOrigin = DescriptionOrigin.USER_MANUAL,
    ) -> Optional[SourceEntry]:
        """为已注册的数据源补充或更新业务语义画像 (PackageProfile)."""
        self._reload_if_changed()
        if not self.get_source(source_id):
            return None

        with self._exclusive_lock():
            entry = self.get_source(source_id)
            if not entry:
                return None

            if name is not None:
                entry.profile.name = name
            if summary is not None:
                entry.profile.summary = summary
            if tags is not None:
                entry.profile.tags = tags
            if questions is not None:
                entry.profile.questions = questions
            if notes is not None:
                entry.profile.notes = notes

            entry.profile.description_origin = origin
            if entry.profile.metadata_status != MetadataStatus.DECLARED:
                entry.profile.metadata_status = MetadataStatus.ENRICHED

            self.save()
            return entry

    def set_source_inputs(
        self,
        source_id: str | Path,
        inputs: List[str],
    ) -> Optional[SourceEntry]:
        """为已注册的数据源配置可调输入参数契约 (SourceEntry.inputs)."""
        self._reload_if_changed()
        entry = self.get_source(source_id)
        if not entry:
            return None

        # 在排他锁外完成耗时的离线 Excel 打开与坐标校验
        abs_fp = self.resolve_absolute_path(entry.file_path)
        parsed_in = ExcelInspector.resolve_and_read_inputs(
            abs_fp,
            inputs_raw=inputs,
        )

        with self._exclusive_lock():
            entry = self.get_source(source_id)
            if not entry:
                return None
            if parsed_in is not None:
                entry.inputs = parsed_in
            self.save()
            return entry

    # -------------------------------------------------------------------------
    # 3. 原地刷新记录 vs. 沙箱副本产物记录
    # -------------------------------------------------------------------------
    def _canonicalize_params(
        self,
        entry: SourceEntry,
        params: Optional[Dict[str, Any]],
        refresh_inputs: bool = True,
        strict: bool = True,
    ) -> Optional[Dict[str, Any]]:
        """生成完整、规范化的全量参数配置字典，用于计算唯一 params_hash 并存入 RunArtifact.

        核心逻辑：
        1. 若数据源未登记任何可调参数 (inputs 为空)，且未传入额外参数时返回 None；
        2. 若登记了多个可调参数，先以所有已登记参数的 (inp.name: inp.current_value) 作为基准全集；
        3. 再用本次传入的 params 覆盖对应项 (支持按参数名或单元格坐标匹配)，并对数值做标量归一化 (如 8.0 -> 8)。
        """
        if entry.inputs and refresh_inputs:
            try:
                fp = self.resolve_absolute_path(entry.file_path)
                if fp.is_file():
                    ExcelInspector.refresh_input_values(fp, entry.inputs)
            except Exception:
                pass

        return build_canonical_params(entry.inputs, params=params, strict=strict)

    def record_in_place_refresh(
        self,
        source_id_or_path: str | Path,
    ) -> Optional[SourceEntry]:
        """记录母版原地刷新结果 (不生成 RunArtifact 副本记录).

        同步更新:
        1. last_refreshed_at 时间戳；
        2. 磁盘最新 file_mtime 与 file_size_bytes；
        3. 若登记了 inputs，同步刷新 ParameterDef.current_value。
        """
        self._reload_if_changed()
        pre_entry = self.get_source(source_id_or_path)
        if not pre_entry:
            return None

        # 在获取排他锁前先完成磁盘 stat 与 Excel 单元格预读，最小化锁持有时间
        fp = self.resolve_absolute_path(pre_entry.file_path)
        file_stat: Optional[os.stat_result] = None
        refreshed_input_vals: Dict[str, Any] = {}
        if fp.is_file():
            file_stat = fp.stat()
            if pre_entry.inputs:
                try:
                    ExcelInspector.refresh_input_values(fp, pre_entry.inputs)
                    refreshed_input_vals = {inp.cell: inp.current_value for inp in pre_entry.inputs}
                except Exception:
                    pass

        with self._exclusive_lock():
            entry = self.get_source(source_id_or_path)
            if not entry:
                return None

            entry.last_refreshed_at = datetime.now().isoformat()
            if file_stat is not None:
                entry.physical.file_size_bytes = file_stat.st_size
                entry.physical.file_mtime = str(file_stat.st_mtime)
            if refreshed_input_vals and entry.inputs:
                for inp in entry.inputs:
                    if inp.cell in refreshed_input_vals:
                        inp.current_value = refreshed_input_vals[inp.cell]

            self.save()
            return entry

    def record_artifact(
        self,
        source_id_or_path: str | Path,
        artifact_path: str | Path,
        params: Optional[Dict[str, Any]] = None,
        elapsed_seconds: float = 0.0,
        keep_formulas: bool = False,
        is_temporary: Optional[bool] = None,
        refresh_inputs: bool = True,
    ) -> Optional[RunArtifact]:
        """按 (母版 source_id, 全量参数配置 X) 唯一标识记录或原地更新沙箱副本产物.

        核心机制：
        1. artifact_id = {source_hash6}_{params_hash6}，由母版 ID 与全量参数 X 唯一确定；
        2. 当同一母版在相同参数配置 X 下于新时间 t 再次刷新时，直接覆盖更新该 artifact_id 的记录
           (若此前存在不同路径的旧临时文件则自动清理)，保证每个 (source_id, X) 仅维护最新一份产物。
        """
        self._reload_if_changed()
        pre_entry = self.get_source(source_id_or_path)
        if not pre_entry:
            return None

        # 在排他锁外完成母版最新基准值预读（若调用方已在刷新前读取过，可通过 refresh_inputs=False 跳过二次读取）
        probed_inputs = pre_entry.inputs
        if probed_inputs and refresh_inputs:
            probed_inputs = [inp.model_copy() for inp in probed_inputs]
            try:
                fp = self.resolve_absolute_path(pre_entry.file_path)
                if fp.is_file():
                    ExcelInspector.refresh_input_values(fp, probed_inputs)
            except Exception:
                pass
        latest_input_vals = {inp.cell: inp.current_value for inp in probed_inputs}

        with self._exclusive_lock():
            entry = self.get_source(source_id_or_path)
            if not entry:
                return None

            updated_inputs = [
                inp.model_copy(update={"current_value": latest_input_vals[inp.cell]})
                if inp.cell in latest_input_vals else inp
                for inp in entry.inputs
            ]

            art_p = Path(artifact_path).resolve()
            storable_art_path, _ = self.to_storable_path(art_p)
            canon_params = build_canonical_params(updated_inputs, params=params, strict=False)
            params_hash = compute_params_hash(canon_params)
            artifact_id = compute_artifact_id(entry.source_id, params_hash)

            if is_temporary is None:
                art_str = str(art_p).lower()
                artifacts_root = str(self.artifacts_dir.resolve()).lower()
                is_temp = art_str.startswith(artifacts_root)
            else:
                is_temp = is_temporary

            artifact = RunArtifact(
                artifact_id=artifact_id,
                source_id=entry.source_id,
                file_path=storable_art_path,
                params=canon_params,
                params_hash=params_hash,
                created_at=datetime.now().isoformat(),
                elapsed_seconds=elapsed_seconds,
                is_temporary=is_temp,
                keep_formulas=keep_formulas,
            )

            previous_inputs = entry.inputs
            previous_artifacts = entry.artifacts
            previous_updated_at = self.store.updated_at
            updated_artifacts = [artifact]
            obsolete_artifacts: List[RunArtifact] = []
            for old_art in previous_artifacts:
                if old_art.artifact_id == artifact_id or len(updated_artifacts) >= 30:
                    obsolete_artifacts.append(old_art)
                else:
                    updated_artifacts.append(old_art)

            entry.inputs = updated_inputs
            entry.artifacts = updated_artifacts
            try:
                self.save()
            except Exception:
                entry.inputs = previous_inputs
                entry.artifacts = previous_artifacts
                self.store.updated_at = previous_updated_at
                raise

            # 目录提交后才删除已移除的临时文件，避免磁盘上的目录指向缺失产物。
            for old_art in obsolete_artifacts:
                if not old_art.is_temporary:
                    continue
                try:
                    old_path = self.resolve_absolute_path(old_art.file_path)
                    if old_path != art_p and old_path.is_file():
                        old_path.unlink()
                except Exception as exc:
                    logger.warning("无法清理旧临时产物 %s：%s", old_art.file_path, exc)
            return artifact

    def list_artifacts(self, source_id_or_path: str | Path, limit: int = 10) -> List[RunArtifact]:
        """列出指定数据源的不同参数配置测算产物清单."""
        entry = self.get_source(source_id_or_path)
        if not entry:
            return []
        return entry.artifacts[:limit]

    def find_artifact(
        self,
        source_id_or_path: str | Path,
        params: Optional[Dict[str, Any]] = None,
        artifact_id: Optional[str] = None,
    ) -> Optional[RunArtifact]:
        """按唯一 artifact_id 或参数配置 X 精确查找对应的产物."""
        entry = self.get_source(source_id_or_path)
        if not entry or not entry.artifacts:
            return None

        if artifact_id:
            aid_clean = artifact_id.strip().lower()
            for art in entry.artifacts:
                if art.artifact_id.lower() == aid_clean or art.params_hash.lower() == aid_clean:
                    return art
            return None

        if params is not None:
            try:
                # 纯只读缓存匹配：显式指定 refresh_inputs=False，严禁触发物理 Excel 文件重新读取。
                # 既避免了非必要的磁盘 I/O 耗时，更防止母版被外部编辑后导致基准值发生偏移而误判历史产物不存在。
                canon_params = self._canonicalize_params(entry, params, refresh_inputs=False, strict=True)
            except ValueError:
                return None
            target_hash = compute_params_hash(canon_params)
            target_aid = compute_artifact_id(entry.source_id, target_hash)
            for art in entry.artifacts:
                if art.artifact_id == target_aid or art.params_hash == target_hash:
                    return art
            return None

        return entry.artifacts[0]

    def find_artifact_by_id(self, artifact_id_or_path: str | Path) -> Optional[RunArtifact]:
        """跨所有数据源根据 artifact_id 或文件物理路径查找产物."""
        self._reload_if_changed()
        key = str(artifact_id_or_path).strip() if artifact_id_or_path is not None else ""
        if not key:
            return None

        target_abs_lower: Optional[str] = None
        if "\\" in key or "/" in key or key.lower().endswith((".xlsx", ".xlsm", ".xls")):
            try:
                target_abs_lower = str(self.resolve_absolute_path(key)).lower()
            except Exception:
                pass

        key_lower = key.lower()
        for entry in self.store.sources.values():
            for art in entry.artifacts:
                if art.artifact_id.lower() == key_lower or art.file_path.lower() == key_lower:
                    return art
                if target_abs_lower:
                    try:
                        art_abs_lower = str(self.resolve_absolute_path(art.file_path)).lower()
                        if art_abs_lower == target_abs_lower:
                            return art
                    except Exception:
                        pass
        return None

    # -------------------------------------------------------------------------
    # 4. 清理失效数据源与过期临时副本 (Clean Stale Sources & Expired Artifacts)
    # -------------------------------------------------------------------------
    def clean_stale_sources(
        self,
        scope_filter: Optional[CatalogScope] = None,
    ) -> List[Tuple[str, str, CatalogScope]]:
        """清理因文件被重命名、移动或删除而失效的 source 条目及其临时沙箱副本.

        Returns:
            被移除的 [(source_id, file_path, scope), ...] 清单
        """
        self._reload_if_changed()
        if scope_filter is not None and self.store.scope != scope_filter:
            return []
        if not self.store.sources:
            return []

        removed: List[Tuple[str, str, CatalogScope]] = []
        with self._exclusive_lock():
            stale_ids: List[str] = []
            for sid, entry in self.store.sources.items():
                abs_fp = self.resolve_absolute_path(entry.file_path)
                if not abs_fp.is_file():
                    stale_ids.append(sid)
                    removed.append((sid, entry.file_path, self.store.scope))
                    # 清理该失效源名下的临时沙箱文件与目录 (仅限本仓 artifacts_dir 内，带防越界校验)
                    self._purge_source_artifacts(sid, list(entry.artifacts))

            for sid in stale_ids:
                del self.store.sources[sid]

            if removed:
                self.save()
        return removed

    def clean_orphan_artifact_dirs(
        self,
        scope_filter: Optional[CatalogScope] = None,
    ) -> List[Tuple[str, str, CatalogScope]]:
        """清理本仓 artifacts_dir 下已不对应任何已注册数据源的孤儿产物目录.

        安全边界：
        - 仅扫描本仓 artifacts_dir 的直接子目录，且仅对照本仓 catalog 判定孤儿；
        - 仅处理名称符合 source_id 格式 ({proj|glob}_{hash6}_{stem}) 的目录，其余目录与散落文件一律不碰；
        - 跳过符号链接与 Windows Junction，防止删除链接指向的真实内容。

        Returns:
            被移除的 [(source_id, 目录绝对路径, scope), ...] 清单
        """
        if scope_filter is not None and self.store.scope != scope_filter:
            return []
        # 先行判断目录存在性，避免对未初始化仓储产生加锁建目录的副作用
        if not self.artifacts_dir.is_dir():
            return []

        removed: List[Tuple[str, str, CatalogScope]] = []
        with self._exclusive_lock():
            # 锁内基于最新磁盘状态判定，防止误删其它进程刚注册数据源的产物目录
            self._reload_if_changed()
            registered = set(self.store.sources.keys())
            try:
                children = list(self.artifacts_dir.iterdir())
            except OSError:
                return []
            for child in children:
                if child.name in registered:
                    continue
                if self._is_link_or_junction(child) or not child.is_dir():
                    continue
                if not self._is_safe_artifact_dir(child):
                    continue
                failures = self._rmtree_collect_failures(child)
                if not child.exists():
                    removed.append((child.name, str(child), self.store.scope))
                elif failures:
                    logger.warning("孤儿产物目录未能完全删除 (可能被占用): %s", child)
        return removed

    def clean_expired_artifacts(
        self,
        source_id_or_path: Optional[str | Path] = None,
        max_age_days: int = 7,
        scope_filter: Optional[CatalogScope] = None,
    ) -> int:
        """清理已过期的临时沙箱副本 (从磁盘物理删除并从 catalog 注销)."""
        self._reload_if_changed()
        if scope_filter is not None and self.store.scope != scope_filter:
            return 0
        if not self.store.sources:
            return 0

        cleaned_count = 0
        with self._exclusive_lock():
            entries = (
                [self.get_source(source_id_or_path)]
                if source_id_or_path
                else list(self.store.sources.values())
            )
            now = datetime.now()

            for entry in entries:
                if not entry:
                    continue
                kept_artifacts: List[RunArtifact] = []
                for art in entry.artifacts:
                    f_path = self.resolve_absolute_path(art.file_path)
                    # 1. 物理实体缺失清理：若产物在磁盘上已不存在，一律清理失效死链，确保 Catalog 账实相符
                    if not f_path.exists():
                        cleaned_count += 1
                        continue

                    # 2. 临时沙箱产物过期清理：超过保留天数则物理删除文件并注销
                    if art.is_temporary:
                        try:
                            art_time = datetime.fromisoformat(art.created_at)
                            if (now - art_time).total_seconds() >= max_age_days * 86400:
                                f_path.unlink(missing_ok=True)
                                cleaned_count += 1
                                continue
                        except Exception:
                            pass
                    kept_artifacts.append(art)
                entry.artifacts = kept_artifacts

            if cleaned_count > 0:
                self.save()
        return cleaned_count


class CascadingCatalog:
    """级联数据源管理器 (Cascading / Composite Catalog Store).

    核心职责：
    1. 统一管理 Primary (主仓，如当前项目) 与 Fallback (兜底仓，如用户全局)；
    2. 穿透式查询与路由：先查主仓，后查兜底仓；
    3. 局部优先遮蔽 (Local Shadows Global)：当同一物理文件同时在两级注册时，以主仓为准去重，防结果双重冗余；
    4. 检索权重倾斜 (Project Boost)：在跨级搜索中对主仓命中结果增加优先权重 (+10.0)，确保项目核心模型不被全局公共库冲淡；
    5. 操作路由归位：更新、刷新记录与沙箱产物自动存入所属仓储，维持各级独立性。
    """

    def __init__(
        self,
        primary: CatalogManager,
        fallback: Optional[CatalogManager] = None,
    ):
        self.primary = primary
        self.fallback = fallback

    @property
    def scope(self) -> CatalogScope:
        """返回当前主仓的作用域 (PROJECT 或 GLOBAL，对齐 CatalogManager 多态)."""
        return self.primary.scope

    def is_initialized(self) -> bool:
        """判断主仓在磁盘上是否已持久化初始化 (对齐 CatalogManager 多态)."""
        return self.primary.is_initialized()

    def initialize(self) -> bool:
        """显式初始化主仓的资产空间目录结构 (对齐 CatalogManager 多态)."""
        return self.primary.initialize()

    def save(self) -> None:
        """强制对主仓执行原子持久化保存 (对齐 CatalogManager 多态)."""
        self.primary.save()

    def reload(self) -> Tuple[CatalogStore, Optional[CatalogStore]]:
        """强制从本地磁盘同步重新加载主仓与兜底仓的最新 catalog.json 状态 (对齐 CatalogManager 多态)."""
        primary_store = self.primary.reload()
        fallback_store = self.fallback.reload() if self.fallback else None
        return primary_store, fallback_store

    def to_storable_path(self, file_path: Path | str) -> Tuple[str, bool]:
        """将物理文件绝对路径转换为适合持久化在主仓的路径格式 (对齐 CatalogManager 多态)."""
        p = Path(file_path).resolve() if isinstance(file_path, (str, Path)) else file_path
        return self.primary.to_storable_path(p)

    def resolve_source_manager(self, source_id_or_path: str | Path) -> Optional[CatalogManager]:
        """确定指定数据源归属于哪一个具体的 CatalogManager (优先主仓)."""
        if self.primary.get_source(source_id_or_path):
            return self.primary
        if self.fallback and self.fallback.get_source(source_id_or_path):
            return self.fallback
        return None

    def get_source(self, source_id_or_path: str | Path) -> Optional[SourceEntry]:
        """根据 source_id 或物理路径查找数据源 (主仓优先，次选兜底仓)."""
        entry = self.primary.get_source(source_id_or_path)
        if entry:
            return entry
        if self.fallback:
            return self.fallback.get_source(source_id_or_path)
        return None

    def get_artifacts_dir(self, source_id: Optional[str | Path] = None) -> Path:
        """获取沙箱测算副本存储目录 (路由至归属仓储)."""
        if source_id:
            target_mgr = self.resolve_source_manager(source_id) or self.primary
            return target_mgr.get_artifacts_dir(str(source_id))
        return self.primary.get_artifacts_dir()

    def resolve_absolute_path(self, target: str | Path | SourceEntry | RunArtifact) -> Path:
        """根据条目或产物实际归属仓动态反解绝对物理路径."""
        if isinstance(target, (SourceEntry, RunArtifact)):
            mgr = self.resolve_source_manager(target.source_id) or self.primary
            return mgr.resolve_absolute_path(target.file_path)
        # 若传入的是文件路径字符串
        return self.primary.resolve_absolute_path(target)

    def register_path(
        self,
        path_str: str | Path,
        custom_tags: Optional[List[str]] = None,
    ) -> List[SourceEntry]:
        """向主仓注册单个 Excel 文件或文件夹."""
        return self.primary.register_path(path_str, custom_tags=custom_tags)

    def list_sources(
        self,
        tag_filter: Optional[str] = None,
        scope_filter: Optional[CatalogScope] = None,
    ) -> List[SourceEntry]:
        """列出数据源清单.

        - scope_filter == PROJECT: 仅返回项目级条目；
        - scope_filter == GLOBAL: 仅返回全局级条目；
        - 默认: 返回两仓并集，若同一物理文件重复注册，主仓自动遮蔽兜底仓 (Local Shadows Global)。
        """
        primary_entries = self.primary.list_sources(tag_filter=tag_filter, scope_filter=scope_filter)
        if not self.fallback:
            return primary_entries

        seen_norm_paths = set()
        results: List[SourceEntry] = []

        for e in primary_entries:
            abs_p = str(self.primary.resolve_absolute_path(e.file_path)).lower()
            seen_norm_paths.add(abs_p)
            results.append(e)

        for e in self.fallback.list_sources(tag_filter=tag_filter, scope_filter=scope_filter):
            abs_p = str(self.fallback.resolve_absolute_path(e.file_path)).lower()
            if abs_p not in seen_norm_paths:
                results.append(e)

        return results

    def search(
        self,
        query: str,
        top_k: int = 5,
        tag_filter: Optional[str] = None,
        scope_filter: Optional[CatalogScope] = None,
    ) -> List[Tuple[SourceEntry, float, str]]:
        """跨仓级联检索，并对主仓数据源赋予 Project Boost (+10.0) 权重加成."""
        primary_matches = self.primary.search(
            query, top_k=top_k * 2, tag_filter=tag_filter, scope_filter=scope_filter
        )
        boosted_primary: List[Tuple[SourceEntry, float, str]] = []
        seen_norm_paths = set()

        is_project_primary = (self.primary.store.scope == CatalogScope.PROJECT)
        boost_amount = 10.0 if is_project_primary and scope_filter is None else 0.0

        for entry, score, reason in primary_matches:
            abs_p = str(self.primary.resolve_absolute_path(entry.file_path)).lower()
            seen_norm_paths.add(abs_p)
            boosted_score = score + boost_amount
            boosted_reason = (
                f"{reason} [本项目优先加权]"
                if (is_project_primary and scope_filter is None)
                else reason
            )
            boosted_primary.append((entry, boosted_score, boosted_reason))

        all_matches = list(boosted_primary)

        if self.fallback:
            fallback_matches = self.fallback.search(
                query, top_k=top_k * 2, tag_filter=tag_filter, scope_filter=scope_filter
            )
            for entry, score, reason in fallback_matches:
                abs_p = str(self.fallback.resolve_absolute_path(entry.file_path)).lower()
                if abs_p not in seen_norm_paths:
                    all_matches.append((entry, score, reason))

        all_matches.sort(key=lambda x: x[1], reverse=True)
        return all_matches[:top_k]

    def unregister(self, source_id_or_path: str | Path) -> bool:
        """注销数据源，优先在主仓执行，次选兜底仓 (由命中仓储自行级联清理其临时沙箱副本)."""
        return self.unregister_detailed(source_id_or_path) is not None

    def unregister_detailed(self, source_id_or_path: str | Path) -> Optional[UnregisterResult]:
        """注销数据源并返回清理明细，优先在主仓执行，次选兜底仓 (各仓只清理自己的产物，绝不跨仓)."""
        result = self.primary.unregister_detailed(source_id_or_path)
        if result is not None:
            return result
        if self.fallback:
            return self.fallback.unregister_detailed(source_id_or_path)
        return None

    def update_source(
        self,
        source_id: str | Path,
        name: Optional[str] = None,
        summary: Optional[str] = None,
        tags: Optional[List[str]] = None,
        questions: Optional[List[str]] = None,
        notes: Optional[str] = None,
        origin: DescriptionOrigin = DescriptionOrigin.USER_MANUAL,
    ) -> Optional[SourceEntry]:
        """根据数据源实际所属仓储更新其业务语义画像."""
        mgr = self.resolve_source_manager(source_id)
        if mgr:
            return mgr.update_source(
                source_id,
                name=name,
                summary=summary,
                tags=tags,
                questions=questions,
                notes=notes,
                origin=origin,
            )
        return None

    def set_source_inputs(
        self,
        source_id: str | Path,
        inputs: List[str],
    ) -> Optional[SourceEntry]:
        """根据数据源实际所属仓储配置其可调输入参数契约."""
        mgr = self.resolve_source_manager(source_id)
        if mgr:
            return mgr.set_source_inputs(source_id, inputs=inputs)
        return None

    def record_in_place_refresh(
        self,
        source_id_or_path: str | Path,
    ) -> Optional[SourceEntry]:
        """同步更新母版原地刷新记录至归属仓储."""
        mgr = self.resolve_source_manager(source_id_or_path)
        if mgr:
            return mgr.record_in_place_refresh(source_id_or_path)
        return None

    def record_artifact(
        self,
        source_id_or_path: str | Path,
        artifact_path: str | Path,
        params: Optional[Dict[str, Any]] = None,
        elapsed_seconds: float = 0.0,
        keep_formulas: bool = False,
        is_temporary: Optional[bool] = None,
        refresh_inputs: bool = True,
    ) -> Optional[RunArtifact]:
        """记录沙箱副本产物至归属仓储."""
        mgr = self.resolve_source_manager(source_id_or_path) or self.primary
        return mgr.record_artifact(
            source_id_or_path=source_id_or_path,
            artifact_path=artifact_path,
            params=params,
            elapsed_seconds=elapsed_seconds,
            keep_formulas=keep_formulas,
            is_temporary=is_temporary,
            refresh_inputs=refresh_inputs,
        )

    def list_artifacts(self, source_id_or_path: str | Path, limit: int = 10) -> List[RunArtifact]:
        """列出产物列表."""
        mgr = self.resolve_source_manager(source_id_or_path)
        if mgr:
            return mgr.list_artifacts(source_id_or_path, limit=limit)
        return []

    def find_artifact(
        self,
        source_id_or_path: str | Path,
        params: Optional[Dict[str, Any]] = None,
        artifact_id: Optional[str] = None,
    ) -> Optional[RunArtifact]:
        mgr = self.resolve_source_manager(source_id_or_path)
        if mgr:
            return mgr.find_artifact(source_id_or_path, params=params, artifact_id=artifact_id)
        return None

    def find_artifact_by_id(self, artifact_id_or_path: str | Path) -> Optional[RunArtifact]:
        art = self.primary.find_artifact_by_id(artifact_id_or_path)
        if art:
            return art
        if self.fallback:
            return self.fallback.find_artifact_by_id(artifact_id_or_path)
        return None

    def clean_stale_sources(
        self,
        scope_filter: Optional[CatalogScope] = None,
    ) -> List[Tuple[str, str, CatalogScope]]:
        """按范围清理失效数据源."""
        removed: List[Tuple[str, str, CatalogScope]] = []
        removed.extend(self.primary.clean_stale_sources(scope_filter=scope_filter))
        if self.fallback:
            removed.extend(self.fallback.clean_stale_sources(scope_filter=scope_filter))
        return removed

    def clean_orphan_artifact_dirs(
        self,
        scope_filter: Optional[CatalogScope] = None,
    ) -> List[Tuple[str, str, CatalogScope]]:
        """按范围清理各仓 artifacts 目录下的孤儿产物目录 (各仓仅对照自身 catalog 判定)."""
        removed: List[Tuple[str, str, CatalogScope]] = []
        removed.extend(self.primary.clean_orphan_artifact_dirs(scope_filter=scope_filter))
        if self.fallback:
            removed.extend(self.fallback.clean_orphan_artifact_dirs(scope_filter=scope_filter))
        return removed

    def clean_expired_artifacts(
        self,
        source_id_or_path: Optional[str | Path] = None,
        max_age_days: int = 7,
        scope_filter: Optional[CatalogScope] = None,
    ) -> int:
        """按范围清理已过期的临时沙箱副本."""
        count = 0
        if scope_filter is None or self.primary.scope == scope_filter:
            count += self.primary.clean_expired_artifacts(
                source_id_or_path, max_age_days=max_age_days, scope_filter=scope_filter
            )
        if self.fallback and (scope_filter is None or self.fallback.scope == scope_filter):
            count += self.fallback.clean_expired_artifacts(
                source_id_or_path, max_age_days=max_age_days, scope_filter=scope_filter
            )
        return count
