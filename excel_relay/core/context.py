"""excel-relay 上下文与项目空间感知工具 (Context & Project Utilities).

提供纯函数工具以支持 Interface 层与第三方 SDK 用户快速感知当前环境与项目空间：
1. 向上寻道探测目标路径所归属的 excel-relay 项目空间根目录 (find_project_root)；
2. 校验指定目录是否为已初始化的项目根目录 (is_project_dir) 或敏感公共目录 (is_sensitive_dir)；
3. 为跨级检索/查询场景提供标准级联加载器 (get_cascading_catalog)。
"""

from __future__ import annotations

from pathlib import Path
from typing import Optional

from excel_relay.core.catalog import CascadingCatalog, CatalogManager


def find_project_root(start_path: Optional[Path | str] = None) -> Optional[Path]:
    """从起始路径自底向上寻道查找 excel-relay 项目根目录.

    寻道规则：
    1. 逐级向上检查当前层级或任一父级是否存在 `.excel-relay/` 目录 (且排除用户主目录 ~/.excel-relay)；
    2. 若到达用户主目录或盘符根目录仍未找到，返回 None (绝不抛出异常)。
    """
    if start_path is not None:
        p = Path(start_path).resolve()
    else:
        p = Path.cwd().resolve()

    if p.is_file():
        p = p.parent

    home_dir = Path.home().resolve()
    current = p

    while True:
        if current == home_dir:
            break

        # 增加防御性 OSError 捕获，避免跨越受系统权限保护的特殊目录时意外中断
        try:
            if (current / ".excel-relay").is_dir():
                return current
        except OSError:
            pass

        parent = current.parent
        if parent == current:
            # 已到达驱动器根目录 (如 C:\ 或 D:\)
            break
        current = parent

    return None


def is_project_dir(path: Path | str) -> bool:
    """便捷判断指定路径自身是否是一个已初始化的 excel-relay 项目根目录.

    规则：
    1. 必须包含已存在的 `.excel-relay` 目录；
    2. 必须严格排除用户主目录 (Path.home())，因为 ~/.excel-relay 属于用户全局资产仓，绝非独立项目空间。
    """
    if not path:
        return False
    p = Path(path).resolve()
    # 文件或用户主目录 (全局注册表宿主) 绝不能被误判为项目目录
    if p.is_file() or p == Path.home().resolve():
        return False
    try:
        return (p / ".excel-relay").is_dir()
    except OSError:
        return False


def is_sensitive_dir(path: Optional[Path | str] = None) -> bool:
    """检查路径是否为敏感/全局性公共目录 (如盘符根目录、用户主目录、桌面、下载等)."""
    if not path:
        return True
    p = Path(path).resolve()
    if p.is_file():
        p = p.parent

    # 1. 盘符根目录 (如 C:\ 或 D:\)
    if p.parent == p:
        return True

    home = Path.home().resolve()
    # 2. 用户主目录 (如 C:\Users\<用户名>)
    if p == home:
        return True

    # 3. 常见全局性公共子目录 (支持中英文，以及 Windows OneDrive 重定向路径)
    sensitive_names = {
        "desktop", "桌面",
        "downloads", "下载",
        "documents", "文档",
        "pictures", "图片",
        "music", "音乐",
        "videos", "视频",
        "appdata",
        "temp", "tmp",
    }
    is_parent_home_or_root = (p.parent == home or p.parent == p)
    is_onedrive_redirected = (
        p.parent.parent == home and "onedrive" in p.parent.name.lower()
    )
    if p.name.lower() in sensitive_names and (is_parent_home_or_root or is_onedrive_redirected):
        return True

    return False


def get_cascading_catalog(workspace_or_cwd: Optional[Path | str] = None) -> CascadingCatalog:
    """构建标准级联数据源读总管 (供 Interface 层与 SDK 查询/检索场景按需调用).

    - 若指定路径 (或缺省当前终端 cwd) 处于某个 excel-relay 项目内部，组装 CascadingCatalog(primary=项目仓, fallback=全局仓)；
    - 否则组装 CascadingCatalog(primary=全局仓, fallback=None)。
    """
    root = find_project_root(workspace_or_cwd)
    global_mgr = CatalogManager()
    if root is not None:
        return CascadingCatalog(primary=CatalogManager(base_dir=root), fallback=global_mgr)
    return CascadingCatalog(primary=global_mgr, fallback=None)
