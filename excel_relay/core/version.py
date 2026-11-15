"""excel-relay 版本信息定义与获取接口 (Core Domain Version).

提供 excel-relay 核心引擎版本号、Catalog 仓储协议格式版本及系统运行环境结构化信息。
纯业务域与环境探测，零外部协议依赖。
"""

from __future__ import annotations

import platform
import sys
from typing import Any, Dict

__version__ = "0.1.0"
CATALOG_PROTOCOL_VERSION = "0.1.0"


def get_version() -> str:
    """获取 excel-relay 核心引擎当前版本号（PEP 440 格式）."""
    return __version__


def get_version_info() -> Dict[str, Any]:
    """获取完整的版本及运行环境结构化元数据字典.

    Returns:
        包含 version, catalog_version, python_version, platform, system 等环境信息的字典。
    """
    return {
        "version": __version__,
        "catalog_version": CATALOG_PROTOCOL_VERSION,
        "python_version": f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}",
        "platform": sys.platform,
        "system": platform.platform(terse=True),
    }
