"""excel-relay: Excel Relay.

将 Excel 模型与插件底稿转化为 AI Agent 可发现、可路由、后台静默刷新的活数据资产。
"""

from excel_relay.core.catalog import CatalogManager
from excel_relay.core.inspector import ExcelInspector
from excel_relay.core.models import (
    CatalogStore,
    DescriptionOrigin,
    InspectionResult,
    MetadataStatus,
    PackageProfile,
    ParameterDef,
    PhysicalInfo,
    RunArtifact,
    SheetSummary,
    SourceEntry,
    compute_artifact_id,
    compute_params_hash,
    compute_source_id,
)
from excel_relay.core.refresher import ExcelComSession, HeadlessExcelRefresher, RefreshError
from excel_relay.core.version import __version__, get_version, get_version_info

__all__ = [
    "CatalogManager",
    "ExcelComSession",
    "ExcelInspector",
    "HeadlessExcelRefresher",
    "RefreshError",
    "CatalogStore",
    "DescriptionOrigin",
    "InspectionResult",
    "MetadataStatus",
    "PackageProfile",
    "ParameterDef",
    "PhysicalInfo",
    "RunArtifact",
    "SheetSummary",
    "SourceEntry",
    "__version__",
    "compute_artifact_id",
    "compute_params_hash",
    "compute_source_id",
    "get_version",
    "get_version_info",
]
