"""excel-relay 核心功能引擎 (Core Domain Package).

包含数据模型、离线元数据探针、注册中心存储及COM静默重算引擎。
"""

from __future__ import annotations

from excel_relay.core.catalog import CascadingCatalog, CatalogManager, UnregisterResult
from excel_relay.core.context import (
    find_project_root,
    get_cascading_catalog,
    is_project_dir,
    is_sensitive_dir,
)
from excel_relay.core.inspector import (
    ExcelInspector,
    build_canonical_params,
    clean_scalar_value,
    coerce_numeric_scalar,
    normalize_cell_ref,
)
from excel_relay.core.models import (
    CatalogScope,
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
from excel_relay.core.version import (
    CATALOG_PROTOCOL_VERSION,
    __version__,
    get_version,
    get_version_info,
)

__all__ = [
    "CATALOG_PROTOCOL_VERSION",
    "CascadingCatalog",
    "CatalogManager",
    "CatalogScope",
    "CatalogStore",
    "DescriptionOrigin",
    "ExcelComSession",
    "ExcelInspector",
    "HeadlessExcelRefresher",
    "InspectionResult",
    "MetadataStatus",
    "PackageProfile",
    "ParameterDef",
    "PhysicalInfo",
    "RefreshError",
    "RunArtifact",
    "SheetSummary",
    "SourceEntry",
    "UnregisterResult",
    "__version__",
    "build_canonical_params",
    "clean_scalar_value",
    "coerce_numeric_scalar",
    "compute_artifact_id",
    "compute_params_hash",
    "compute_source_id",
    "find_project_root",
    "get_cascading_catalog",
    "get_version",
    "get_version_info",
    "is_project_dir",
    "is_sensitive_dir",
    "normalize_cell_ref",
]
