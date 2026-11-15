"""Pydantic 数据模型定义 (Models & Schemas).

围绕 catalog.json 持久化协议构建核心领域模型：
1. source_id 由文件规范化路径对称确定性生成 (项目级: proj_{6位Hash}_{文件名} / 全局级: glob_{6位Hash}_{文件名})；
2. SourceEntry 采用仓储身份 + 四大正交领域组件 (physical, profile, inputs, artifacts) 组合设计；
3. 遵循极简减负原则设计 ParameterDef，不在 catalog 中存储残缺的数据输出快照。
"""

from __future__ import annotations

import hashlib
import json
import re
from datetime import datetime
from enum import Enum
from pathlib import Path
from typing import Any, Dict, List, Optional

from pydantic import BaseModel, Field

from excel_relay.core.version import CATALOG_PROTOCOL_VERSION


class MetadataStatus(str, Enum):
    """元数据完整度与置信度状态."""
    RAW = "raw"            # 仅包含纯代码提取的基础事实 (文件名/Sheet名/表头/公式)
    ENRICHED = "enriched"  # 经由宿主 Agent 或用户通过 Catalog 补充了业务语义
    DECLARED = "declared"  # 模型作者在 Excel 内置 _excel_relay 表中规范声明


class DescriptionOrigin(str, Enum):
    """描述信息的生成来源."""
    EXTRACTED = "extracted"          # 程序从结构中自动提取拼接的事实
    AGENT_GENERATED = "agent"        # 宿主 Agent (LLM) 提炼生成的业务解释
    AUTHOR_DECLARED = "author"       # 制作分析师在 _excel_relay 表中声明的
    USER_MANUAL = "user"             # 用户手动通过 CLI 设置的


class CatalogScope(str, Enum):
    """数据源与注册中心作用域分层."""
    PROJECT = "project"  # 项目专属数据源
    GLOBAL = "global"    # 用户全局数据源


def compute_source_id(
    file_path: str | Path,
    root_dir: Optional[Path] = None,
) -> str:
    """根据文件路径与作用域规范化生成对称确定性 source_id.

    规则 (严格对称的 {4字母作用域前缀}_{6位路径Hash}_{清洗后文件名} 结构):
    1. 若指定了 root_dir 且文件位于 root_dir 下 (项目专属)，按相对于 root_dir 的相对路径计算哈希:
       格式: proj_{6位相对路径Hash}_{清洗后文件名}，如 'proj_a1f9c2_贵州茅台估值模型'。
       项目目录在不同机器间克隆或移动时，source_id 保持不变，实现完全便携。
    2. 否则 (全局数据源)，按绝对路径计算哈希:
       格式: glob_{6位绝对路径Hash}_{清洗后文件名}，如 'glob_e83b12_贵州茅台估值模型'。
    """
    resolved = Path(file_path).resolve()
    clean_stem = re.sub(r"[^\w\.-]", "_", resolved.stem).strip("_") or "workbook"

    if root_dir is not None:
        try:
            rel = resolved.relative_to(Path(root_dir).resolve())
            norm_rel = str(rel).replace("\\", "/").lower()
            hash6 = hashlib.sha256(norm_rel.encode("utf-8")).hexdigest()[:6]
            return f"proj_{hash6}_{clean_stem}"
        except ValueError:
            pass

    # 统一将 Windows 反斜杠转换为正斜杠并小写化，确保 Windows/WSL/Linux 跨平台绝对路径哈希中立一致
    norm_path_str = str(resolved).replace("\\", "/").lower()
    hash6 = hashlib.sha256(norm_path_str.encode("utf-8")).hexdigest()[:6]
    return f"glob_{hash6}_{clean_stem}"


def compute_params_hash(params: Optional[Dict[str, Any]]) -> str:
    """计算全量参数配置字典 X 的唯一规范化签名哈希值 (6位十六进制)."""
    if not params:
        return "base"
    try:
        normalized = json.dumps(params, sort_keys=True, ensure_ascii=False)
        return hashlib.sha256(normalized.encode("utf-8")).hexdigest()[:6]
    except Exception:
        return "custom"


def compute_artifact_id(source_id: str, params_hash: str) -> str:
    """根据母版 source_id 与全量参数哈希 params_hash 构建唯一确定性 artifact_id.

    格式 (严格对称):
    - 项目级: proj_{6位相对Hash}_{参数Hash}，如 'proj_a1f9c2_8b3e10'
    - 全局级: glob_{6位绝对Hash}_{参数Hash}，如 'glob_e83b12_8b3e10'
    同一母版在相同全量参数配置 X 下永远对应唯一一个 artifact_id。
    """
    parts = source_id.split("_", 2)
    source_prefix = f"{parts[0]}_{parts[1]}" if len(parts) >= 2 else source_id
    return f"{source_prefix}_{params_hash}"


class ParameterDef(BaseModel):
    """可调输入参数定义 (极简 4 字段设计，零语法负担)."""
    cell: str = Field(..., description="参数所在单个单元格的规范化坐标 'Sheet!Cell'，如 假设表!C4 (不支持命名区域)")
    name: str = Field(..., description="参数名称，如 提价幅度 (未指定时默认使用 cell)")
    current_value: Optional[Any] = Field(None, description="当前单元格基准值 (由探针自动从表格提取)")
    description: Optional[str] = Field(None, description="补充说明：单位、合理区间或业务口径，如 '单位%，合理范围 0~15'")


class SheetSummary(BaseModel):
    """工作表物理结构摘要."""
    name: str = Field(..., description="工作表名称 (Sheet Name)")
    row_count: int = Field(0, description="有效数据行数")
    col_count: int = Field(0, description="有效数据列数")
    header_row: int = Field(1, description="探测到的横向表头所在行号 (1-indexed)")
    columns: List[str] = Field(default_factory=list, description="横向表头字段列表")
    row_headers: List[str] = Field(default_factory=list, description="纵向首列/次列指标名列表 (适配财务三表与估值表)")
    has_formulas: bool = Field(False, description="工作表是否包含动态公式计算")


class PhysicalInfo(BaseModel):
    """文件客观物理事实与工作表三维结构摘要 (100% 由 Inspector 从物理文件探测提取，只读)."""
    file_size_bytes: int = Field(0, description="物理文件大小 (字节)")
    file_mtime: str = Field("", description="文件最后修改时间戳 (秒级浮点数字符串)")
    vendor: str = Field("none", description="自动嗅探识别的数据插件依赖: wind / ifind / none")
    has_live_plugins: bool = Field(False, description="是否包含万得或同花顺等实时网络刷新公式插件")
    has_official_spec: bool = Field(False, description="源文件物理上是否包含内置的 _excel_relay 规范工作表")
    sheets: List[SheetSummary] = Field(default_factory=list, description="物理工作表结构与表头字段摘要列表")


class PackageProfile(BaseModel):
    """数据源业务语义画像卡片 (由 _excel_relay 表提取或通过 Catalog 补充)."""
    name: str = Field(..., description="模型中文名称")
    summary: str = Field(..., description="核心功能一句话摘要，供 AI 检索路由")
    tags: List[str] = Field(default_factory=list, description="检索标签列表")
    questions: List[str] = Field(default_factory=list, description="典型提问方式列表")
    notes: Optional[str] = Field(None, description="测算口径说明、数据来源或作者署名备注")
    metadata_status: MetadataStatus = Field(MetadataStatus.RAW, description="元数据完备度: raw / enriched / declared")
    description_origin: DescriptionOrigin = Field(
        DescriptionOrigin.EXTRACTED,
        description="描述来源: extracted / agent / author / user",
    )


class RunArtifact(BaseModel):
    """基于 (母版 source_id, 全量参数配置 X) 唯一标识的沙箱副本产物记录."""
    artifact_id: str = Field(..., description="由作用域、母版短Hash与全量参数Hash构成的唯一确定性标识，如 proj_a1f9c2_8b3e10")
    source_id: str = Field(..., description="关联的数据源母版 ID")
    file_path: str = Field(..., description="副本产物规范化路径 (项目级存储相对路径，全局级存储绝对路径)")
    params: Optional[Dict[str, Any]] = Field(None, description="本次测算的全量参数配置字典 X (未传参项自动补全母版基准值)")
    params_hash: str = Field("base", description="全量参数配置 X 的规范化签名哈希值 (6位十六进制)")
    created_at: str = Field(default_factory=lambda: datetime.now().isoformat(), description="最新一次计算刷新时间戳 t")
    elapsed_seconds: float = Field(0.0, description="刷新耗时秒数")
    is_temporary: bool = Field(True, description="是否位于系统临时沙箱中")
    keep_formulas: bool = Field(False, description="是否保留公式 (False表示已脱水为纯数值快照)")


class InspectionResult(BaseModel):
    """ExcelInspector.inspect_file() 的离线探测返回对象 (仅存在于内存，不含任何 Catalog 注册表身份)."""
    file_path: str = Field(..., description="被探测文件的规范化物理绝对路径")
    physical: PhysicalInfo = Field(..., description="文件客观物理事实与工作表结构")
    profile: PackageProfile = Field(..., description="从文件提取的初始业务语义画像卡片")
    inputs: List[ParameterDef] = Field(default_factory=list, description="从文件内部 _excel_relay 规范表显式声明中提取的可调输入参数列表")


class SourceEntry(BaseModel):
    """注册表核心数据源条目 (Catalog Entry): 仓储身份 + 四大正交领域组件."""
    source_id: str = Field(..., description="数据源唯一标识 (项目级: proj_{hash6}_{stem} / 全局级: glob_{hash6}_{stem})")
    file_path: str = Field(..., description="文件规范化路径 (项目级存储相对路径，全局级存储绝对路径)")
    scope: CatalogScope = Field(CatalogScope.GLOBAL, description="数据源作用域: project (项目专属) / global (用户全局)")
    is_relative: bool = Field(False, description="磁盘上是否以相对路径持久化")
    project_name: Optional[str] = Field(None, description="所属项目别名")
    registered_at: str = Field(default_factory=lambda: datetime.now().isoformat(), description="首次注册登记到 Catalog 的时间戳")
    last_refreshed_at: Optional[str] = Field(None, description="母版最近一次原地刷新的时间戳")
    physical: PhysicalInfo = Field(..., description="组件 1: 文件客观物理事实与表结构 (只读，重扫时无条件刷最新)")
    profile: PackageProfile = Field(..., description="组件 2: 数据源业务语义画像卡片 (支持 Agent 赋能与 Smart Merge)")
    inputs: List[ParameterDef] = Field(
        default_factory=list,
        description="组件 3: 显式登记的可调输入参数契约列表 (若为空则为无参模型 f(t)，默认原地刷新且禁止传参；若非空则为带参模型 f(t,X)，每次刷新均生成沙箱副本 Artifact)",
    )
    artifacts: List[RunArtifact] = Field(default_factory=list, description="组件 4: 沙箱改参或另存生成的历史副本列表")


class CatalogStore(BaseModel):
    """注册中心持久化仓储结构 (catalog.json)."""
    version: str = Field(CATALOG_PROTOCOL_VERSION, description="Catalog 仓储协议格式版本号")
    scope: CatalogScope = Field(CatalogScope.GLOBAL, description="仓储级别: project / global")
    project_name: Optional[str] = Field(None, description="所属项目名称 (仅在 project 级生效)")
    updated_at: str = Field(default_factory=lambda: datetime.now().isoformat(), description="注册表最后一次持久化更新时间戳")
    sources: Dict[str, SourceEntry] = Field(default_factory=dict, description="已登记的数据源字典 {source_id: SourceEntry}")
