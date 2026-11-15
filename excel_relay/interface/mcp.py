"""excel-relay MCP 服务端实现 (Model Context Protocol Server).

基于 mcp.server.fastmcp 提供面向 AI Agent 的标准 RPC 工具集。
支持工作区上下文感知 (Workspace Root) 与项目专属/用户全局级联数据源架构。
"""

from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Any, Dict, List, Optional

from mcp.server.fastmcp import FastMCP

from excel_relay.core import (
    CascadingCatalog,
    CatalogManager,
    CatalogScope,
    DescriptionOrigin,
    HeadlessExcelRefresher,
    find_project_root,
    get_cascading_catalog,
    get_version as core_get_version,
    get_version_info,
)

logger = logging.getLogger("excel-relay.mcp")

MCP_INSTRUCTIONS = """[Excel Relay安全与操作守则]
【显式触发条件】
当用户提示词中出现“使用excel-relay”、“使用excel-relay数据源”、“使用Excel数据源”或涉及本地投研底稿（宏观/行业高频数据表、固收/量化底稿、参数化取数模板、股票估值模型等）时，必须自动激活并优先使用 excel-relay MCP 工具集。

【投研数据作业标准三步法 (Mandatory 3-Step Workflow)】
所有通过 excel-relay 处理投研数据的任务，必须严格遵循以下三步线性闭环，严禁跳步或直接读取需动态刷新的母版原文件：
1. 第一步【查找定位】：优先调用 search_data_sources 检索目标底稿与参数契约，获取 source_id、外部插件依赖状态 (has_plugin_dependency) 及已登记输入参数清单 (inputs)。
2. 第二步【刷新/带参计算】：调用 refresh_data_source 驱动真实 Excel 后台刷新或带参计算（若含万得/同花顺插件或需最新动态数据则无参刷新；若需传入查询参数提取特定数据或调整假设测算，则传入白名单 params 在沙箱副本中计算）。严禁跳过本步直接读取需动态计算的原文件！
3. 第三步【自主读取与交付】：等待第二步返回物理结果文件 (result_file_path) 后，使用你自身的文件读取/Python 能力打开该结果文件，按用户实际意图提取目标数据或交付分析报告（若仅是函数式传参取数，直接交付数据即可，无需做多余的敏感性对比）。

macOS 使用桌面 Excel 与 Wind 加载项刷新 .xlsx，结果保存为保留公式的副本。

【底层安全红线】
- 母版只读防御：严禁使用通用文件写入工具（如 openpyxl、Python 文件写入等）直接修改已注册底稿母版原文件！
- 白名单契约：params 传参必须严格受限于 inputs 白名单，未登记参数禁止私自改写母版公式。
"""

mcp = FastMCP("excel-relay", instructions=MCP_INSTRUCTIONS)


@mcp.tool()
def search_data_sources(
    query: str,
    tag: Optional[str] = None,
    workspace_path: Optional[str] = None,
) -> str:
    """根据投研任务意图检索最匹配的已注册 Excel 数据源 (支持项目专属与全局数据源分层检索与优先加权).

    路由与数据读取重要规则:
    - 外部插件与鲜度优先: 若返回的 has_plugin_dependency 为 True (依赖万得/同花顺插件)，或用户需要最新/当前/实时行情与指标，严禁直接读取原文件！必须先调用 refresh_data_source 刷新，再读取返回的物理文件。
    - 静态直接读取: 仅当 has_plugin_dependency 为 False 且用户仅查询静态历史归档数据时，才可使用自身的 Python 能力直接读取 absolute_path。

    Args:
        query: 任务问题、标的、指标或关键词，涵盖宏观经济、行业高频、个股财务、固收利率、量化策略或任意投研数据表 (如 '宏观流动性跟踪'、'茅台出厂价'、'光伏排产'、'中债收益率')
        tag: 可选标签过滤
        workspace_path: 可选当前工程工作区路径，用于动态识别本项目上下文

    Returns:
        JSON 格式的最匹配数据源列表 (包含 scope 作用域、可调参数 inputs、插件依赖状态 has_plugin_dependency、鲜度警告及物理路径 absolute_path)
    """
    cascading = get_cascading_catalog(workspace_or_cwd=workspace_path)
    matches = cascading.search(query, top_k=5, tag_filter=tag)

    if not matches:
        return f"未找到与 '{query}' 相关的 Excel 数据源。"

    results = []
    for entry, score, reason in matches:
        abs_p = str(cascading.resolve_absolute_path(entry))
        has_plugin = entry.physical.vendor != "none"
        results.append({
            "source_id": entry.source_id,
            "scope": entry.scope.value,
            "project_name": entry.project_name,
            "name": entry.profile.name,
            "summary": entry.profile.summary,
            "match_score": round(score, 1),
            "match_reason": reason,
            "vendor": entry.physical.vendor,
            "has_plugin_dependency": has_plugin,
            "data_freshness_warning": (
                f"⚠️ 警告: 该底稿依赖 {entry.physical.vendor} 金融插件动态取数公式 (=[1]!...)！"
                "直接用 Python 读取原文件极易获取到历史静态残留或未解析数据。若用户需要最新行情/指标或重新计算，"
                "必须首先调用 refresh_data_source(source_id) 执行后台静默刷新，再读取生成的结果文件！"
                if has_plugin
                else "该底稿无外部插件依赖。若需调整参数测算请调 refresh_data_source，若仅查询历史静态数据可直接读取 absolute_path。"
            ),
            "recommended_questions": entry.profile.questions,
            "inputs": [i.model_dump() for i in entry.inputs],
            "file_path": entry.file_path,
            "absolute_path": abs_p,
        })

    return json.dumps(results, ensure_ascii=False, indent=2)


@mcp.tool()
def get_source_detail(source_id: str, workspace_path: Optional[str] = None) -> str:
    """获取指定 Excel 数据源的完整元数据、可调参数列表、工作表结构、外部插件依赖及物理绝对路径 (resolved_absolute_path).

    路由与数据读取重要规则:
    - 外部插件与鲜度优先: 若 has_plugin_dependency 为 True (依赖万得/同花顺插件)，或用户需要最新/当前/实时行情与指标，严禁直接读取原文件！必须先调用 refresh_data_source 刷新，再读取返回的物理文件。
    - 静态直接读取: 仅当 has_plugin_dependency 为 False 且用户仅查询静态历史归档数据时，才可使用自身的 Python 能力直接读取 resolved_absolute_path。

    Args:
        source_id: 数据源唯一 ID (支持完整 ID、前 6 位短 Hash 或文件路径)
        workspace_path: 可选工作区路径，用于定位项目专属数据源
    """
    cascading = get_cascading_catalog(workspace_or_cwd=workspace_path)
    entry = cascading.get_source(source_id)
    if not entry:
        return f"错误: 未找到 ID 为 '{source_id}' 的数据源。"

    data = entry.model_dump()
    has_plugin = entry.physical.vendor != "none"
    data["has_plugin_dependency"] = has_plugin
    data["data_freshness_warning"] = (
        f"⚠️ 警告: 该底稿依赖 {entry.physical.vendor} 金融插件动态取数公式 (=[1]!...)！"
        "直接用 Python 读取原文件极易获取到历史静态残留或未解析数据。若用户需要最新行情/指标或重新计算，"
        "必须首先调用 refresh_data_source(source_id) 执行后台静默刷新，再读取生成的结果文件！"
        if has_plugin
        else "该底稿无外部插件依赖。若需调整参数测算请调 refresh_data_source，若仅查询历史静态数据可直接读取 resolved_absolute_path。"
    )
    data["resolved_absolute_path"] = str(cascading.resolve_absolute_path(entry))
    return json.dumps(data, ensure_ascii=False, indent=2)


@mcp.tool()
def refresh_data_source(
    source_id: str,
    params: Optional[Dict[str, Any]] = None,
    output_path: Optional[str] = None,
    timeout_sec: int = 120,
    keep_formulas: bool = False,
    workspace_path: Optional[str] = None,
) -> str:
    """刷新指定 Excel 数据源，或使用已登记的参数计算副本。

    安全约束与使用指引:
    - macOS 使用桌面 Excel 和 Wind 加载项刷新 .xlsx，并将结果保存为保留公式的副本。
    - 【涉插件/最新数据无参刷新】: 若用户提问涉及'最新'、'当前'、'实时'数据，或底稿含有万得/同花顺插件，即使不改参数也必须无参调用本工具 (params=None) 进行最新数据收敛刷新；
    - 【带参沙箱计算 (函数式取数 / 假设测算)】: 仅当数据源在 Catalog 的 inputs 中明确登记了输入参数时，才允许传入 params（既可用于传入代码/日期/口径提取特定数据，也可用于修改假设进行情景测算）；
    - params 的键必须严格对应已登记参数的 name 或 cell 坐标，未登记参数的数据源仅支持无参刷新；
    - 若传参被白名单拦截，严禁私自调用 configure_source_inputs 修改白名单，必须向用户报告当前允许的参数范围并询问用户；
    - 产物副本将自动写入所属仓储的沙箱目录 (项目专属底稿存入项目 .excel-relay/artifacts/)；
    - 计算完成后，本工具返回结果文件物理路径 (result_file_path)，请务必使用你自身的 Excel/Python 读取能力打开该文件读取结果。

    Args:
        source_id: 数据源唯一 ID (支持完整 ID 或前 6 位短 Hash)
        params: 拟传入的参数字典，如 {"证券代码": "600519.SH"}、{"提价幅度": 5} 或 {"假设表!C4": 5}
        output_path: 可选另存路径。macOS 未指定时自动生成独立结果路径。
        timeout_sec: 刷新超时秒数，默认 120。Windows 在插件持续返回数据时可顺延。
        keep_formulas: Windows 副本是否保留公式，默认为 False。macOS 副本始终保留公式。
        workspace_path: 可选工作区路径，用于定位项目专属数据源
    """
    cascading = get_cascading_catalog(workspace_or_cwd=workspace_path)
    entry = cascading.get_source(source_id)
    if not entry:
        return f"错误: 未找到 ID 为 '{source_id}' 的数据源。"

    try:
        out_file, elapsed, run_artifact = HeadlessExcelRefresher.refresh(
            source_id_or_path=entry.source_id,
            params=params,
            output_path=output_path,
            timeout_sec=timeout_sec,
            keep_formulas=keep_formulas,
            catalog_mgr=cascading,
        )

        is_sandbox = run_artifact is not None
        payload: Dict[str, Any] = {
            "status": "success",
            "source_id": entry.source_id,
            "scope": entry.scope.value,
            "mode": "sandbox_copy" if is_sandbox else "in_place_refresh",
            "artifact_id": run_artifact.artifact_id if run_artifact else None,
            "result_file_path": out_file,
            "params": run_artifact.params if run_artifact else params,
            "params_hash": run_artifact.params_hash if run_artifact else "base",
            "keep_formulas": run_artifact.keep_formulas if run_artifact else True,
            "elapsed_seconds": elapsed,
            "message": f"成功在后台刷新数据！耗时 {elapsed} 秒，结果文件已保存至 {out_file}，请直接读取该 Excel 文件获取完整计算结果。",
        }

        return json.dumps(payload, ensure_ascii=False, indent=2)
    except Exception as e:
        logger.error("刷新数据源异常 (%s): %s", entry.source_id, e)
        return json.dumps({
            "status": "error",
            "source_id": entry.source_id,
            "error": str(e),
        }, ensure_ascii=False, indent=2)


@mcp.tool()
def list_source_artifacts(
    source_id: str,
    limit: int = 10,
    workspace_path: Optional[str] = None,
) -> str:
    """查询指定数据源的不同参数配置测算产物清单.

    Args:
        source_id: 数据源唯一 ID (支持完整 ID 或前 6 位短 Hash)
        limit: 返回条数上限，默认 10
        workspace_path: 可选工作区路径
    """
    cascading = get_cascading_catalog(workspace_or_cwd=workspace_path)
    entry = cascading.get_source(source_id)
    if not entry:
        return f"错误: 未找到 ID 为 '{source_id}' 的数据源。"

    artifacts = cascading.list_artifacts(entry.source_id, limit=limit)
    art_list = []
    for art in artifacts:
        d = art.model_dump()
        d["resolved_absolute_path"] = str(cascading.resolve_absolute_path(art))
        art_list.append(d)
    return json.dumps({
        "status": "success",
        "source_id": entry.source_id,
        "scope": entry.scope.value,
        "name": entry.profile.name,
        "count": len(art_list),
        "artifacts": art_list,
    }, ensure_ascii=False, indent=2)


@mcp.tool()
def get_artifact(
    source_id: str,
    artifact_id: Optional[str] = None,
    params: Optional[Dict[str, Any]] = None,
    workspace_path: Optional[str] = None,
) -> str:
    """根据产物唯一标识 (artifact_id) 或特定参数组合 X 精确查找对应的测算产物.

    Args:
        source_id: 数据源唯一 ID (支持完整 ID 或前 6 位短 Hash)
        artifact_id: 可选产物唯一标识，如 "proj_a1f9c2_8b3e10"
        params: 可选参数字典，如 {"提价幅度": 5}
        workspace_path: 可选工作区路径
    """
    cascading = get_cascading_catalog(workspace_or_cwd=workspace_path)
    entry = cascading.get_source(source_id)
    if not entry:
        return f"错误: 未找到 ID 为 '{source_id}' 的数据源。"

    art = cascading.find_artifact(entry.source_id, params=params, artifact_id=artifact_id)
    if not art:
        return json.dumps({
            "status": "not_found",
            "source_id": entry.source_id,
            "message": f"未找到符合条件的测算产物 (artifact_id={artifact_id}, params={params})",
        }, ensure_ascii=False, indent=2)

    art_data = art.model_dump()
    art_data["resolved_absolute_path"] = str(cascading.resolve_absolute_path(art))
    return json.dumps({
        "status": "success",
        "source_id": entry.source_id,
        "artifact": art_data,
    }, ensure_ascii=False, indent=2)


@mcp.tool()
def register_data_source(
    path: str,
    tags: Optional[List[str]] = None,
    workspace_path: Optional[str] = None,
    register_as_global: bool = False,
) -> str:
    """在对话中动态注册新的 Excel 文件或目录为数据源 (自动感知项目/全局).

    Args:
        path: 本地 Excel 文件或文件夹绝对路径
        tags: 可选自定义标签，如 ["汽车", "新势力"]
        workspace_path: 可选工作区路径，若在项目内默认注册为项目专属数据源
        register_as_global: 若为 True，强制注册为用户全局数据源
    """
    try:
        p = Path(path)
        if not p.is_absolute() and workspace_path:
            target_abs = (Path(workspace_path) / p).resolve()
        else:
            target_abs = p.resolve()

        if register_as_global:
            target_mgr = CatalogManager()
        else:
            proj_root = find_project_root(target_abs)
            target_mgr = CatalogManager(base_dir=proj_root) if proj_root else CatalogManager()

        results = target_mgr.register_path(target_abs, custom_tags=tags)
        if not results:
            return f"警告: 在 {path} 中未发现有效的 Excel 文件 (.xlsx/.xlsm/.xls)。"

        return json.dumps({
            "status": "success",
            "target_scope": target_mgr.store.scope.value,
            "registered_count": len(results),
            "sources": [
                {
                    "source_id": r.source_id,
                    "scope": r.scope.value,
                    "name": r.profile.name,
                    "summary": r.profile.summary,
                    "metadata_status": r.profile.metadata_status.value,
                    "description_origin": r.profile.description_origin.value,
                    "sheets": [s.name for s in r.physical.sheets],
                    "inputs_count": len(r.inputs),
                }
                for r in results
            ],
        }, ensure_ascii=False, indent=2)
    except Exception as e:
        logger.error("注册数据源异常 (%s): %s", path, e)
        return json.dumps({
            "status": "error",
            "error": str(e),
        }, ensure_ascii=False, indent=2)


@mcp.tool()
def update_source_metadata(
    source_id: str,
    name: Optional[str] = None,
    summary: Optional[str] = None,
    tags: Optional[List[str]] = None,
    questions: Optional[List[str]] = None,
    notes: Optional[str] = None,
    workspace_path: Optional[str] = None,
) -> str:
    """为已注册数据源补充或优化业务语义画像 (PackageProfile).

    本工具仅更新描述性业务语义 (名称、摘要、标签、推荐问法、口径备注)，不涉及可调参数契约修改，
    允许 Agent 在读取理解 Excel 内容后自主调用以丰富数据源画像。

    Args:
        source_id: 数据源唯一 ID (支持完整 ID 或前 6 位短 Hash)
        name: 可选模型中文名称
        summary: 补充的业务功能摘要
        tags: 补充的检索标签列表
        questions: 补充的推荐问法列表
        notes: 补充的测算口径或背景说明
        workspace_path: 可选工作区路径
    """
    cascading = get_cascading_catalog(workspace_or_cwd=workspace_path)
    try:
        entry = cascading.update_source(
            source_id=source_id,
            name=name,
            summary=summary,
            tags=tags,
            questions=questions,
            notes=notes,
            origin=DescriptionOrigin.AGENT_GENERATED,
        )
    except Exception as e:
        logger.error("更新数据源业务语义异常 (%s): %s", source_id, e)
        return json.dumps({
            "status": "error",
            "source_id": source_id,
            "error": str(e),
        }, ensure_ascii=False, indent=2)

    if not entry:
        return f"错误: 未找到 ID 为 '{source_id}' 的数据源。"

    return json.dumps({
        "status": "success",
        "source_id": entry.source_id,
        "scope": entry.scope.value,
        "metadata_status": entry.profile.metadata_status.value,
        "description_origin": entry.profile.description_origin.value,
        "name": entry.profile.name,
        "summary": entry.profile.summary,
        "tags": entry.profile.tags,
        "questions": entry.profile.questions,
        "notes": entry.profile.notes,
        "message": f"成功为数据源 {entry.source_id} ({entry.scope.value}) 更新业务语义画像！",
    }, ensure_ascii=False, indent=2)


@mcp.tool()
def configure_source_inputs(
    source_id: str,
    inputs: List[str],
    workspace_path: Optional[str] = None,
) -> str:
    """为已注册的 Excel 数据源配置可调输入参数白名单契约 (SourceEntry.inputs).

    安全红线 (严格遵守):
    - 严禁 Agent 自行通过启发式猜测调用本工具登记可调参数；
    - 当 refresh_data_source 因参数未登记而报错拦截时，严禁私自调用本工具绕过拦截；
    - 仅当用户在对话中明确指定了参数单元格坐标 (或经用户明确确认同意后)，才允许调用本工具。

    Args:
        source_id: 数据源唯一 ID (支持完整 ID 或前 6 位短 Hash)
        inputs: 用户明确指定的可调输入参数列表，如 ["提价幅度=假设表!C4 (单位%)", "折现率=WACC!B8"]；传入空列表 [] 可清空可调参数
        workspace_path: 可选工作区路径
    """
    cascading = get_cascading_catalog(workspace_or_cwd=workspace_path)
    try:
        entry = cascading.set_source_inputs(
            source_id=source_id,
            inputs=inputs,
        )
    except Exception as e:
        logger.error("更新数据源业务语义异常 (%s): %s", source_id, e)
        return json.dumps({
            "status": "error",
            "source_id": source_id,
            "error": str(e),
        }, ensure_ascii=False, indent=2)

    if not entry:
        return f"错误: 未找到 ID 为 '{source_id}' 的数据源。"

    return json.dumps({
        "status": "success",
        "source_id": entry.source_id,
        "scope": entry.scope.value,
        "inputs": [i.model_dump() for i in entry.inputs],
        "message": f"成功为数据源 {entry.source_id} ({entry.scope.value}) 配置 {len(entry.inputs)} 个可调输入参数！",
    }, ensure_ascii=False, indent=2)


@mcp.tool()
def clean_catalog(
    clean_stale_sources: bool = True,
    max_artifact_age_days: int = 7,
    workspace_path: Optional[str] = None,
) -> str:
    """清理 Catalog 中因文件移动/改名/删除而失效的 source 条目、孤儿产物目录，以及过期的临时沙箱副本.

    Args:
        clean_stale_sources: 是否清理源文件已不存在的失效数据源条目，默认 True
        max_artifact_age_days: 清理超过指定天数的临时沙箱副本，默认 7 天
        workspace_path: 可选工作区路径
    """
    cascading = get_cascading_catalog(workspace_or_cwd=workspace_path)
    removed_sources = cascading.clean_stale_sources() if clean_stale_sources else []
    removed_orphans = cascading.clean_orphan_artifact_dirs()
    cleaned_artifacts = cascading.clean_expired_artifacts(max_age_days=max_artifact_age_days)

    return json.dumps({
        "status": "success",
        "removed_stale_sources_count": len(removed_sources),
        "removed_stale_sources": [
            {"source_id": sid, "file_path": fp, "scope": sc.value}
            for sid, fp, sc in removed_sources
        ],
        "removed_orphan_dirs_count": len(removed_orphans),
        "removed_orphan_dirs": [
            {"source_id": sid, "dir_path": dp, "scope": sc.value}
            for sid, dp, sc in removed_orphans
        ],
        "cleaned_expired_artifacts_count": cleaned_artifacts,
    }, ensure_ascii=False, indent=2)


@mcp.tool()
def get_version() -> str:
    """获取 excel-relay 引擎版本、Catalog 仓储协议格式版本及系统运行环境信息。

    Returns:
        JSON 格式的版本与环境详情，包含 version, catalog_version, python_version, platform, system 等字段。
    """
    info = get_version_info()
    return json.dumps(info, ensure_ascii=False, indent=2)


def run_server():
    """启动 MCP Server."""
    mcp.run()
