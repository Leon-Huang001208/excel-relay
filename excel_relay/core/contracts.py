"""项目底稿的日期参数和固定指标登记。"""

from __future__ import annotations

from datetime import date, time, timedelta
import json
import math
import os
from pathlib import Path
import re
import tempfile
from typing import Any
from xml.etree import ElementTree
from zipfile import BadZipFile, ZipFile

from openpyxl import load_workbook
from openpyxl.utils.cell import coordinate_to_tuple
from openpyxl.utils.datetime import MAC_EPOCH, WINDOWS_EPOCH, to_excel
from openpyxl.utils.exceptions import InvalidFileException
from pydantic import BaseModel, ConfigDict, Field, TypeAdapter, field_validator

from excel_relay.core.catalog import CascadingCatalog, CatalogManager
from excel_relay.core.inspector import coerce_numeric_scalar, is_valid_cell_ref, normalize_cell_ref
from excel_relay.core.models import CatalogScope, ParameterDef, SourceEntry


class OutputContract(BaseModel):
    """已登记的固定指标。单位描述原始单元格值的业务口径。"""

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    name: str = Field(min_length=1)
    cell: str
    unit: str

    @field_validator("cell")
    @classmethod
    def validate_cell(cls, value: str) -> str:
        cell = normalize_cell_ref(value)
        if "!" in cell and is_valid_cell_ref(cell):
            row, col = coordinate_to_tuple(cell.rsplit("!", 1)[1])
            if 1 <= row <= 1048576 and 1 <= col <= 16384:
                return cell
        raise ValueError("指标位置必须是有效的单个工作表坐标，例如 海外!C2")


class SourceContract(BaseModel):
    """日期参数类型和结果单元格，不包含本次或历史结果。"""

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)

    date_inputs: list[str] = Field(default_factory=list)
    outputs: list[OutputContract] = Field(default_factory=list)

    @field_validator("outputs")
    @classmethod
    def distinct_output_names(cls, value: list[OutputContract]) -> list[OutputContract]:
        if len({output.name.casefold() for output in value}) != len(value):
            raise ValueError("指标名称不能重复")
        return value


_CONTRACTS = TypeAdapter(dict[str, SourceContract])
_WORKBOOK_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def _project(catalog: CatalogManager | CascadingCatalog) -> CatalogManager | None:
    primary = catalog.primary if isinstance(catalog, CascadingCatalog) else catalog
    if primary.scope == CatalogScope.PROJECT and primary.is_initialized():
        return primary
    return None


def _load_contracts(path: Path) -> dict[str, SourceContract]:
    try:
        return _CONTRACTS.validate_json(path.read_bytes())
    except FileNotFoundError:
        return {}
    except (OSError, ValueError) as error:
        raise ValueError(f"无法读取项目指标登记 {path}: {error}") from error


def _match_input(inputs: list[ParameterDef], key: str) -> ParameterDef | None:
    name = key.strip().lower()
    cell = normalize_cell_ref(key).lower()
    return next(
        (item for item in inputs if name == item.name.strip().lower()
         or cell == normalize_cell_ref(item.cell).lower()),
        None,
    )


def _date_inputs(entry: SourceEntry, contract: SourceContract) -> list[ParameterDef]:
    selected = []
    for name in contract.date_inputs:
        item = _match_input(entry.inputs, name)
        if item is None:
            raise ValueError(f"日期参数 '{name}' 未在底稿输入白名单中登记")
        if item in selected:
            raise ValueError(f"日期参数 '{item.name}' 已重复声明")
        selected.append(item)
    return selected


def configure_source_contract(
    catalog: CatalogManager | CascadingCatalog,
    source_id: str,
    spec: dict[str, Any] | SourceContract,
) -> SourceContract:
    """为已登记的项目底稿原子保存合同，供旧版目录和新版同时使用。"""
    project = _project(catalog)
    if project is None:
        raise ValueError("指标登记需要已初始化的项目目录")
    contract = SourceContract.model_validate(spec)
    with project._exclusive_lock():
        entry = project.get_source(source_id)
        if entry is None or entry.scope != CatalogScope.PROJECT:
            raise ValueError(f"当前项目未登记底稿 '{source_id}'")
        if project.resolve_absolute_path(entry).suffix.lower() not in {".xlsx", ".xlsm"}:
            raise ValueError("指标登记支持 .xlsx 和 .xlsm 工作簿")
        sheet_names = {sheet.name.casefold(): sheet.name for sheet in entry.physical.sheets}
        outputs = []
        for output in contract.outputs:
            sheet, coordinate = output.cell.rsplit("!", 1)
            if sheet.casefold() not in sheet_names:
                raise ValueError(f"指标 '{output.name}' 引用的工作表 '{sheet}' 未登记")
            outputs.append(output.model_copy(update={"cell": f"{sheet_names[sheet.casefold()]}!{coordinate}"}))
        contract = SourceContract(
            date_inputs=[item.name for item in _date_inputs(entry, contract)],
            outputs=outputs,
        )
        path = project.catalog_dir / "contracts.json"
        contracts = _load_contracts(path)
        contracts[entry.source_id] = contract
        tmp_path = None
        try:
            with tempfile.NamedTemporaryFile(
                mode="w", encoding="utf-8", prefix="contracts.", suffix=".tmp",
                dir=project.catalog_dir, delete=False,
            ) as tmp:
                tmp_path = Path(tmp.name)
                json.dump(_CONTRACTS.dump_python(contracts, mode="json"), tmp, ensure_ascii=False, indent=2)
                tmp.write("\n")
                tmp.flush()
                os.fsync(tmp.fileno())
            os.replace(tmp_path, path)
        finally:
            if tmp_path is not None:
                tmp_path.unlink(missing_ok=True)
    return contract


def get_source_contract(
    catalog: CatalogManager | CascadingCatalog,
    entry: SourceEntry,
) -> SourceContract | None:
    """从当前项目读取一份底稿的指标登记。"""
    project = _project(catalog)
    if project is None or entry.scope != CatalogScope.PROJECT:
        return None
    return _load_contracts(project.catalog_dir / "contracts.json").get(entry.source_id)


def _workbook_epoch(path: Path):
    """仅解压工作簿元数据，轻量读取 1900/1904 日期系统。"""
    try:
        with ZipFile(path) as workbook:
            root = ElementTree.fromstring(workbook.read("xl/workbook.xml"))
        props = root.find(f"{_WORKBOOK_NS}workbookPr")
        if props is not None and props.get("date1904") in {"1", "true"}:
            return MAC_EPOCH
        return WINDOWS_EPOCH
    except (OSError, KeyError, BadZipFile, ElementTree.ParseError) as error:
        raise ValueError(f"无法读取工作簿的 Excel 日期系统: {path}") from error


def normalize_contract_params(
    path: Path,
    entry: SourceEntry,
    params: dict[str, Any] | None,
    contract: SourceContract | None,
) -> dict[str, Any] | None:
    """为已声明的日期参数转换 ISO 日期，保留原参数键和数字序列值。"""
    if not contract or not contract.date_inputs:
        return params
    dates = _date_inputs(entry, contract)
    if not params:
        return params
    result = params.copy()
    epoch = None
    for key, value in params.items():
        if _match_input(entry.inputs, key) not in dates:
            continue
        numeric = coerce_numeric_scalar(value)
        if type(numeric) is int or (type(numeric) is float and math.isfinite(numeric)):
            result[key] = numeric
            continue
        error = f"日期参数 '{key}' 必须是有效的 YYYY-MM-DD 日期或 Excel 序列值"
        if not isinstance(value, str) or not re.fullmatch(r"[0-9]{4}-[0-9]{2}-[0-9]{2}", value):
            raise ValueError(error)
        try:
            day = date.fromisoformat(value)
        except ValueError as invalid_date:
            raise ValueError(error) from invalid_date
        if epoch is None:
            epoch = _workbook_epoch(path)
        result[key] = int(to_excel(day, epoch))
    return result


def _output_value(workbook, output: OutputContract) -> Any:
    sheet, coordinate = output.cell.rsplit("!", 1)
    if sheet not in workbook.sheetnames:
        raise ValueError(f"工作表 '{sheet}' 缺失")
    worksheet = workbook[sheet]
    row, col = coordinate_to_tuple(coordinate)
    if (worksheet.max_row is not None and row > worksheet.max_row) or (
        worksheet.max_column is not None and col > worksheet.max_column
    ):
        raise ValueError("指标单元格位置缺失")
    cell = worksheet[coordinate]
    if cell.data_type == "e":
        raise ValueError(f"Excel 单元格报错: {cell.value}")
    value = cell.value
    if value is None or value == "":
        raise ValueError("指标单元格空白或公式缓存缺失")
    if isinstance(value, (date, time)):
        return value.isoformat()
    if isinstance(value, timedelta):
        return str(value)
    if isinstance(value, float) and not math.isfinite(value):
        raise ValueError("指标单元格存在非有限数值")
    return value


def read_contract_outputs(path: Path, contract: SourceContract) -> list[dict[str, Any]]:
    """逐项读取本次已保存副本的缓存值，规整空白与 Excel 错误。"""
    if not contract.outputs:
        return []
    read_errors = (OSError, KeyError, ValueError, BadZipFile, ElementTree.ParseError, InvalidFileException)
    try:
        workbook = load_workbook(path, data_only=True, read_only=True, keep_links=False)
    except read_errors as error:
        return [
            {**output.model_dump(), "status": "error", "error": f"无法读取本次结果文件: {error}"}
            for output in contract.outputs
        ]

    outcomes = []
    try:
        for output in contract.outputs:
            item = output.model_dump()
            try:
                value = _output_value(workbook, output)
            except read_errors as error:
                outcomes.append({**item, "status": "error", "error": str(error)})
            else:
                outcomes.append({**item, "status": "ok", "value": value})
    finally:
        workbook.close()
    return outcomes
