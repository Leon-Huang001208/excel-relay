"""Excel 元数据离线探针与规范解析器 (Inspector & Spec Parser).

核心设计职责：
1. 纯离线毫秒级解析 (.xlsx/.xlsm 用 openpyxl，.xls 用 xlrd)，不启动 Excel 进程；
2. 遵循职责隔离原则，仅产出纯文件探测对象 InspectionResult (physical + profile + inputs)，不涉及任何 Catalog 身份；
3. 解析全中文两段式 _excel_relay 规范表；
4. 保守确定性提取 ParameterDef (仅支持按 'Sheet!Cell' 单元格坐标声明的单格可调输入参数，不支持命名区域) 及其当前基准值；
5. 智能识别工作表真实表头行 (columns)、纵向首列指标名 (row_headers) 与金融插件依赖。
"""

from __future__ import annotations

import datetime as dt
import logging
import math
import os
import re
import shutil
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import coordinate_to_tuple
from openpyxl.workbook import Workbook
import xlrd

from excel_relay.core.models import (
    DescriptionOrigin,
    InspectionResult,
    MetadataStatus,
    PackageProfile,
    ParameterDef,
    PhysicalInfo,
    SheetSummary,
)

logger = logging.getLogger("excel-relay.inspector")

WIND_FORMULA_PATTERNS = [
    re.compile(r"\bWSD\s*\(", re.IGNORECASE),
    re.compile(r"\bWSS\s*\(", re.IGNORECASE),
    re.compile(r"\bWSQ\s*\(", re.IGNORECASE),
    re.compile(r"\bEDB\s*\(", re.IGNORECASE),
    re.compile(r"\bWSET\s*\(", re.IGNORECASE),
    re.compile(r"\bWSI\s*\(", re.IGNORECASE),
    re.compile(r"\bWST\s*\(", re.IGNORECASE),
    re.compile(r"\bWTD\s*\(", re.IGNORECASE),
    re.compile(r"\bWPF\s*\(", re.IGNORECASE),
    re.compile(r"\bTDAYS\s*\(", re.IGNORECASE),
    re.compile(r"\bTDAYSOFFSET\s*\(", re.IGNORECASE),
    re.compile(r"\bTDAYSCOUNT\s*\(", re.IGNORECASE),
    re.compile(r"WindFunc(?:\.xla)?", re.IGNORECASE),
]

IFIND_FORMULA_PATTERNS = [
    re.compile(r"\bTHS_\w+\s*\(", re.IGNORECASE),
    re.compile(r"\bTHS\s*\(", re.IGNORECASE),
    re.compile(r"ThsXls(?:\.xla)?", re.IGNORECASE),
]

# 单元格坐标与非法工作表字符匹配正则
CELL_COORD_RE = re.compile(r"^\$?[A-Za-z]{1,3}\$?\d+$")
INVALID_SHEET_CHARS_RE = re.compile(r"[:\\/\?\*\[\]]")


def coerce_numeric_scalar(val: Any) -> Any:
    """尝试将纯数字字符串安全提升为 int 或 float (如 '5' -> 5, '0.05' -> 0.05).

    避免外部传入字符串数字写入 Excel 单元格导致文本格式数字公式失效，并确保参数哈希一致性。
    """
    if val is None or isinstance(val, (int, float, bool)):
        return val
    if isinstance(val, str):
        s = val.strip()
        # 匹配整型数字
        if re.match(r"^-?\d+$", s):
            try:
                return int(s)
            except Exception:
                pass
        # 匹配标准浮点数字
        elif re.match(r"^-?\d+\.\d+$", s):
            try:
                return float(s)
            except Exception:
                pass
    return val


def clean_scalar_value(val: Any) -> Any:
    """将从 Excel 读取的单元格值转换为可 JSON 序列化的干净标量."""
    if val is None:
        return None
    if isinstance(val, (dt.datetime, dt.date, dt.time)):
        return val.isoformat()
    if isinstance(val, float):
        if math.isnan(val) or math.isinf(val):
            return None
        if val.is_integer():
            return int(val)
        return round(val, 6)
    if isinstance(val, (int, bool)):
        return val
    s = str(val).strip()
    return s if s != "" else None


def _cell_to_str_openpyxl(val: Any) -> str:
    """安全提取 openpyxl 单元格文本或公式字符串."""
    if val is None:
        return ""
    if hasattr(val, "text"):
        val = val.text
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    return str(val).strip()


def normalize_cell_ref(raw_ref: str) -> str:
    """规范化用户在 _excel_relay 表或 CLI 中填写的单个单元格坐标引用.

    支持:
    - '=假设表!$C$4' -> '假设表!C4'
    - "='假设表'!$C$4" -> '假设表!C4'
    - "='Q1 2025'!$C$4" -> 'Q1 2025!C4'
    - '假设表!C4' -> '假设表!C4'
    """
    if not raw_ref:
        return ""
    s = str(raw_ref).strip()
    if s.startswith("="):
        s = s[1:].strip()

    # 去除绝对引用符号 $
    s = s.replace("$", "")

    # 规范化 Sheet!Addr 中的单引号并去除单元格绝对引用
    if "!" in s:
        sheet_part, addr_part = s.rsplit("!", 1)
        sheet_part = sheet_part.strip()
        if sheet_part.startswith("'") and sheet_part.endswith("'") and len(sheet_part) >= 2:
            sheet_part = sheet_part[1:-1].replace("''", "'")
        return f"{sheet_part}!{addr_part.strip().upper()}"

    return s.upper()


def is_valid_cell_ref(ref: str) -> bool:
    """判断字符串是否为合法的单个单元格坐标语法 (支持带工作表前缀或裸坐标).

    规则:
    1. 严禁包含区域冒号 ':' (仅支持单个单元格)；
    2. 若包含 '!'：
       - 工作表部分: 长度 1~31，不能包含 Excel 非法字符 ':\\/?*[]'，不能为 '_excel_relay'；
         允许包含空格、括号、横杠等合法字符 (如 'Q1 2025!C4', '假设表 (最新)!C4')；
       - 单元格坐标部分: 符合 1~3 位列字母 + 行号 (如 C4, AA10, XFD1048576)；
    3. 若不含 '!'：
       - 单元格坐标部分符合 1~3 位列字母 + 行号。
    """
    if not ref or not str(ref).strip() or ":" in ref:
        return False

    s = str(ref).strip()
    if "!" in s:
        sheet_part, addr_part = s.rsplit("!", 1)
        sheet_part = sheet_part.strip()
        addr_part = addr_part.strip()
        # 校验工作表命名合规性
        if not sheet_part or len(sheet_part) > 31 or INVALID_SHEET_CHARS_RE.search(sheet_part):
            return False
        if sheet_part.lower() == "_excel_relay":
            return False
        return bool(CELL_COORD_RE.match(addr_part))

    return bool(CELL_COORD_RE.match(s))


def build_canonical_params(
    inputs: Optional[List[ParameterDef]],
    params: Optional[Dict[str, Any]] = None,
    strict: bool = True,
) -> Optional[Dict[str, Any]]:
    """构建带参模型 f(t, X) 的全量规范化参数配置字典 X (Single Source of Truth).

    - 若无参模型 f(t) (inputs 为空)：
      - 当 params 为空时返回 None；
      - 当 params 非空且 strict=True 时抛出 ValueError；strict=False 时清洗并返回。
    - 若带参模型 f(t, X) (inputs 非空)：
      - 以所有已登记参数的 current_value 构建基准字典；
      - 用传入的 params 按参数名或标准化坐标匹配覆盖对应项，并统一调用 clean_scalar_value；
      - 若存在未登记的参数键且 strict=True 时抛出 ValueError。
    """
    if not inputs:
        if not params:
            return None
        if strict:
            raise ValueError("该数据源未登记任何可调输入参数，不支持传入参数字典。")
        return {str(k).strip(): clean_scalar_value(coerce_numeric_scalar(v)) for k, v in params.items()}

    full_params: Dict[str, Any] = {
        inp.name: clean_scalar_value(coerce_numeric_scalar(inp.current_value))
        for inp in inputs
    }
    if params:
        for raw_key, val in params.items():
            key_clean = str(raw_key).strip()
            key_lower = key_clean.lower()
            key_norm_cell = normalize_cell_ref(key_clean).lower()
            matched_inp: Optional[ParameterDef] = None
            for inp in inputs:
                if inp.name.strip().lower() == key_lower or normalize_cell_ref(inp.cell).lower() == key_norm_cell:
                    matched_inp = inp
                    break
            if matched_inp is None:
                if strict:
                    valid_desc = ", ".join([f"{i.name}({i.cell})" for i in inputs])
                    raise ValueError(
                        f"参数 '{raw_key}' 未在该数据源的可调参数白名单中登记！当前允许的参数: [{valid_desc}]。"
                    )
                mapped_name = key_clean
            else:
                mapped_name = matched_inp.name
            # 数值类型安全提升：纯数字字符串自动转换为标准数字，确保参数字典规范哈希一致
            full_params[mapped_name] = clean_scalar_value(coerce_numeric_scalar(val))
    return full_params


def qualify_and_validate_ref(
    raw_ref: str,
    sheet_names: List[str],
) -> Optional[str]:
    """校验并规范化单个单元格引用，确保目标工作表真实存在，且自动补全缺省工作表名.

    仅接受单元格坐标 (不支持命名区域)：
    1. 若包含 'Sheet!Addr'，校验 Sheet 是否存在于 sheet_names 中 (排除 _excel_relay 规范表)，不存在则返回 None；
    2. 若仅填写了裸坐标 (如 'C4')，自动补全首个业务数据工作表名 (如 '假设表!C4')。
    """
    if not raw_ref or not str(raw_ref).strip():
        return None

    clean = normalize_cell_ref(str(raw_ref).strip())
    if not clean or not is_valid_cell_ref(clean):
        return None

    data_sheets = [sh for sh in sheet_names if sh.strip().lower() != "_excel_relay"]

    if "!" in clean:
        # 从右向左切分工作表与单元格坐标，保持全包切分口径严格一致
        sheet_part, addr_part = clean.rsplit("!", 1)
        matched_sheet: Optional[str] = None
        for sh in sheet_names:
            if sh.strip().lower() == sheet_part.strip().lower() and sh.strip().lower() != "_excel_relay":
                matched_sheet = sh
                break
        if not matched_sheet:
            return None
        return f"{matched_sheet}!{addr_part}"
    else:
        default_sheet = data_sheets[0] if data_sheets else (sheet_names[0] if sheet_names else None)
        if not default_sheet or default_sheet.strip().lower() == "_excel_relay":
            return None
        return f"{default_sheet}!{clean}"


def parse_inline_input_items(raw_text: str) -> List[Tuple[str, str, Optional[str]]]:
    """解析 CLI/MCP 传入的可调参数声明字符串.

    支持格式:
    1. '提价幅度=假设表!C4 (单位%, 0~15)'
    2. '假设表!C4 (单位%, 0~15)'
    3. '假设表!C4'

    返回: List[(name, cell, description)]
    """
    items: List[Tuple[str, str, Optional[str]]] = []
    if not raw_text or not str(raw_text).strip():
        return items

    for segment in re.split(r"[;；\n]+", str(raw_text)):
        seg = segment.strip()
        if not seg:
            continue

        # 提取括号内的备注说明
        desc: Optional[str] = None
        m_paren = re.search(r"[\(（]([^\)）]+)[\)）]\s*$", seg)
        if m_paren:
            desc = m_paren.group(1).strip()
            seg = seg[: m_paren.start()].strip()

        if "=" in seg and not seg.startswith("="):
            left, right = seg.split("=", 1)
            name = left.strip()
            cell = normalize_cell_ref(right.strip())
        else:
            cell = normalize_cell_ref(seg)
            name = cell

        if not cell:
            continue

        items.append((name or cell, cell, desc))
    return items


class ExcelInspector:
    """Excel 离线探测与规范解析器 (全面支持 .xlsx, .xlsm, .xls)."""

    @staticmethod
    def inspect_file(file_path: str | Path) -> InspectionResult:
        """完整探测单个 Excel 文件并生成纯文件探测数据包 InspectionResult."""
        path = Path(file_path).resolve()
        if not path.is_file():
            raise FileNotFoundError(f"文件不存在: {path}")

        stat = path.stat()
        ext = path.suffix.lower()

        if ext == ".xls":
            return ExcelInspector._inspect_xls(path, stat)
        elif ext in [".xlsx", ".xlsm"]:
            return ExcelInspector._inspect_openpyxl(path, stat)
        else:
            raise ValueError(f"不支持的文件格式: {path.suffix} (仅支持 .xlsx, .xlsm, .xls)")

    # -------------------------------------------------------------------------
    # openpyxl 实现 (.xlsx, .xlsm)
    # -------------------------------------------------------------------------
    @staticmethod
    def _inspect_openpyxl(path: Path, stat: os.stat_result) -> InspectionResult:
        wb_formula = openpyxl.load_workbook(path, read_only=False, data_only=False)
        try:
            sheet_names = list(wb_formula.sheetnames)

            # 1. 检查是否存在 _excel_relay 官方规范工作表 (大小写不敏感匹配，可调参数仅来源于 _excel_relay 表显式声明)
            spec_sheet_name = next(
                (s for s in sheet_names if s.strip().lower() == "_excel_relay"), None
            )
            has_official_spec = spec_sheet_name is not None
            if has_official_spec:
                rows = ExcelInspector._read_excel_relay_rows_openpyxl(wb_formula[spec_sheet_name])
                profile, inputs = ExcelInspector._build_profile_from_excel_relay_rows(
                    rows, path, sheet_names
                )
            else:
                profile = ExcelInspector._generate_fallback_profile(sheet_names, path)
                inputs = []

            # 2. 收集所有数据工作表的摘要信息与金融插件特征
            sheets_summary, detected_vendors = ExcelInspector._inspect_sheets_openpyxl(wb_formula)

            # 3. 深度嗅探 OOXML 的 externalLinks 外部加载项依赖 (覆盖 WindFunc.xla / iFinD 等)
            ext_vendors = ExcelInspector._inspect_external_links_ooxml(path)
            detected_vendors.update(ext_vendors)

            has_live_plugin = len(detected_vendors) > 0
            vendor_str = ",".join(sorted(detected_vendors)) if has_live_plugin else "none"
        finally:
            wb_formula.close()

        # 4. 若存在显式声明的 inputs，读取其当前基准数值 (data_only=True)
        if inputs:
            ExcelInspector._populate_input_values_openpyxl(path, inputs)

        physical = PhysicalInfo(
            file_size_bytes=stat.st_size,
            file_mtime=str(stat.st_mtime),
            vendor=vendor_str,
            has_live_plugins=has_live_plugin,
            has_official_spec=has_official_spec,
            sheets=sheets_summary,
        )
        return InspectionResult(
            file_path=str(path),
            physical=physical,
            profile=profile,
            inputs=inputs,
        )

    @staticmethod
    def _read_excel_relay_rows_openpyxl(ws) -> List[Tuple[str, str, str]]:
        """读取 _excel_relay 表 A/B/C 三列文本 (需以 data_only=False 加载，以保留 =Sheet!Cell 公式引用文本)."""
        rows: List[Tuple[str, str, str]] = []
        for row in ws.iter_rows(values_only=True):
            if not row:
                continue
            c0 = _cell_to_str_openpyxl(row[0]) if len(row) > 0 else ""
            c1 = _cell_to_str_openpyxl(row[1]) if len(row) > 1 else ""
            c2 = _cell_to_str_openpyxl(row[2]) if len(row) > 2 else ""
            rows.append((c0, c1, c2))
        return rows

    @staticmethod
    def _inspect_external_links_ooxml(path: Path) -> set[str]:
        """毫秒级解包解析 OOXML (.xlsx/.xlsm) 的 externalLinks，探测金融加载项依赖 (Wind/iFinD).

        通过直接读取 zip 包中的 xl/externalLinks/_rels/*.rels 与 xl/externalLinks/*.xml，
        精准识别对 WindFunc.xla、iFinD 等外部加载项的依赖，彻底避免受限于局部单元格采样行数与特定公式名称。
        """
        detected: set[str] = set()
        try:
            with zipfile.ZipFile(path, "r") as z:
                for name in z.namelist():
                    n_lower = name.lower()
                    # 1. 扫描外部链接关系文件 Target: xl/externalLinks/_rels/*.rels
                    if n_lower.startswith("xl/externallinks/_rels/") and n_lower.endswith(".rels"):
                        try:
                            content = z.read(name).decode("utf-8", errors="ignore").lower()
                            if any(k in content for k in ["windfunc", "wdf.addin", "windnet", "wind.net", "\\wind\\", "/wind/"]):
                                detected.add("wind")
                            if any(k in content for k in ["ths", "ifind", "thsxls", "thsslg", "\\ifind\\", "\\ths\\", "/ifind/", "/ths/"]):
                                detected.add("ifind")
                        except Exception:
                            pass
                    # 2. 扫描外部链接元数据: xl/externalLinks/*.xml (定义的外部宏/函数名称 definedNames)
                    elif n_lower.startswith("xl/externallinks/") and n_lower.endswith(".xml") and not n_lower.startswith("xl/externallinks/_rels/"):
                        try:
                            content = z.read(name).decode("utf-8", errors="ignore").lower()
                            if "windfunc" in content:
                                detected.add("wind")
                            if any(k in content for k in ["ths", "ifind"]):
                                detected.add("ifind")
                        except Exception:
                            pass
        except Exception:
            pass
        return detected

    @staticmethod
    def _inspect_sheets_openpyxl(wb: Workbook) -> Tuple[List[SheetSummary], set[str]]:
        summaries: List[SheetSummary] = []
        detected_vendors: set[str] = set()

        for name in wb.sheetnames:
            if name.strip().lower() == "_excel_relay":
                continue
            ws = wb[name]
            max_r = ws.max_row or 0
            max_c = ws.max_column or 0

            # 扫描前 40 行，提取真实表头行、首列纵向指标名以及插件公式
            scan_limit = min(max_r, 40)
            sampled_rows: List[List[Any]] = []
            sheet_has_formula = False

            for row in ws.iter_rows(min_row=1, max_row=max(scan_limit, 1), max_col=min(max_c or 1, 30), values_only=False):
                row_vals = []
                for cell in row:
                    val = cell.value
                    row_vals.append(_cell_to_str_openpyxl(val) if hasattr(val, "text") else val)
                    if val is not None:
                        s_val = _cell_to_str_openpyxl(val)
                        if s_val.startswith("="):
                            sheet_has_formula = True
                            for pat in WIND_FORMULA_PATTERNS:
                                if pat.search(s_val):
                                    detected_vendors.add("wind")
                            for pat in IFIND_FORMULA_PATTERNS:
                                if pat.search(s_val):
                                    detected_vendors.add("ifind")
                sampled_rows.append(row_vals)

            header_row_idx, columns, row_headers = ExcelInspector._extract_headers_from_grid(sampled_rows)

            summaries.append(
                SheetSummary(
                    name=name,
                    row_count=max_r,
                    col_count=max_c,
                    header_row=header_row_idx,
                    columns=columns,
                    row_headers=row_headers,
                    has_formulas=sheet_has_formula,
                )
            )

        return summaries, detected_vendors

    @staticmethod
    def _populate_input_values_openpyxl(
        path: Path,
        inputs: List[ParameterDef],
    ) -> None:
        """以 data_only=True 模式单次加载工作簿，读取 inputs 的当前基准数值."""
        wb_val = openpyxl.load_workbook(path, read_only=False, data_only=True)
        try:
            for inp in inputs:
                inp.current_value = ExcelInspector._read_ref_openpyxl(wb_val, inp.cell)
        finally:
            wb_val.close()

    @staticmethod
    def _read_ref_openpyxl(wb: Workbook, ref: str) -> Any:
        """从 openpyxl Workbook 中按 'Sheet!A1' 读取单个单元格标量基准值."""
        clean = normalize_cell_ref(ref)
        if not clean:
            return None

        sheet_name = None
        addr = clean
        if "!" in clean:
            sheet_name, addr = clean.rsplit("!", 1)

        if sheet_name:
            # Excel 官方工作表大小写无关 (Case-Insensitive)，通过小写比对准确定位工作表对象
            matched_ws = None
            target_sh_lower = sheet_name.strip().lower()
            for s in wb.sheetnames:
                if s.strip().lower() == target_sh_lower:
                    matched_ws = wb[s]
                    break
            if matched_ws is None:
                return None
            ws = matched_ws
        else:
            data_sheets = [s for s in wb.sheetnames if s.strip().lower() != "_excel_relay"]
            ws = wb[data_sheets[0]] if data_sheets else wb.active

        try:
            cell_obj = ws[addr]
            return clean_scalar_value(cell_obj.value)
        except Exception:
            return None

    # -------------------------------------------------------------------------
    # xlrd 实现 (.xls)
    # -------------------------------------------------------------------------
    @staticmethod
    def _inspect_xls(path: Path, stat: os.stat_result) -> InspectionResult:
        wb = xlrd.open_workbook(str(path))
        sheet_names = list(wb.sheet_names())

        # 1. 检查是否存在 _excel_relay 官方规范工作表 (大小写不敏感匹配，可调参数仅来源于 _excel_relay 表显式声明)
        spec_sheet_name = next(
            (s for s in sheet_names if s.strip().lower() == "_excel_relay"), None
        )
        has_official_spec = spec_sheet_name is not None
        if has_official_spec:
            rows = ExcelInspector._read_excel_relay_rows_xls(
                wb.sheet_by_name(spec_sheet_name), path, sheet_names
            )
            profile, inputs = ExcelInspector._build_profile_from_excel_relay_rows(
                rows, path, sheet_names
            )
        else:
            profile = ExcelInspector._generate_fallback_profile(sheet_names, path)
            inputs = []

        # 2. 收集所有数据工作表的摘要信息与金融插件特征
        sheets_summary, detected_vendors = ExcelInspector._inspect_sheets_xls(wb)
        has_live_plugin = len(detected_vendors) > 0
        vendor_str = ",".join(sorted(detected_vendors)) if has_live_plugin else "none"

        # 3. 读取 inputs 的当前基准值
        if inputs:
            ExcelInspector._populate_input_values_xls(wb, inputs)

        physical = PhysicalInfo(
            file_size_bytes=stat.st_size,
            file_mtime=str(stat.st_mtime),
            vendor=vendor_str,
            has_live_plugins=has_live_plugin,
            has_official_spec=has_official_spec,
            sheets=sheets_summary,
        )
        return InspectionResult(
            file_path=str(path),
            physical=physical,
            profile=profile,
            inputs=inputs,
        )

    @staticmethod
    def _read_excel_relay_rows_xls(
        ws: xlrd.sheet.Sheet,
        file_path: Path,
        sheet_names: List[str],
    ) -> List[Tuple[str, str, str]]:
        """读取 .xls 的 _excel_relay 表 A/B/C 三列文本 (默认纯离线，仅在必要时按需降级 COM).

        xlrd 无法读取公式文本：若用户在【二、可调输入参数】中以 =假设表!C4 鼠标点选坐标，
        xlrd 只能拿到求值结果，此时坐标校验会失败。仅当 inputs 区块存在无法离线识别的非空坐标时，
        才拉起隐藏 Excel COM 读取原始公式文本。
        """
        rows: List[Tuple[str, str, str]] = []
        section = "meta"
        needs_com_formula_fallback = False

        for r in range(ws.nrows):
            vals = []
            for c in range(3):
                if c < ws.ncols:
                    ctype = ws.cell_type(r, c)
                    v = ws.cell_value(r, c)
                    if ctype == xlrd.XL_CELL_ERROR:
                        vals.append("#ERROR!")
                    elif isinstance(v, float) and v.is_integer():
                        vals.append(str(int(v)))
                    else:
                        vals.append(str(v).strip() if v is not None else "")
                else:
                    vals.append("")

            c0, c1, c2 = vals[0], vals[1], vals[2]
            if any(k in c0 for k in ["【一", "一、模型基本信息"]):
                section = "meta"
            elif any(k in c0 for k in ["【二", "二、可调输入参数", "可调输入参数"]):
                section = "inputs"
            elif section == "inputs" and c1:
                if not any(h in c0 or h in c1 for h in ["参数名称", "所在单元格", "【"]):
                    if qualify_and_validate_ref(c1, sheet_names) is None:
                        needs_com_formula_fallback = True

            rows.append((c0, c1, c2))

        if needs_com_formula_fallback:
            formula_rows = ExcelInspector._read_excel_relay_formulas_xls_com(file_path, ws.nrows)
            if formula_rows:
                rows = formula_rows

        return rows

    @staticmethod
    def _read_excel_relay_formulas_xls_com(file_path: Path, nrows: int) -> Optional[List[Tuple[str, str, str]]]:
        """当 .xls 文件的 _excel_relay 表包含 =Sheet!Cell 公式引用时，通过隐藏 COM 读取原始公式文本."""
        if sys.platform != "win32":
            return None

        from excel_relay.core.refresher import ExcelComSession

        try:
            with ExcelComSession(
                visible=False,
                display_alerts=False,
                screen_updating=False,
                ask_to_update_links=False,
            ) as session:
                wb = session.open_workbook(file_path, update_links=0, read_only=True)
                # 大小写不敏感查找规范表名称
                target_sheet_name = "_excel_relay"
                for sh_idx in range(1, wb.Worksheets.Count + 1):
                    if wb.Worksheets(sh_idx).Name.strip().lower() == "_excel_relay":
                        target_sheet_name = wb.Worksheets(sh_idx).Name
                        break
                ws = wb.Worksheets(target_sheet_name)
                max_r = max(nrows, 1)
                rng_formulas = ws.Range(ws.Cells(1, 1), ws.Cells(max_r, 3)).Formula
                if not isinstance(rng_formulas, tuple):
                    return None

                rows: List[Tuple[str, str, str]] = []
                for r_tuple in rng_formulas:
                    if not isinstance(r_tuple, tuple):
                        continue
                    c0 = str(r_tuple[0]).strip() if len(r_tuple) > 0 and r_tuple[0] is not None else ""
                    c1 = str(r_tuple[1]).strip() if len(r_tuple) > 1 and r_tuple[1] is not None else ""
                    c2 = str(r_tuple[2]).strip() if len(r_tuple) > 2 and r_tuple[2] is not None else ""
                    rows.append((c0, c1, c2))
                return rows
        except Exception as e:
            logger.debug("通过 COM 读取 .xls 公式失败 (%s): %s", file_path, e)
            return None

    @staticmethod
    def _inspect_sheets_xls(wb: xlrd.Book) -> Tuple[List[SheetSummary], set[str]]:
        summaries: List[SheetSummary] = []
        detected_vendors: set[str] = set()

        # 检查二进制工作簿中的外部名称/AddIn 引用特征
        for nobj in getattr(wb, "name_obj_list", []):
            n_lower = getattr(nobj, "name", "").lower()
            if (
                n_lower in ["wsd", "wss", "wsq", "edb", "wset", "wsi", "wst", "tdays", "tdaysoffset", "tdayscount"]
                or "windfunc" in n_lower
            ):
                detected_vendors.add("wind")
            elif n_lower.startswith("ths") or "ifind" in n_lower:
                detected_vendors.add("ifind")

        for name in wb.sheet_names():
            if name.strip().lower() == "_excel_relay":
                continue
            ws = wb.sheet_by_name(name)

            scan_limit = min(ws.nrows, 40)
            sampled_rows: List[List[Any]] = []
            sheet_has_formula = len(detected_vendors) > 0

            for r in range(scan_limit):
                row_vals = []
                for c in range(min(ws.ncols, 30)):
                    val = ws.cell_value(r, c)
                    # 识别 xlrd 错误单元格类型 (如 #NAME?, #N/A)，将其映射为标准错误文本
                    if hasattr(ws, "cell_type") and ws.cell_type(r, c) == xlrd.XL_CELL_ERROR:
                        val = xlrd.biffh.error_text_from_code.get(val, "#ERROR!")
                    row_vals.append(val)
                    s_val = str(val)
                    if (
                        s_val.startswith("=")
                        or "#NAME?" in s_val
                        or any(w in s_val.lower() for w in ["fetching", "retrieving", "正在检索", "正在获取"])
                    ):
                        sheet_has_formula = True
                        for pat in WIND_FORMULA_PATTERNS:
                            if pat.search(s_val):
                                detected_vendors.add("wind")
                        for pat in IFIND_FORMULA_PATTERNS:
                            if pat.search(s_val):
                                detected_vendors.add("ifind")
                sampled_rows.append(row_vals)

            header_row_idx, columns, row_headers = ExcelInspector._extract_headers_from_grid(sampled_rows)

            summaries.append(
                SheetSummary(
                    name=name,
                    row_count=ws.nrows,
                    col_count=ws.ncols,
                    header_row=header_row_idx,
                    columns=columns,
                    row_headers=row_headers,
                    has_formulas=sheet_has_formula,
                )
            )

        return summaries, detected_vendors

    @staticmethod
    def _populate_input_values_xls(
        wb: xlrd.Book,
        inputs: List[ParameterDef],
    ) -> None:
        for inp in inputs:
            inp.current_value = ExcelInspector._read_ref_xls(wb, inp.cell)

    @staticmethod
    def _read_ref_xls(wb: xlrd.Book, ref: str) -> Any:
        clean = normalize_cell_ref(ref)
        if not clean:
            return None

        sheet_name = None
        addr = clean
        if "!" in clean:
            sheet_name, addr = clean.rsplit("!", 1)

        all_sheets = wb.sheet_names()
        if sheet_name:
            # Excel 官方工作表大小写无关 (Case-Insensitive)，通过小写比对准确定位工作表对象
            matched_ws = None
            target_sh_lower = sheet_name.strip().lower()
            for s in all_sheets:
                if s.strip().lower() == target_sh_lower:
                    matched_ws = wb.sheet_by_name(s)
                    break
            if matched_ws is None:
                return None
            ws = matched_ws
        else:
            data_sheets = [s for s in all_sheets if s.strip().lower() != "_excel_relay"]
            ws = wb.sheet_by_name(data_sheets[0] if data_sheets else all_sheets[0])

        try:
            r, c = coordinate_to_tuple(addr)
            if 0 <= r - 1 < ws.nrows and 0 <= c - 1 < ws.ncols:
                # 准确识别 xlrd 错误单元格类型，将其映射为标准的错误字符串 (如 '#N/A', '#DIV/0!')，防止其被误作为纯数字提取
                if hasattr(ws, "cell_type") and ws.cell_type(r - 1, c - 1) == xlrd.XL_CELL_ERROR:
                    err_code = ws.cell_value(r - 1, c - 1)
                    return xlrd.biffh.error_text_from_code.get(err_code, "#ERROR!")
                return clean_scalar_value(ws.cell_value(r - 1, c - 1))
            return None
        except Exception:
            return None

    # -------------------------------------------------------------------------
    # 智能表头与纵向指标提取 & _excel_relay 解析核心
    # -------------------------------------------------------------------------
    @staticmethod
    def _extract_headers_from_grid(
        grid: List[List[Any]],
    ) -> Tuple[int, List[str], List[str]]:
        """从采样网格中智能定位真实横向表头行 (columns) 并提取首列/次列纵向指标名 (row_headers)."""
        if not grid:
            return 1, [], []

        best_row_idx = 0
        best_score = -1

        # 在前 12 行中寻找“短文本字段数量最多”的行作为真实表头行
        for r_idx in range(min(len(grid), 12)):
            row = grid[r_idx]
            text_count = 0
            non_empty_count = 0
            for val in row:
                if val is None:
                    continue
                # 浮点整数归整 (如 2024.0 -> '2024')，避免财模年份被识别为浮点字符串
                if isinstance(val, float) and val.is_integer():
                    s = str(int(val))
                else:
                    s = str(val).strip()
                if not s or s.startswith("="):
                    continue
                non_empty_count += 1
                # 排除长篇说明段落，偏好短字段名或年份表头 (如 2025E, 2025)
                if len(s) <= 25:
                    text_count += 1
            score = text_count * 2 + non_empty_count
            if non_empty_count >= 2 and score > best_score:
                best_score = score
                best_row_idx = r_idx

        columns: List[str] = []
        for val in grid[best_row_idx]:
            if val is not None:
                # 浮点整数归整 (如 2024.0 -> '2024')，确保财模年份表头字段美观且精准匹配
                if isinstance(val, float) and val.is_integer():
                    s = str(int(val))
                else:
                    s = str(val).strip()
                if s and not s.startswith("=") and len(s) <= 35:
                    columns.append(s)
        columns = columns[:25]

        # 提取表头行下方的首列/次列纵向指标名 (适配三大财务报表与估值表)
        row_headers: List[str] = []
        seen_rh: set[str] = set(columns)
        for r_idx in range(best_row_idx + 1, len(grid)):
            row = grid[r_idx]
            if not row:
                continue
            candidate = None
            for c_idx in range(min(len(row), 2)):
                val = row[c_idx]
                if val is None or isinstance(val, (int, float, bool)):
                    continue
                s = str(val).strip()
                # 过滤纯数字序号、公式、空串或过长句子
                if not s or s.startswith("=") or len(s) > 28:
                    continue
                if re.match(r"^[\d\.\-\/\s%]+$", s):
                    continue
                candidate = s
                break
            if candidate and candidate not in seen_rh:
                seen_rh.add(candidate)
                row_headers.append(candidate)
                if len(row_headers) >= 25:
                    break

        return best_row_idx + 1, columns, row_headers

    @staticmethod
    def _build_profile_from_excel_relay_rows(
        rows: List[Tuple[str, str, str]],
        file_path: Path,
        sheet_names: List[str],
    ) -> Tuple[PackageProfile, List[ParameterDef]]:
        """解析全中文 _excel_relay 表 (严格对齐 A/B/C 三列结构)."""
        meta: Dict[str, str] = {}
        declared_inputs: List[ParameterDef] = []

        section = "meta"  # "meta" | "inputs"

        META_KEYS = {
            "模型名称": "name",
            "功能摘要": "summary",
            "检索标签": "tags",
            "常见问题": "questions",
            "口径备注": "notes",
        }

        for c0, c1, c2 in rows:
            if not c0 and not c1:
                continue

            # 1. 识别区块切换标题
            if any(k in c0 for k in ["【一", "一、模型基本信息"]):
                section = "meta"
                continue
            if any(k in c0 for k in ["【二", "二、可调输入参数", "可调输入参数"]):
                section = "inputs"
                continue

            # 2. 按当前所在区块解析
            if section == "meta":
                mapped_key = META_KEYS.get(c0)
                if mapped_key:
                    meta[mapped_key] = c1
            elif section == "inputs":
                # 跳过子表头行或未填写B列坐标的空行
                if "参数名称" in c0 or "所在单元格" in c1 or not c1:
                    continue

                valid_cell = qualify_and_validate_ref(c1, sheet_names)
                if valid_cell:
                    declared_inputs.append(
                        ParameterDef(
                            cell=valid_cell,
                            name=c0 or valid_cell,
                            description=c2 or None,
                        )
                    )

        # 3. 按 cell/name 去重 (防止作者在 _excel_relay 表中重复声明同一参数)
        merged_inputs = ExcelInspector._dedupe_inputs(declared_inputs)

        # 4. 判定【一、模型基本信息】是否被作者实质填写
        # 若作者仅注入了模板用于填写【二、可调输入参数】，而未填写摘要、标签、问题、备注且未修改默认模型名称，
        # 则回退使用物理结构提取的兜底 Profile (状态为 RAW)，以便后续允许 Agent 自主赋能 (ENRICHED)。
        has_declared_meta = bool(
            meta.get("summary")
            or meta.get("tags")
            or meta.get("questions")
            or meta.get("notes")
            or (meta.get("name") and meta.get("name") != file_path.stem)
        )
        fallback_profile = ExcelInspector._generate_fallback_profile(sheet_names, file_path)
        if not has_declared_meta:
            return fallback_profile, merged_inputs

        name = meta.get("name") or file_path.stem
        summary = meta.get("summary") or fallback_profile.summary
        questions_raw = meta.get("questions", "")
        questions = [q.strip() for q in re.split(r"[;；\n]+", questions_raw) if q.strip()]
        tags_raw = meta.get("tags", "")
        tags = [t.strip() for t in re.split(r"[,，\s]+", tags_raw) if t.strip()] or fallback_profile.tags
        notes = meta.get("notes") or None

        profile = PackageProfile(
            name=name,
            summary=summary,
            tags=tags,
            questions=questions,
            notes=notes,
            metadata_status=MetadataStatus.DECLARED,
            description_origin=DescriptionOrigin.AUTHOR_DECLARED,
        )
        return profile, merged_inputs

    @staticmethod
    def _dedupe_inputs(items: List[ParameterDef]) -> List[ParameterDef]:
        """按单元格坐标与参数名去重 (保留首次声明)."""
        result: List[ParameterDef] = []
        seen_cells: set[str] = set()
        seen_names: set[str] = set()

        for item in items:
            c_key = item.cell.lower()
            n_key = item.name.lower()
            if c_key in seen_cells or n_key in seen_names:
                continue
            seen_cells.add(c_key)
            seen_names.add(n_key)
            result.append(item)
        return result

    @staticmethod
    def _generate_fallback_profile(
        sheet_names: List[str],
        file_path: Path,
    ) -> PackageProfile:
        """对未包含 _excel_relay 表的普通文件生成保守兜底语义卡片."""
        stem = file_path.stem
        tags = [w for w in re.split(r"[_\-\s\(\)（）]+", stem) if len(w) >= 2]
        data_sheets = [s for s in sheet_names if s.strip().lower() != "_excel_relay"]
        summary = f"{stem} (包含工作表: {', '.join(data_sheets[:4])})"

        return PackageProfile(
            name=stem,
            summary=summary,
            questions=[f"查看{stem}中的数据", f"刷新{stem}获取最新行情"],
            tags=tags,
            metadata_status=MetadataStatus.RAW,
            description_origin=DescriptionOrigin.EXTRACTED,
        )

    @staticmethod
    def resolve_and_read_inputs(
        file_path: str | Path,
        inputs_raw: Optional[List[str]] = None,
    ) -> Optional[List[ParameterDef]]:
        """解析通过 CLI/MCP 传入的参数坐标声明字符串列表，严格校验工作表存在性并读取当前基准值.

        供 CatalogManager.update_source() 在外部登记可调输入参数时调用。
        若坐标语法非法或目标工作表不存在，直接抛出 ValueError。
        """
        if inputs_raw is None:
            return None

        path = Path(file_path).resolve()
        if not path.is_file():
            raise FileNotFoundError(f"源文件不存在: {path}")

        ext = path.suffix.lower()
        if ext not in [".xls", ".xlsx", ".xlsm"]:
            raise ValueError(f"不支持的文件格式: {path.suffix}")

        # 单次加载工作簿：同时完成工作表存在性校验与当前基准值读取
        if ext == ".xls":
            wb_xls = xlrd.open_workbook(str(path))
            sheet_names = list(wb_xls.sheet_names())
            parsed_inputs = ExcelInspector._parse_inline_inputs(inputs_raw, sheet_names)
            ExcelInspector._populate_input_values_xls(wb_xls, parsed_inputs)
        else:
            wb_val = openpyxl.load_workbook(path, read_only=False, data_only=True)
            try:
                sheet_names = list(wb_val.sheetnames)
                parsed_inputs = ExcelInspector._parse_inline_inputs(inputs_raw, sheet_names)
                for inp in parsed_inputs:
                    inp.current_value = ExcelInspector._read_ref_openpyxl(wb_val, inp.cell)
            finally:
                wb_val.close()
        return parsed_inputs

    @staticmethod
    def _parse_inline_inputs(
        inputs_raw: List[str],
        sheet_names: List[str],
    ) -> List[ParameterDef]:
        """将参数声明字符串解析为 ParameterDef 列表，并严格校验坐标语法与工作表存在性."""
        data_sheets = [s for s in sheet_names if s.strip().lower() != "_excel_relay"]
        parsed_inputs: List[ParameterDef] = []
        for item_str in inputs_raw:
            for p_name, p_cell, p_desc in parse_inline_input_items(item_str):
                valid_cell = qualify_and_validate_ref(p_cell, sheet_names)
                if not valid_cell:
                    raise ValueError(
                        f"无效的输入参数坐标或工作表不存在: '{p_cell}' "
                        f"(仅支持单元格坐标如 '假设表!C4'，不支持命名区域；可用工作表: {data_sheets})"
                    )
                final_name = valid_cell if p_name == p_cell else p_name
                parsed_inputs.append(ParameterDef(cell=valid_cell, name=final_name, description=p_desc))
        return parsed_inputs

    @staticmethod
    def refresh_input_values(
        file_path: str | Path,
        inputs: List[ParameterDef],
    ) -> None:
        """直接刷新已有 ParameterDef 列表的 current_value (单次加载工作簿，按 'Sheet!Cell' 坐标读取).

        内部安全闭环保护：作为轻量只读嗅探探针，遇到底层暂态文件锁或格式异常时静默跳过，绝不崩溃上层调用。
        """
        path = Path(file_path).resolve()
        if not path.is_file() or not inputs:
            return
        ext = path.suffix.lower()
        try:
            if ext == ".xls":
                wb = xlrd.open_workbook(str(path))
                ExcelInspector._populate_input_values_xls(wb, inputs)
            elif ext in [".xlsx", ".xlsm"]:
                ExcelInspector._populate_input_values_openpyxl(path, inputs)
        except Exception:
            pass

    @staticmethod
    def validate_excel_relay_coordinates(file_path: str | Path) -> List[str]:
        """检查 _excel_relay 表中填写的参数坐标是否存在语法错误或指向不存在的工作表."""
        path = Path(file_path).resolve()
        if not path.is_file():
            return [f"文件不存在: {path}"]

        ext = path.suffix.lower()
        if ext == ".xls":
            wb = xlrd.open_workbook(str(path))
            sheet_names = list(wb.sheet_names())
            # 大小写不敏感查找规范表名称
            spec_sheet_name = next(
                (s for s in sheet_names if s.strip().lower() == "_excel_relay"), None
            )
            if not spec_sheet_name:
                return []
            rows = ExcelInspector._read_excel_relay_rows_xls(wb.sheet_by_name(spec_sheet_name), path, sheet_names)
        else:
            wb_f = openpyxl.load_workbook(path, read_only=False, data_only=False)
            try:
                sheet_names = list(wb_f.sheetnames)
                # 大小写不敏感查找规范表名称
                spec_sheet_name = next(
                    (s for s in sheet_names if s.strip().lower() == "_excel_relay"), None
                )
                if not spec_sheet_name:
                    return []
                rows = ExcelInspector._read_excel_relay_rows_openpyxl(wb_f[spec_sheet_name])
            finally:
                wb_f.close()

        errors: List[str] = []
        section = "meta"
        data_sheets = [s for s in sheet_names if s.strip().lower() != "_excel_relay"]

        for r_idx, (c0, c1, _c2) in enumerate(rows, start=1):
            if not c0 and not c1:
                continue
            if any(k in c0 for k in ["【一", "一、模型基本信息"]):
                section = "meta"
                continue
            if any(k in c0 for k in ["【二", "二、可调输入参数", "可调输入参数"]):
                section = "inputs"
                continue

            if section == "inputs":
                if "参数名称" in c0 or "所在单元格" in c1:
                    continue
                if not c1 or not qualify_and_validate_ref(c1, sheet_names):
                    errors.append(
                        f"_excel_relay 第 {r_idx} 行输入参数坐标无效或工作表不存在: '{c1}' "
                        f"(仅支持单元格坐标如 '假设表!C4'，不支持命名区域；可用工作表: {data_sheets})"
                    )

        return errors

    @staticmethod
    def validate_package(file_path: str | Path) -> Tuple[InspectionResult, List[str], List[str]]:
        """完整校验 Excel 数据源包的 _excel_relay 规范表、接口坐标与元数据质量.

        Returns:
            Tuple[InspectionResult, errors列表, warnings列表]
        """
        inspected = ExcelInspector.inspect_file(file_path)
        p = inspected.profile
        errors: List[str] = []
        warnings: List[str] = []

        if not inspected.physical.has_official_spec:
            errors.append(
                "缺少 _excel_relay 规范工作表 (可右键选择 '[ExcelRelay] 注入“_excel_relay”元数据模板' 或运行 excel-relay pack init 生成)"
            )
        else:
            coord_errors = ExcelInspector.validate_excel_relay_coordinates(file_path)
            errors.extend(coord_errors)
            if p.metadata_status == MetadataStatus.RAW:
                warnings.append("【一、模型基本信息】尚未填写，当前将回退使用物理表头兜底摘要（或由 AI 后续自动补全）")
            else:
                if len(p.summary) < 8:
                    warnings.append("功能摘要 (summary) 过短，不利于 AI 精准路由，建议补充完整一句话业务说明")
                if not p.questions:
                    warnings.append("未填写常见问题 (questions)，建议补充 1~2 个典型提问示例")

        return inspected, errors, warnings

    # -------------------------------------------------------------------------
    # 两段式可视化规范表注入 (inject_spec_sheet)
    # -------------------------------------------------------------------------
    @staticmethod
    def _build_template_rows(stem: str) -> List[Tuple[str, str, str, str]]:
        """生成全中文 _excel_relay 表单行定义: (row_type, colA, colB, colC).

        row_type: 'section' | 'subheader' | 'meta_item' | 'io_item' | 'blank'
        """
        return [
            ("section", "【一、模型基本信息】", "填写内容 (浅黄色单元格为填写区，选填)", "填写说明与示例 (灰色提示，无需修改)"),
            ("meta_item", "模型名称", stem, "默认取文件名，可修改为正式中文全称"),
            ("meta_item", "功能摘要", "", "★推荐：一句话说明表的功能，供AI搜索定位（留空则由AI或物理结构自动提取）"),
            ("meta_item", "检索标签", "", "多个标签用逗号分隔，如：白酒, 茅台, 估值"),
            ("meta_item", "常见问题", "", "典型提问方式，多个问题用分号 ; 分隔"),
            ("meta_item", "口径备注", "", "测算口径说明、数据来源或作者署名 (选填)"),
            ("blank", "", "", ""),
            ("section", "【二、可调输入参数】", "单元格位置 (可输 = 用鼠标点选单个格，无参数可留空)", "单位与范围说明 (选填)"),
            ("subheader", "参数名称 (选填)", "所在单元格 (必填，如 假设表!C4)", "补充说明 (如：单位%，合理范围 0~15)"),
            ("io_item", "", "", ""),
            ("io_item", "", "", ""),
            ("io_item", "", "", ""),
        ]

    @staticmethod
    def inject_spec_sheet(file_path: str | Path, hidden: bool = False) -> bool:
        """向目标 Excel 文件注入三段式可视化 _excel_relay 模板表.

        Args:
            file_path: 目标 Excel 文件路径
            hidden: 是否将注入的 _excel_relay 工作表设为隐藏。默认为 False (可见并置于首位，方便用户立即填写)。
        """
        path = Path(file_path).resolve()
        if not path.is_file():
            raise FileNotFoundError(f"文件不存在: {path}")

        ext = path.suffix.lower()
        if ext not in [".xlsx", ".xlsm", ".xls"]:
            raise ValueError(f"不支持的文件格式: {path.suffix} (仅支持 .xlsx, .xlsm, .xls)")

        # 在 Windows 平台下，优先通过 Excel COM 原生无损注入
        # 彻底规避 openpyxl 重新序列化导致的图表/绘图部件损坏 (如 drawing1.xml 丢失) 及条件格式丢失
        if sys.platform == "win32":
            try:
                return ExcelInspector._inject_spec_sheet_com(path, hidden=hidden)
            except Exception as e:
                logger.warning(
                    "尝试通过 Windows Excel COM 注入 _excel_relay 失败 (%s: %s)，正在降级为 openpyxl 模式处理...",
                    path.name,
                    e,
                )
                if ext == ".xls":
                    raise

        if ext in [".xlsx", ".xlsm"]:
            return ExcelInspector._inject_spec_sheet_openpyxl(path, hidden=hidden)
        else:
            return ExcelInspector._inject_spec_sheet_com(path, hidden=hidden)

    @staticmethod
    def fix_external_links_consistency(file_path: str | Path) -> bool:
        """修复 OOXML (.xlsx / .xlsm) 中 externalLinks 与其 _rels 关系 ID 不一致的问题.

        openpyxl 在保存包含外部数据源引用的工作簿 (如 WindFunc.xla / iFinD / 外部模型公式) 时，
        极易出现 xl/externalLinks/externalLinkN.xml 中的 <externalBook r:id="..."> 与
        xl/externalLinks/_rels/externalLinkN.xml.rels 中的 <Relationship Id="..."> 不一致的缺陷
        (例如 externalBook 指向 rId1，而 _rels 中被写为 rId3)，导致 Excel 打开时弹窗提示：
        'Repaired Records: External formula reference from /xl/externalLinks/externalLink1.xml part (Cached values from external formula reference)'。

        此函数进行底层 ZIP/XML 级原子对齐校验与静默自愈。
        """
        path = Path(file_path).resolve()
        if not path.is_file() or path.suffix.lower() not in [".xlsx", ".xlsm"]:
            return False

        temp_dir = tempfile.mkdtemp()
        temp_zip = Path(temp_dir) / f"repaired_{path.name}"
        fixed_any = False

        try:
            with zipfile.ZipFile(path, "r") as zin:
                namelist = zin.namelist()
                ext_xmls = [
                    n for n in namelist
                    if re.match(r"^xl/externallinks/externallink\d+\.xml$", n, re.IGNORECASE)
                ]
                if not ext_xmls:
                    return False

                updates = {}
                for ext_xml_name in ext_xmls:
                    basename = os.path.basename(ext_xml_name)
                    expected_rels = f"xl/externalLinks/_rels/{basename}.rels"
                    matching_rels = [n for n in namelist if n.lower() == expected_rels.lower()]
                    if not matching_rels:
                        continue
                    rels_name = matching_rels[0]

                    try:
                        xml_data = zin.read(ext_xml_name)
                        root = ET.fromstring(xml_data)
                        ext_book = root.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}externalBook")
                        if ext_book is None:
                            ext_book = root.find("externalBook")
                        if ext_book is None:
                            continue

                        rel_id = ext_book.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                        if not rel_id:
                            continue

                        rels_data = zin.read(rels_name)
                        rels_root = ET.fromstring(rels_data)
                        need_update = False

                        all_rels = rels_root.findall("{http://schemas.openxmlformats.org/package/2006/relationships}Relationship")
                        existing_ids = {r.attrib.get("Id") for r in all_rels}

                        # 1. 确保 Relationship Id 与 externalBook r:id 强一致 (杜绝 Excel 报外链缓存损坏)
                        # 仅在 rels 中完全缺失 rel_id 时才进行对齐修正，杜绝破坏包含 alternateUrls (rId1, rId2, rId3) 的现代 Excel 多关系定义
                        if rel_id not in existing_ids:
                            alt_ids = set()
                            for child in ext_book:
                                if child.tag.endswith("alternateUrls"):
                                    for sub_elem in child:
                                        sub_id = sub_elem.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                                        if sub_id:
                                            alt_ids.add(sub_id)

                            candidates = [r for r in all_rels if r.attrib.get("Id") not in alt_ids]
                            target_rel = candidates[0] if candidates else (all_rels[0] if all_rels else None)
                            if target_rel is not None:
                                target_rel.attrib["Id"] = rel_id
                                need_update = True
                                fixed_any = True

                        # 2. 检查并自愈金融插件路径 Target (防止跨盘符移动时相对路径失效导致公式附加长绝对路径前缀)
                        for rel_elem in all_rels:
                            target = rel_elem.attrib.get("Target", "")
                            target_lower = target.lower()
                            if "windfunc" in target_lower:
                                wind_candidates = [
                                    r"C:\Wind\Wind.NET.Client\WindNET\DataBrowse\XLA\WindFunc.xla",
                                    r"D:\Wind\Wind.NET.Client\WindNET\DataBrowse\XLA\WindFunc.xla",
                                ]
                                for cand in wind_candidates:
                                    if os.path.isfile(cand):
                                        cand_url = f"file:///{cand.replace(os.sep, '/')}"
                                        if target != cand_url:
                                            rel_elem.attrib["Target"] = cand_url
                                            need_update = True
                                            fixed_any = True
                                        break
                            elif any(k in target_lower for k in ["thsxls", "thsslg"]):
                                ifind_candidates = [
                                    r"C:\同花顺\iFinD\bin\thsxls.xla",
                                    r"D:\同花顺\iFinD\bin\thsxls.xla",
                                    r"C:\iFinD\bin\thsxls.xla",
                                ]
                                for cand in ifind_candidates:
                                    if os.path.isfile(cand):
                                        cand_url = f"file:///{cand.replace(os.sep, '/')}"
                                        if target != cand_url:
                                            rel_elem.attrib["Target"] = cand_url
                                            need_update = True
                                            fixed_any = True
                                        break

                        if need_update:
                            updates[rels_name] = ET.tostring(rels_root, encoding="utf-8", xml_declaration=True)
                    except Exception:
                        continue

                if not fixed_any:
                    return False

                with zipfile.ZipFile(temp_zip, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                    for item in zin.infolist():
                        if item.filename in updates:
                            zout.writestr(item.filename, updates[item.filename])
                        else:
                            zout.writestr(item, zin.read(item.filename))

            shutil.move(str(temp_zip), str(path))
            return True
        except Exception:
            return False
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    @staticmethod
    def _inject_spec_sheet_openpyxl(path: Path, hidden: bool = False) -> bool:
        wb = openpyxl.load_workbook(path)
        saved = False
        try:
            # 大小写不敏感排查：若已有规范表 (无论大小写)，则跳过注入，保持原文件完全不变
            if any(s.strip().lower() == "_excel_relay" for s in wb.sheetnames):
                return False

            ws = wb.create_sheet(title="_excel_relay", index=0)
            ws.sheet_state = "hidden" if hidden else "visible"

            # 激活与工作表选中状态治理：确保单表激活，杜绝多表同时选中进入 [工作表组] (Group) 联动覆写模式
            if not hidden:
                wb.active = ws
                if ws.views.sheetView:
                    ws.views.sheetView[0].tabSelected = True
                for other_ws in wb.worksheets:
                    if other_ws != ws:
                        for v in other_ws.views.sheetView:
                            v.tabSelected = None
            else:
                if ws.views.sheetView:
                    ws.views.sheetView[0].tabSelected = None
                first_visible = next((s for s in wb.worksheets if s != ws and s.sheet_state == "visible"), None)
                if first_visible:
                    wb.active = first_visible
                    if first_visible.views.sheetView:
                        first_visible.views.sheetView[0].tabSelected = True

            # 设置列宽
            ws.column_dimensions["A"].width = 24
            ws.column_dimensions["B"].width = 50
            ws.column_dimensions["C"].width = 44

            # 定义专业配色与样式
            fill_section = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            fill_sub = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
            fill_edit = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
            fill_label = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

            font_section = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
            font_sub = Font(name="微软雅黑", size=10, bold=True, color="1F4E78")
            font_label = Font(name="微软雅黑", size=10, bold=True, color="333333")
            font_edit = Font(name="微软雅黑", size=10, bold=False, color="000000")
            font_hint = Font(name="微软雅黑", size=9, italic=False, color="595959")

            thin_side = Side(border_style="thin", color="BFBFBF")
            border_grid = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
            align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

            template_rows = ExcelInspector._build_template_rows(path.stem)
            for r_idx, (r_type, col_a, col_b, col_c) in enumerate(template_rows, start=1):
                ws.cell(row=r_idx, column=1, value=col_a)
                ws.cell(row=r_idx, column=2, value=col_b)
                ws.cell(row=r_idx, column=3, value=col_c)

                if r_type == "blank":
                    ws.row_dimensions[r_idx].height = 12
                    continue

                ws.row_dimensions[r_idx].height = 24 if r_type == "section" else 21
                for c_idx in range(1, 4):
                    cell = ws.cell(row=r_idx, column=c_idx)
                    cell.alignment = align_left
                    cell.border = border_grid

                    if r_type == "section":
                        cell.fill = fill_section
                        cell.font = font_section
                    elif r_type == "subheader":
                        cell.fill = fill_sub
                        cell.font = font_sub
                    elif r_type == "meta_item":
                        if c_idx == 1:
                            cell.fill = fill_label
                            cell.font = font_label
                        elif c_idx == 2:
                            cell.fill = fill_edit
                            cell.font = font_edit
                        else:
                            cell.font = font_hint
                    elif r_type == "io_item":
                        cell.fill = fill_edit
                        cell.font = font_edit

            # 内存级同步 externalLinks 关系 ID，降低底层写出歧义
            if hasattr(wb, "_external_links"):
                for link in wb._external_links:
                    if hasattr(link, "file_link") and hasattr(link, "externalBook") and link.file_link and link.externalBook:
                        if link.externalBook.id:
                            link.file_link.Id = link.externalBook.id
                        elif link.file_link.Id:
                            link.externalBook.id = link.file_link.Id

            wb.save(path)
            saved = True
            return True
        finally:
            wb.close()
            # 仅在确实保存了新工作表时，才对 OOXML 级 externalLinks 进行对齐自愈
            if saved:
                ExcelInspector.fix_external_links_consistency(path)

    @staticmethod
    def _inject_spec_sheet_com(path: Path, hidden: bool = False) -> bool:
        """通过 Windows Excel COM 原生无损注入 _excel_relay 模板表.

        适用于 .xlsx, .xlsm, .xls 等所有格式。由 Excel 官方内核读写，
        彻底杜绝 openpyxl 重写导致的图表损坏 (drawing1.xml 缺失) 及条件格式丢失。
        """
        if sys.platform != "win32":
            raise OSError("写入 Excel 模板表依赖 Windows Excel COM。")

        from excel_relay.core.refresher import ExcelComSession

        try:
            with ExcelComSession(
                visible=False,
                display_alerts=False,
                screen_updating=False,
                ask_to_update_links=False,
            ) as session:
                wb = session.open_workbook(path, update_links=0, read_only=False)
                for i in range(1, wb.Worksheets.Count + 1):
                    if wb.Worksheets(i).Name.strip().lower() == "_excel_relay":
                        return False

                ws = wb.Worksheets.Add(Before=wb.Worksheets(1))
                ws.Name = "_excel_relay"

                ws.Columns(1).ColumnWidth = 24
                ws.Columns(2).ColumnWidth = 50
                ws.Columns(3).ColumnWidth = 44

                template_rows = ExcelInspector._build_template_rows(path.stem)

                # 全表基础字体统一设置为微软雅黑
                try:
                    ws.Range(ws.Cells(1, 1), ws.Cells(len(template_rows), 3)).Font.Name = "微软雅黑"
                except Exception:
                    pass

                for r_idx, (r_type, col_a, col_b, col_c) in enumerate(template_rows, start=1):
                    ws.Cells(r_idx, 1).Value = col_a
                    ws.Cells(r_idx, 2).Value = col_b
                    ws.Cells(r_idx, 3).Value = col_c

                    if r_type == "blank":
                        continue

                    row_range = ws.Range(ws.Cells(r_idx, 1), ws.Cells(r_idx, 3))
                    row_range.Borders.LineStyle = 1  # xlContinuous

                    if r_type == "section":
                        row_range.Interior.Color = 0x784E1F  # BGR for #1F4E78
                        row_range.Font.Color = 0xFFFFFF
                        row_range.Font.Bold = True
                    elif r_type == "subheader":
                        row_range.Interior.Color = 0xF2E1D9  # BGR for #D9E1F2
                        row_range.Font.Bold = True
                    elif r_type == "meta_item":
                        ws.Cells(r_idx, 1).Interior.Color = 0xF2F2F2
                        ws.Cells(r_idx, 1).Font.Bold = True
                        ws.Cells(r_idx, 2).Interior.Color = 0xCCF2FF  # BGR for #FFF2CC
                        ws.Cells(r_idx, 3).Font.Color = 0x595959
                    elif r_type == "io_item":
                        row_range.Interior.Color = 0xCCF2FF  # BGR for #FFF2CC

                if hidden:
                    ws.Visible = 0  # 0 = xlSheetHidden
                else:
                    ws.Activate()
                wb.Save()
                return True
        except Exception as e:
            logger.debug("向 Excel 注入 _excel_relay 模板表失败 (%s): %s", path, e)
            raise

    # 保持向后兼容的静态方法别名
    _inject_spec_sheet_xls = _inject_spec_sheet_com
