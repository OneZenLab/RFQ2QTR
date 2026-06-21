import re
import argparse
from pathlib import Path
from typing import Any, Dict, List, Tuple

import pandas as pd
import yaml


BASE_DIR = Path(__file__).resolve().parent
LEXER_PATH = BASE_DIR / "Lexer.yaml"


def load_lexer() -> Dict[str, Any]:
    with open(LEXER_PATH, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f)
    return data.get("lexer", {})


def compile_enum_patterns(enum_conf: Dict[str, Any]) -> Dict[str, List[Tuple[str, re.Pattern]]]:
    """Compile regex patterns for enums like UNI_TYPE, UNI_OD, etc."""
    compiled: Dict[str, List[Tuple[str, re.Pattern]]] = {}
    for enum_name, conf in enum_conf.items():
        mappings = conf.get("mappings", {})
        bucket: List[Tuple[str, re.Pattern]] = []
        for value_name, pat_conf in mappings.items():
            pattern = pat_conf.get("pattern", "")
            flags = pat_conf.get("flags", [])
            flag_val = 0
            for fl in flags:
                if fl.upper() == "IGNORECASE":
                    flag_val |= re.IGNORECASE
            if not pattern:
                continue
            bucket.append((value_name, re.compile(pattern, flag_val)))
        # For UNI_MATERIAL, use the exact order from Lexer.yaml
        if enum_name == "UNI_MATERIAL":
            # Get the exact order from the YAML file
            yaml_order = list(mappings.keys())
            # Create a priority dictionary based on YAML order
            priority_dict = {value: i for i, value in enumerate(yaml_order)}
            # Sort bucket based on YAML order, with unknown values at the end
            bucket.sort(key=lambda x: priority_dict.get(x[0], len(yaml_order)))
            print(f"Debug: UNI_MATERIAL patterns order: {[value for value, _ in bucket]}")
        # Debug: Print the order of patterns for UNI_ANGLE
        elif enum_name == "UNI_ANGLE":
            print(f"Debug: UNI_ANGLE patterns order: {[value for value, _ in bucket]}")
        compiled[enum_name] = bucket
    return compiled


# 尺寸范围 "N*M" / "NxM" 中第二段分数/整数的标准化（用于输出 UNI_OD 值名）
OD_FRACTION_NORMALIZE = {
    "1 1/2": "1-1/2", "1 1/4": "1-1/4", "2 1/2": "2-1/2", "3 1/2": "3-1/2",
    "1-1/2": "1-1/2", "1-1/4": "1-1/4", "2-1/2": "2-1/2", "3-1/2": "3-1/2",
}


def fraction_to_decimal(fraction_str: str) -> float | None:
    """将分数字符串转换为小数，例如 "1-1/2" -> 1.5，"1 1/2" -> 1.5，"1/2" -> 0.5"""
    if not fraction_str:
        return None
    
    # 处理带整数部分的分数，如 "1-1/2" 或 "1 1/2"
    if "-" in fraction_str:
        parts = fraction_str.split("-")
        if len(parts) == 2:
            try:
                integer_part = float(parts[0])
                frac_part = parts[1]
                if "/" in frac_part:
                    numerator, denominator = map(float, frac_part.split("/"))
                    if denominator != 0:
                        return integer_part + (numerator / denominator)
            except (ValueError, ZeroDivisionError):
                pass
    elif " " in fraction_str and "/" in fraction_str:
        # 处理空格分隔的分数，如 "1 1/2"
        parts = fraction_str.split(" ")
        if len(parts) == 2:
            try:
                integer_part = float(parts[0])
                frac_part = parts[1]
                if "/" in frac_part:
                    numerator, denominator = map(float, frac_part.split("/"))
                    if denominator != 0:
                        return integer_part + (numerator / denominator)
            except (ValueError, ZeroDivisionError):
                pass
    # 处理纯分数，如 "1/2"
    elif "/" in fraction_str:
        try:
            numerator, denominator = map(float, fraction_str.split("/"))
            if denominator != 0:
                return numerator / denominator
        except (ValueError, ZeroDivisionError):
            pass
    # 处理整数，如 "4"
    else:
        try:
            return float(fraction_str)
        except ValueError:
            pass
    
    return None


def normalize_od_for_comparison(od_value: str) -> float | str:
    """归一化 OD 值用于对比，将分数转换为小数"""
    decimal_val = fraction_to_decimal(od_value)
    if decimal_val is not None:
        return decimal_val
    return od_value


def parse_od_range(size_text: str) -> List[str] | None:
    """
    从 Size 文本中解析 "N*M" / "NxM" 形式的两段尺寸，返回 [val1, val2]（标准 OD 值名），
    无法解析时返回 None。支持 4X1、4*1、4 * 1 1/2" 等。
    """
    if not size_text or not str(size_text).strip():
        return None
    text = str(size_text).strip()
    # 支持：纯数字*纯数字（4X1、4*1）、或 数字 * 数字 分数（4 * 1 1/2）；第二段用贪婪匹配以免只拿到 "1"
    m = re.search(r"\b(\d+)\s*[\*Xx×]\s*([\d\s\-/]+)(?:\s*[\"']|NB|DN|\s*$|\s*\))?", text, re.IGNORECASE)
    if not m:
        return None
    first = m.group(1).strip()
    second_raw = m.group(2).strip()
    # 分数形式 "1 1/2" -> "1-1/2"
    if re.match(r"^\d+\s+\d+/\d+$", second_raw):
        second = second_raw.replace(" ", "-")
    else:
        second = OD_FRACTION_NORMALIZE.get(second_raw, second_raw)
    return [first, second]


def compile_column_mapping(rules_conf: Dict[str, Any]) -> List[Tuple[re.Pattern, str]]:
    """Compile column name mapping rules from Lexer.yaml."""
    col_map = rules_conf.get("column_map", {})
    mappings = col_map.get("mappings", [])
    compiled: List[Tuple[re.Pattern, str]] = []
    for m in mappings:
        src_pat = m.get("source", "")
        target = m.get("target", "")
        flags = m.get("flags", [])
        flag_val = 0
        for fl in flags:
            if fl.upper() == "IGNORECASE":
                flag_val |= re.IGNORECASE
        if src_pat and target:
            compiled.append((re.compile(src_pat, flag_val), target))
    return compiled


def detect_header_row(df: pd.DataFrame, lexer_rules: Dict[str, Any]) -> int:
    """Detect the header row index heuristically using keywords from lexer."""
    attr_keywords = lexer_rules.get("attr_table_keywords", {}).get("keywords", [])
    kw_upper = [str(k).upper() for k in attr_keywords]

    def score_row(values: List[Any]) -> int:
        score = 0
        non_empty = 0
        for v in values:
            text = str(v).strip()
            if not text:
                continue
            non_empty += 1
            up = text.upper()
            if any(k in up for k in kw_upper):
                score += 2
        if non_empty >= 3:
            score += 1
        return score

    best_idx = 0
    best_score = -1
    for i in range(min(len(df), 20)):  # only look at first 20 rows
        row_vals = df.iloc[i].tolist()
        sc = score_row(row_vals)
        if sc > best_score:
            best_score = sc
            best_idx = i
    return best_idx


def infer_column_semantics(columns: List[str], compiled_col_maps: List[Tuple[re.Pattern, str]]) -> Dict[str, str]:
    """Map raw column names to internal field names like QTR_LINE, QTR_QTY, etc."""
    mapping: Dict[str, str] = {}
    for col in columns:
        col_str = "" if col is None else str(col)
        for pat, target in compiled_col_maps:
            if pat.match(col_str):
                mapping[col_str] = target
                break
    return mapping


def is_summary_row(row: pd.Series) -> bool:
    text = " ".join(str(v) for v in row.values)
    up = text.upper()
    if not text.strip():
        return True
    summary_keywords = ["TOTAL", "SUBTOTAL", "合计", "小计", "REMARK", "备注"]
    if any(k in up for k in summary_keywords):
        return True
    return False


def is_parent_row(row: pd.Series, col_semantics: Dict[str, str], enum_patterns: Dict[str, List[Tuple[str, re.Pattern]]]) -> bool:
    """
    判断是否为父行（通常是包含属性信息但没有数量或价格的行）
    父行特征：
    1. 有描述信息
    2. 数量为空（核心原则：没有数量就是父行）
    3. 价格为空
    4. 可能包含完整的产品规格信息
    5. 行号格式通常表示更高层级
    """
    # 1. 基于列语义获取列名
    print(f"【父行】: 检查行是否为父行，行数据: {row.to_dict()}")
    desc_cols = []
    qty_cols = []
    price_cols = []
    line_cols = []
    
    for col_name, semantic in col_semantics.items():
        print(f"【父行】[desc_cols]: 检查列语义，列名: {col_name}, 语义: {semantic}")
        if semantic == "QTR_LINE":
            line_cols.append(col_name)
        elif semantic == "QTR_QTY":
            qty_cols.append(col_name)
        elif semantic == "QTR_PRICE":
            price_cols.append(col_name)
        # 其他列可能包含描述信息【】
        else:
            desc_cols.append(col_name)
    
    # 2. 检查描述信息
    has_desc = False
    desc_text = ""
    # 优先使用 Item 列（如果存在）
    if 'Item' in row.index:
        item_val = row.get('Item', '')
        desc_text = str(item_val).strip()
        has_desc = bool(desc_text)
    else:
        # 尝试其他可能的描述列
        print(f"【父行】: 检查描述列，描述列: {desc_cols}")
        for col in desc_cols:
            if col in row.index:
                val = row.get(col, '')
                text = str(val).strip()
                if text:
                    desc_text = text
                    has_desc = True
                    break
    
    if not has_desc:
        return False
    
    # 3. 检查数量（核心原则：没有数量就是父行）
    has_qty = False
    for col in qty_cols:
        if col in row.index:
            val = row.get(col, '')
            if bool(str(val).strip()):
                has_qty = True
                break
    
    if has_qty:
        return False
    
    # 4. 检查价格
    has_price = False
    for col in price_cols:
        if col in row.index:
            val = row.get(col, '')
            if bool(str(val).strip()):
                has_price = True
                break
    
    if has_price:
        return False
    
    # 5. 行号格式分析
    line_num = ""
    for col in line_cols:
        if col in row.index:
            val = row.get(col, '')
            text = str(val).strip()
            if text:
                line_num = text
                break
    
    # 分析行号格式，父行通常层级较少（如 P3.1 vs P3.1.1）
    if line_num:
        # 计算行号中的层级分隔符数量
        level_count = line_num.count('.')
        # 父行通常层级数较少（≤1），子行层级数较多（≥2）
        if level_count > 1:
            return False
    
    # 6. 描述文本分析
    # 检查是否包含来自 Lexer.yaml 中定义的 UNI_* 类型的值
    upper_desc = desc_text.upper()
    
    # 需要检查的 UNI_* 类型
    uni_types_to_check = [
        "UNI_TYPE",
        "UNI_ANGLE",
        "UNI_RADIUS",
        "UNI_MISC",
        "UNI_DIM_SPEC",
        "UNI_MATERIAL",
        "UNI_CONSTRUCTION",
        "UNI_END_PREPARATION"
    ]
    
    # 检查描述文本是否包含任何来自指定 UNI_* 类型的值
    uni_value_count = 0
    for uni_type in uni_types_to_check:
        if uni_type in enum_patterns:
            for value_name, pat in enum_patterns[uni_type]:
                if pat.search(upper_desc):
                    uni_value_count += 1
                    break
    
    # 如果描述文本中包含的 UNI_* 值较少，可能不是父行
    if uni_value_count < 2:
        return False
    
    # 7. 最终判定
    return True


def compile_ignore_cols_pattern(rules_conf: Dict[str, Any]) -> re.Pattern:
    """Compile pattern for columns that should be ignored when building description."""
    ignore_conf = rules_conf.get("ignore_cols", {})
    pattern = ignore_conf.get("pattern")
    flags_list = ignore_conf.get("flags", [])
    if not pattern:
        return re.compile(r"$a")  # match nothing
    flag_val = 0
    for fl in flags_list:
        if fl.upper() == "IGNORECASE":
            flag_val |= re.IGNORECASE
    return re.compile(pattern, flag_val)


def build_desc_text(row: pd.Series, col_semantics: Dict[str, str], ignore_cols_re: re.Pattern) -> str:
    """Build a description/spec text from non-ignored, non-metadata columns."""
    parts: List[str] = []
    for col_name in row.index:
        name_str = "" if col_name is None else str(col_name)
        # 忽略显式配置为 ignore_cols 的列
        if ignore_cols_re and ignore_cols_re.match(name_str):
            continue
        internal = col_semantics.get(name_str, "")
        # 即使 Item 列被映射到 QTR_LINE，也包含其值
        if internal == "QTR_LINE" and name_str.lower() == "item":
            # 包含 Item 列的值
            val = row.get(col_name, "")
            if not pd.isna(val):
                txt = str(val).strip()
                if txt:
                    parts.append(txt)
            continue
        # 忽略明显不是描述的字段：行号、数量、单价
        if internal in {"QTR_LINE", "QTR_QTY", "QTR_PRICE"}:
            continue
        val = row.get(col_name, "")
        if pd.isna(val):
            continue
        txt = str(val).strip()
        if txt:
            # 如果列名包含 DN，则在值前添加 DN
            if "DN" in name_str.upper():
                print(f"Debug: DN column '{name_str}' found with value: '{txt}' -> formatted as: 'DN{txt}'")
                txt = f"DN{txt}"
            parts.append(txt)
    return " ".join(parts)


def match_first_enum(text: str, enum_patterns: List[Tuple[str, re.Pattern]]) -> str:
    for value_name, pat in enum_patterns:
        match = pat.search(text)
        if match:
            print(f"Debug: Enum match found - Value: '{value_name}', Pattern: {pat}, Match: '{match.group(0)}' in text: '{text}'")
            return value_name
    return ""


def parse_row_to_uni(
    row: pd.Series,
    row_index: int,
    col_semantics: Dict[str, str],
    enum_patterns: Dict[str, List[Tuple[str, re.Pattern]]],
    ignore_cols_re: re.Pattern,
    size_col: str | None,
) -> Dict[str, Any]:
    """Parse a single row into UNI_* fields + QTR_*."""
    # 初始化所有字段为空
    uni: Dict[str, Any] = {
        "UNI_TYPE": "",
        "UNI_ANGLE": "",
        "UNI_RADIUS": "",
        "UNI_MISC": "",
        "UNI_OD1": "",
        "UNI_OD2": "",
        "UNI_OD3": "",
        "UNI_WT1": "",
        "UNI_WT2": "",
        "UNI_WT3": "",
        "UNI_DIM_SPEC": "",
        "UNI_MATERIAL": "",
        "UNI_CONSTRUCTION": "",
        "UNI_END_PREPARATION": "",
        "QTR_LINE": "",
        "QTR_QTY": "",
        "QTR_PRICE": "",
    }

    # 行号
    line_col = next((c for c, t in col_semantics.items() if t == "QTR_LINE"), None)
    if line_col is not None:
        v = row.get(line_col, "")
        uni["QTR_LINE"] = v
    else:
        uni["QTR_LINE"] = (row_index + 1) * 10  # fallback: data row index starting from 1

    # 数量
    qty_col = next((c for c, t in col_semantics.items() if t == "QTR_QTY"), None)
    if qty_col is not None:
        v = row.get(qty_col, "")
        uni["QTR_QTY"] = str(v).strip()

    # 单价
    # 遍历所有映射到 QTR_PRICE 的列，找到第一个有值的列
    price_value = ""
    price_cols = [c for c, t in col_semantics.items() if t == "QTR_PRICE"]
    for col in price_cols:
        v = row.get(col, "")
        if v and not pd.isna(v):
            price_value = v
            break
    uni["QTR_PRICE"] = price_value

    # 处理直接映射到 UNI_WT1 和 UNI_WT2 的列（如 SCH1 和 SCH2）
    # 注意：这里不直接存储原始值，而是在后续的 WT 提取和分配部分处理
    # 这样可以确保最终存储的是解析后的值，而不是原始值
    pass

    # 构建描述文本
    desc_text = build_desc_text(row, col_semantics, ignore_cols_re)
    # 初始设置 UNI_DIM_SPEC 为空
    uni["UNI_DIM_SPEC"] = ""

    upper_desc = desc_text.upper()

    # Size 列文本（优先用于 OD 解析）
    size_text = ""
    if size_col and size_col in row.index:
        size_val = row.get(size_col, "")
        if not pd.isna(size_val):
            size_text = str(size_val)
    upper_size = size_text.upper()

    # 使用枚举规则提取类型、角度、半径、材质等
    if "UNI_TYPE" in enum_patterns:
        uni["UNI_TYPE"] = match_first_enum(upper_desc, enum_patterns["UNI_TYPE"])
    if "UNI_ANGLE" in enum_patterns:
        uni["UNI_ANGLE"] = match_first_enum(upper_desc, enum_patterns["UNI_ANGLE"])
    if "UNI_RADIUS" in enum_patterns:
        uni["UNI_RADIUS"] = match_first_enum(upper_desc, enum_patterns["UNI_RADIUS"])
    if "UNI_MISC" in enum_patterns:
        uni["UNI_MISC"] = match_first_enum(upper_desc, enum_patterns["UNI_MISC"])
    if "UNI_MATERIAL" in enum_patterns:
        uni["UNI_MATERIAL"] = match_first_enum(upper_desc, enum_patterns["UNI_MATERIAL"])
    if "UNI_CONSTRUCTION" in enum_patterns:
        uni["UNI_CONSTRUCTION"] = match_first_enum(upper_desc, enum_patterns["UNI_CONSTRUCTION"])
    if "UNI_END_PREPARATION" in enum_patterns:
        uni["UNI_END_PREPARATION"] = match_first_enum(upper_desc, enum_patterns["UNI_END_PREPARATION"])
    if "UNI_DIM_SPEC" in enum_patterns:
        dim_spec_val = match_first_enum(upper_desc, enum_patterns["UNI_DIM_SPEC"])
        if dim_spec_val:
            uni["UNI_DIM_SPEC"] = dim_spec_val

    # 特例处理：描述中出现“RED TEE”时，UNI_TYPE 应为 TEE，RED 作为 UNI_MISC
    if "RED TEE" in upper_desc:
        if uni.get("UNI_TYPE") == "RED":
            uni["UNI_TYPE"] = "TEE"
        # 确保 RED 记录到 UNI_MISC
        misc_val = str(uni.get("UNI_MISC", "") or "")
        if "RED" not in misc_val:
            uni["UNI_MISC"] = "RED" if not misc_val else misc_val

    # OD 和 WT，先收集列表，再根据规则赋值
    wts: List[str] = []
    # 用于存储 WT 匹配的位置，避免被 OD 提取误识别
    wt_match_ranges = []
    # 先提取 WT 值，避免被 OD 提取误识别
    if "UNI_WT" in enum_patterns:
        # 1. 从描述文本中提取 WT 值
        for value_name, pat in enum_patterns["UNI_WT"]:
            match = pat.search(upper_desc)
            if match:
                wts.append(value_name)
                # 记录 WT 匹配的位置
                wt_match_ranges.append(match.span())
        # 2. 从直接映射的列中提取 WT 值（如 SCH1 和 SCH2）
        for col_name, internal in col_semantics.items():
            if internal in {"UNI_WT1", "UNI_WT2"}:
                wt_val = row.get(col_name, "")
                if not pd.isna(wt_val):
                    txt = str(wt_val).strip()
                    if txt:
                        upper_wt_val = txt.upper()
                        for value_name, pat in enum_patterns["UNI_WT"]:
                            if pat.search(upper_wt_val):
                                wts.append(value_name)
                                break
    wts = list(dict.fromkeys(wts))
    # 权重清洗：例如同时匹配到 XXS 和 XS 时，优先保留 XXS
    # 处理其他重叠情况，保留更具体的规格
    if "XXS" in wts and "XS" in wts:
        wts = [w for w in wts if w != "XS"]
    if "S40S" in wts and "S40" in wts:
        wts = [w for w in wts if w != "S40"]
    if "S80S" in wts and "S80" in wts:
        wts = [w for w in wts if w != "S80"]
    if "S10S" in wts and "S10" in wts:
        wts = [w for w in wts if w != "S10"]
    print(f"Debug: Extracted WTs: {wts}")

    ods: List[str] = []
    search_text = upper_size or upper_desc
    print(f"Debug: Search text for OD extraction: '{search_text}'")
    # 优先用 "N*M" / "NxM" 结构解析（如 4X1、4 * 1 1/2），保证 OD1/OD2 正确
    range_ods = parse_od_range(size_text if size_text else desc_text)
    is_range_pattern = False
    if range_ods:
        print(f"Debug: Range ODs parsed: {range_ods}")
        ods = range_ods
        is_range_pattern = True
    elif "UNI_OD" in enum_patterns:
        # 只在 Size/描述 中解析 OD
        print(f"Debug: Using UNI_OD enum patterns for extraction")
        # 先收集所有匹配项，包括它们在文本中的位置
        all_matches = []
        # 用于存储已匹配的位置范围，避免重复匹配子部分
        matched_ranges = []
        
        def add_match_if_not_overlapped(match, value_name):
            """添加匹配项，确保不与已有的匹配项重叠"""
            start, end = match.span()
            # 检查当前匹配是否完全包含在已有的匹配项中
            for existing_start, existing_end in matched_ranges:
                if start >= existing_start and end <= existing_end:
                    print(f"Debug: Skipping nested match '{value_name}' at ({start}, {end}) inside existing match at ({existing_start}, {existing_end})")
                    return
            # 检查当前匹配是否与 WT 匹配重叠
            for wt_start, wt_end in wt_match_ranges:
                if start < wt_end and end > wt_start:
                    print(f"Debug: Skipping OD match '{value_name}' at ({start}, {end}) overlapping with WT match at ({wt_start}, {wt_end})")
                    return
            # 检查当前匹配是否是小数的一部分（例如 5.5 中的 5）
            # 检查匹配的前一个字符是否是小数点或数字（如果匹配在文本开头则跳过）
            if start > 0:
                prev_char = search_text[start - 1]
                if prev_char == '.' or prev_char.isdigit():
                    print(f"Debug: Skipping OD match '{value_name}' at ({start}, {end}) because it's part of a decimal number")
                    return
            # 检查匹配的后一个字符是否是小数点（如果匹配在文本结尾则跳过）
            if end < len(search_text):
                next_char = search_text[end]
                if next_char == '.':
                    print(f"Debug: Skipping OD match '{value_name}' at ({start}, {end}) because it's part of a decimal number")
                    return
            # 如果当前匹配没有被包含在已有的匹配项中，添加它
            matched_ranges.append((start, end))
            all_matches.append((start, -len(match.group(0)), value_name))
        
        # 先尝试从 search_text 中提取 OD 值
        for value_name, pat in enum_patterns["UNI_OD"]:
            # 查找所有匹配项，而不仅仅是第一个
            for match in pat.finditer(search_text):
                print(f"Debug: Match found for value '{value_name}' with pattern {pat}")
                print(f"Debug: Match group: '{match.group(0)}' at position {match.span()}")
                add_match_if_not_overlapped(match, value_name)
        
        # 如果没有从 search_text 中提取到 OD 值，再尝试从 desc_text 中提取
        if not all_matches:
            print(f"Debug: No OD matches found in search_text, trying desc_text: '{upper_desc}'")
            # 重置匹配范围，避免影响 desc_text 的匹配
            matched_ranges = []
            # 为 desc_text 创建一个新的 add_match 函数，使用 upper_desc 作为搜索文本
            def add_match_if_not_overlapped_desc(match, value_name):
                """添加匹配项，确保不与已有的匹配项重叠（用于 desc_text）"""
                start, end = match.span()
                # 检查当前匹配是否完全包含在已有的匹配项中
                for existing_start, existing_end in matched_ranges:
                    if start >= existing_start and end <= existing_end:
                        print(f"Debug: Skipping nested match '{value_name}' at ({start}, {end}) inside existing match at ({existing_start}, {existing_end})")
                        return
                # 检查当前匹配是否与 WT 匹配重叠
                for wt_start, wt_end in wt_match_ranges:
                    if start < wt_end and end > wt_start:
                        print(f"Debug: Skipping OD match '{value_name}' at ({start}, {end}) overlapping with WT match at ({wt_start}, {wt_end})")
                        return
                # 检查当前匹配是否是小数的一部分（例如 5.5 中的 5）
                # 检查匹配的前一个字符是否是小数点或数字（如果匹配在文本开头则跳过）
                if start > 0:
                    prev_char = upper_desc[start - 1]
                    if prev_char == '.' or prev_char.isdigit():
                        print(f"Debug: Skipping OD match '{value_name}' at ({start}, {end}) because it's part of a decimal number")
                        return
                # 检查匹配的后一个字符是否是小数点（如果匹配在文本结尾则跳过）
                if end < len(upper_desc):
                    next_char = upper_desc[end]
                    if next_char == '.':
                        print(f"Debug: Skipping OD match '{value_name}' at ({start}, {end}) because it's part of a decimal number")
                        return
                # 如果当前匹配没有被包含在已有的匹配项中，添加它
                matched_ranges.append((start, end))
                all_matches.append((start, -len(match.group(0)), value_name))
            
            for value_name, pat in enum_patterns["UNI_OD"]:
                # 查找所有匹配项，而不仅仅是第一个
                for match in pat.finditer(upper_desc):
                    print(f"Debug: Match found for value '{value_name}' with pattern {pat}")
                    print(f"Debug: Match group: '{match.group(0)}' at position {match.span()}")
                    add_match_if_not_overlapped_desc(match, value_name)
        # 按匹配位置升序排序，位置相同的按匹配长度降序排序
        all_matches.sort()
        # 提取值名，去重并保持顺序
        seen = set()
        ods = []
        for _, _, value_name in all_matches:
            if value_name not in seen:
                seen.add(value_name)
                ods.append(value_name)
        print(f"Debug: Extracted ODs: {ods}")

    # 如果还未识别到 WT，尝试根据 “BW 数字 尺寸” 结构推断 S40/S80 等
    if not wts:
        # 例如: "... BW 40 4 ..." 或 "... BW 80 3 ..."
        m = re.search(r"\bBW\s+(\d+)\b", upper_desc)
        if m:
            sched_num = m.group(1)
            if sched_num in {"10", "20", "40", "60", "80", "160"}:
                wts.append(f"S{sched_num}")

    # 应用 OD 分配规则（考虑 TEE、RED）
    uni_type = uni.get("UNI_TYPE", "").upper()
    if ods:
        # 减径管件优先让较大口径作为 OD1
        # 只对范围模式（如 4X1）进行排序，对 DN 值不排序
        if uni_type == "RED" and is_range_pattern:
            try:
                ods_sorted = sorted(ods, key=lambda x: float(str(x).replace("-", ".")), reverse=True)
                ods = ods_sorted
            except Exception:
                pass

        if uni_type == "TEE":
            # 三通特例
            if len(ods) == 1:
                uni["UNI_OD1"] = ods[0]
            else:
                uni["UNI_OD1"] = ods[0]
                if len(ods) >= 2:
                    uni["UNI_OD3"] = ods[1]
        else:
            if len(ods) == 1:
                uni["UNI_OD1"] = ods[0]
            elif len(ods) == 2:
                uni["UNI_OD1"] = ods[0]
                uni["UNI_OD2"] = ods[1]
            else:
                uni["UNI_OD1"] = ods[0]
                uni["UNI_OD2"] = ods[1]
                uni["UNI_OD3"] = ods[2]

    # WT 分配
    if wts:
        # 计算解析出的 OD 数量
        od_count = 0
        for i in range(1, 4):
            if uni.get(f"UNI_OD{i}"):
                od_count += 1
        print(f"Debug: OD count: {od_count}")
        
        if uni_type == "TEE":
            # TEE 类型的特殊处理
            if od_count == 2 and len(wts) >= 2:
                # 解析出 2 个 OD 且 TYPE=TEE，显示 WT1 WT3
                uni["UNI_WT1"] = wts[0]
                uni["UNI_WT3"] = wts[1]
                print(f"Debug: TEE with 2 ODs - Setting WT1={wts[0]}, WT3={wts[1]}")
            else:
                # 其他情况，只显示 WT1
                if len(wts) >= 1:
                    uni["UNI_WT1"] = wts[0]
                    print(f"Debug: TEE with {od_count} ODs - Only setting WT1={wts[0]}")
        else:
            # 非 TEE 类型
            if od_count == 1:
                # 解析出 1 个 OD，只显示 WT1
                if len(wts) >= 1:
                    uni["UNI_WT1"] = wts[0]
                    print(f"Debug: Non-TEE with 1 OD - Only setting WT1={wts[0]}")
            else:
                # 解析出多个 OD，按正常逻辑分配
                if len(wts) >= 1:
                    uni["UNI_WT1"] = wts[0]
                if len(wts) >= 2:
                    # 当 WT1 和 WT2 相同时，只显示 WT1
                    if wts[0] != wts[1]:
                        uni["UNI_WT2"] = wts[1]
                    else:
                        print(f"Debug: WT1 and WT2 are the same: {wts[0]}, only setting WT1")
                if len(wts) >= 3:
                    uni["UNI_WT3"] = wts[2]

    return uni


def build_output_dataframe(parsed_rows: List[Dict[str, Any]], uni_to_code_map: Dict[str, str]) -> pd.DataFrame:
    """Convert internal UNI_* + QTR_* dicts to DataFrame with DESC2CODE-style columns."""
    rows_for_df: List[Dict[str, Any]] = []
    for uni in parsed_rows:
        row_out: Dict[str, Any] = {}
        # UNI_* 映射到中文列名
        for uni_field, cn_name in uni_to_code_map.items():
            val = uni.get(uni_field, "")
            if val is None:
                val = ""
            if val != "":
                val = str(val)
            row_out[cn_name] = val
        # 额外的 QTR 字段（如果没有在 uni_to_code_map 里）
        if "QTR_LINE" not in uni_to_code_map:
            row_out["QTR_LINE"] = uni.get("QTR_LINE", "")
        if "QTR_QTY" not in uni_to_code_map:
            row_out["QTR_QTY"] = uni.get("QTR_QTY", "")
        if "QTR_PRICE" not in uni_to_code_map:
            row_out["QTR_PRICE"] = uni.get("QTR_PRICE", "")
        rows_for_df.append(row_out)
    return pd.DataFrame(rows_for_df)


def main() -> None:
    # 解析命令行参数
    parser = argparse.ArgumentParser(description="Parse Excel file and compare with target template")
    parser.add_argument("excel_file", type=str, help="Excel file path (contains both source and target sheets)")
    args = parser.parse_args()
    
    excel_file = Path(args.excel_file)
    
    # 验证文件存在
    if not excel_file.exists():
        print(f"Error: Excel file not found: {excel_file}")
        return

    lexer_conf = load_lexer()
    rules_conf = lexer_conf.get("rules", {})
    enums_conf = lexer_conf.get("enums", {})
    config_conf = lexer_conf.get("config", {})

    enum_patterns = compile_enum_patterns(enums_conf)
    compiled_col_maps = compile_column_mapping(rules_conf)
    uni_to_code_map = rules_conf.get("uni_to_code_map", {})
    ignore_cols_re = compile_ignore_cols_pattern(rules_conf)

    # 1. 粗读，检测表头行（SRC 是表2，索引为1）
    raw_df = pd.read_excel(excel_file, sheet_name=1, header=None, dtype=str)
    header_row = detect_header_row(raw_df, rules_conf)

    print(f"检测到表头行为: {header_row}")

    # 2. 带表头再读一次
    data_df = pd.read_excel(excel_file, sheet_name=1, header=header_row, dtype=str)
    data_df = data_df.fillna("")

    # 3. 列语义映射
    col_semantics = infer_column_semantics(list(data_df.columns), compiled_col_maps)
    print("原始列名:", list(data_df.columns))
    print("列语义映射:", col_semantics)
    try:
        print("前几行 No./Item 列：")
        print(data_df[["No.", "Item"]].head())
    except Exception:
        pass

    # Size 列（若存在）——用于 OD 解析
    size_col = None
    for c in data_df.columns:
        if isinstance(c, str) and "SIZE" in c.upper():
            size_col = c
            break
    if size_col:
        print(f"检测到 Size 列: {size_col}")

    # 4. 行解析（支持树形结构）
    parsed_rows: List[Dict[str, Any]] = []
    current_parent_attributes: Dict[str, Any] = {}
    
    print(f"【父行】: 开始解析行数据")
    for idx, (_, row) in enumerate(data_df.iterrows()):
        print(f"【父行】: 处理第 {idx} 行，行数据: {row.to_dict()}")
        if is_summary_row(row):
            continue
        
        # 检查是否为父行
        print(f"【父行】: 检查行是否为父行，行数据: {row.to_dict()}")
        if is_parent_row(row, col_semantics, enum_patterns):
            # 直接使用 Item 列的值作为父行描述
            parent_item = row.get('Item', '')
            print(f"Debug: 父行 Item 值: '{parent_item}'")
            # 创建一个临时行，将 Item 列的值复制到一个不会被忽略的列
            temp_row = row.copy()
            # 解析父行，提取属性信息
            parent_uni = parse_row_to_uni(temp_row, idx, col_semantics, enum_patterns, re.compile(r"$a"), size_col)
            # 保存非空的属性
            current_parent_attributes = {}
            for key, value in parent_uni.items():
                if value and key not in {"QTR_LINE", "QTR_QTY", "QTR_PRICE"}:
                    current_parent_attributes[key] = value
            print(f"Debug: 检测到父行，提取属性: {current_parent_attributes}")
            continue
        
        # 粗过滤：数量 + 描述至少有一个不为空
        desc_text = build_desc_text(row, col_semantics, ignore_cols_re)
        has_desc = bool(str(desc_text).strip())
        qty_col = next((c for c, t in col_semantics.items() if t == "QTR_QTY"), None)
        qty_val = row.get(qty_col, "") if qty_col is not None else ""
        has_qty = bool(str(qty_val).strip())
        if not (has_desc or has_qty):
            continue
        
        # 解析子行
        uni_row = parse_row_to_uni(row, idx, col_semantics, enum_patterns, ignore_cols_re, size_col)
        
        # 应用父行属性到子行（仅当子行对应属性为空时）
        for key, value in current_parent_attributes.items():
            if not uni_row.get(key, ""):
                uni_row[key] = value
                print(f"Debug: 应用父行属性 {key}={value} 到子行")
        
        parsed_rows.append(uni_row)

    # 5. 构建输出 DataFrame
    out_df = build_output_dataframe(parsed_rows, uni_to_code_map)
    # 使用时间戳避免文件被占用导致写入失败
    import time
    out_path = BASE_DIR / f"{excel_file.stem}_{int(time.time())}.xlsx"
    out_df.to_excel(out_path, index=False)
    print(f"解析结果已输出到: {out_path}")

    # 6. 与目标模板对比（TARGET 是表1，索引为0）
    try:
        target_df = pd.read_excel(excel_file, sheet_name=0, dtype=str).fillna("")
        # 归一化目标模板中的材质值，确保与解析结果格式一致
        material_col = uni_to_code_map.get("UNI_MATERIAL")
        if material_col and material_col in target_df.columns:
            for i, val in enumerate(target_df[material_col]):
                if val:
                    # 使用与解析相同的逻辑归一化材质值
                    normalized_val = match_first_enum(val, enum_patterns.get("UNI_MATERIAL", []))
                    if normalized_val:
                        target_df.at[i, material_col] = normalized_val
        # 对齐列名和排序，包含所有列，包括行号和价格
        common_cols = [c for c in target_df.columns if c in out_df.columns]
        out_aligned = out_df[common_cols].reset_index(drop=True)
        target_aligned = target_df[common_cols].reset_index(drop=True)
        
        # 归一化 OD 列值，将分数转换为小数进行对比
        od_columns = [col for col in common_cols if col.startswith('OD')]
        if od_columns:
            print(f"\n归一化 OD 列进行对比: {od_columns}")
            # 创建归一化后的数据框用于比较
            out_normalized = out_aligned.copy()
            target_normalized = target_aligned.copy()
            
            # 归一化 UNIXLS 解析结果中的 OD 值
            for col in od_columns:
                for i in range(len(out_normalized)):
                    od_val = out_normalized.at[i, col]
                    normalized_val = normalize_od_for_comparison(str(od_val))
                    out_normalized.at[i, col] = normalized_val
            
            # 归一化目标模板中的 OD 值
            for col in od_columns:
                for i in range(len(target_normalized)):
                    od_val = target_normalized.at[i, col]
                    normalized_val = normalize_od_for_comparison(str(od_val))
                    target_normalized.at[i, col] = normalized_val
            
            # 使用归一化后的值进行对比
            if out_normalized.equals(target_normalized):
                print("解析结果与目标模板在所有公共列上完全一致（包括行号和价格，OD 值已归一化对比）。")
            else:
                print("解析结果与目标模板存在差异（包括行号和价格，OD 值已归一化对比）。")
                print("公共列：", common_cols)
                # 打印前几行对比
                max_rows = min(len(out_aligned), len(target_aligned), 10)
                for i in range(max_rows):
                    if not out_normalized.iloc[i].equals(target_normalized.iloc[i]):
                        print(f"--- 第 {i+1} 行不一致 ---")
                        print("UNIXLS:", dict(out_aligned.iloc[i]))
                        print("TARGET:", dict(target_aligned.iloc[i]))
                        print("UNIXLS (归一化):", dict(out_normalized.iloc[i]))
                        print("TARGET (归一化):", dict(target_normalized.iloc[i]))
        else:
            # 没有 OD 列，直接使用原始值对比
            if out_aligned.equals(target_aligned):
                print("解析结果与目标模板在所有公共列上完全一致（包括行号和价格）。")
            else:
                print("解析结果与目标模板存在差异（包括行号和价格）。")
                print("公共列：", common_cols)
                # 打印前几行对比
                max_rows = min(len(out_aligned), len(target_aligned), 10)
                for i in range(max_rows):
                    if not out_aligned.iloc[i].equals(target_aligned.iloc[i]):
                        print(f"--- 第 {i+1} 行不一致 ---")
                        print("UNIXLS:", dict(out_aligned.iloc[i]))
                        print("TARGET:", dict(target_aligned.iloc[i]))
    except Exception as e:
        print(f"对比目标模板时出错: {e}")


if __name__ == "__main__":
    main()