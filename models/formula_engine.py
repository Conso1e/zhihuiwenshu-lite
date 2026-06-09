"""
公式引擎 — 解析用户自定义公式，匹配 Excel 列，执行计算。

公式格式: {{变量1}} 运算符 {{变量2}} = {{结果变量}}
示例: {{单价}}*{{建筑面积}}*12/365*{{欠费天数}}={{物业费总额}}
支持: + - * / ( )
"""

import re
import pandas as pd


def parse_formula(formula_str):
    """解析一行公式字符串。

    Args:
        formula_str: 如 "{{单价}}*{{数量}}={{总价}}"

    Returns:
        (variables, result, expr, var_map) 或 None
        variables: ["单价", "数量"]
        result: "总价"
        expr: "_v0*_v1" (安全占位表达式)
        var_map: {"_v0": "单价", "_v1": "数量"}
    """
    formula_str = formula_str.strip()
    if not formula_str or '=' not in formula_str:
        return None

    left, right = formula_str.rsplit('=', 1)
    left = left.strip()
    right = right.strip()

    result_match = re.search(r'\{\{(.+?)\}\}', right)
    if not result_match:
        return None
    result_var = result_match.group(1).strip()

    var_matches = re.findall(r'\{\{(.+?)\}\}', left)
    variables = [v.strip() for v in var_matches]
    if not variables:
        return None

    # 用安全占位名替换中文变量名
    expr = left
    var_map = {}
    for i, v in enumerate(variables):
        safe = f'_v{i}'
        var_map[safe] = v
        expr = expr.replace('{{' + v + '}}', safe)

    return variables, result_var, expr, var_map


def evaluate_row(row, formula_vars, expression, var_map):
    """对 DataFrame 的一行执行公式计算。

    Args:
        row: pandas Series (一行数据)
        formula_vars: 原始变量名列表
        expression: 安全占位表达式 (如 "_v0*_v1")
        var_map: 安全名到原始名的映射

    Returns:
        计算结果 (float)，失败返回 0.0
    """
    local_vars = {}
    for safe, real in var_map.items():
        val = row.get(real, 0)
        try:
            local_vars[safe] = float(val) if val != '' and pd.notna(val) else 0.0
        except (ValueError, TypeError):
            local_vars[safe] = 0.0

    try:
        result = eval(expression, {"__builtins__": {}}, local_vars)
        return round(float(result), 2)
    except Exception as e:
        print(f"公式计算错误: {expression}, 错误: {e}")
        return 0.0


def process_formulas(df, formulas):
    """批量执行公式计算。

    Args:
        df: pandas DataFrame
        formulas: 公式字符串列表

    Returns:
        (df, all_variables, results)
    """
    parsed = []
    all_vars = set()
    results = []

    for f in formulas:
        p = parse_formula(f)
        if p:
            variables, result_var, expr, var_map = p
            parsed.append((variables, result_var, expr, var_map))
            all_vars.update(variables)
            results.append(result_var)
        else:
            print(f"公式解析失败: {f}")

    for variables, result_var, expr, var_map in parsed:
        def calc(row, _expr=expr, _vmap=var_map):
            return evaluate_row(row, variables, _expr, _vmap)
        df[result_var] = df.apply(calc, axis=1)

    return df, all_vars, results


def match_variables(word_placeholders, excel_columns, formula_results):
    """变量匹配检测 — 找出 Word 模板需要的变量与 Excel 列的对应关系。

    Args:
        word_placeholders: Word 模板中 {{ }} 占位符列表
        excel_columns: Excel 所有列名列表
        formula_results: 公式引擎产生的结果列名列表

    Returns:
        (matched, unmatched, summary)
        matched: [(变量名, 来源), ...] — 来源: "excel" 或 "formula"
        unmatched: [变量名, ...] — 既不在 Excel 也不在公式结果中的变量
        summary: dict — {total, matched_count, unmatched_count}
    """
    matched = []
    unmatched = []
    excel_set = set(excel_columns)
    result_set = set(formula_results)

    for ph in word_placeholders:
        if ph in excel_set:
            matched.append((ph, "excel"))
        elif ph in result_set:
            matched.append((ph, "formula"))
        else:
            unmatched.append(ph)

    summary = {
        'total': len(word_placeholders),
        'matched_count': len(matched),
        'unmatched_count': len(unmatched),
    }
    return matched, unmatched, summary


def validate_formula(formula_str, valid_variables):
    """校验单条公式，返回 (is_valid, error_message).

    校验规则：
    1. 必须包含恰好一个 =
    2. 右侧必须恰好有一个 {{结果变量}}
    3. 左侧必须至少有一个 {{变量}}
    4. 所有 {{变量}}（除结果变量外）必须在 valid_variables 中
    5. 结果变量如果是新变量（不在 valid_variables 中），允许但提示

    Args:
        formula_str: 公式字符串，如 "{{单价}}*{{建筑面积}}={{物业费}}"
        valid_variables: 有效变量名集合（Word占位符 ∪ Excel列名）

    Returns:
        (True, "") 或 (False, "错误描述")
    """
    formula_str = formula_str.strip()

    if '=' not in formula_str:
        return False, "缺少等号（=），格式应为：{{变量}} 运算符 {{变量}} = {{结果}}"

    parts = formula_str.split('=')
    if len(parts) > 2:
        return False, f"包含多个等号（{len(parts) - 1}个），公式中只能有一个等号"

    left = parts[0].strip()
    right = parts[1].strip()

    if not left:
        return False, "等号左侧为空，需要至少一个 {{变量}} 或运算表达式"

    if not right:
        return False, "等号右侧为空，需要一个 {{结果变量}}"

    # 提取右侧结果变量
    right_vars = re.findall(r'\{\{(.+?)\}\}', right)
    if len(right_vars) == 0:
        return False, "等号右侧必须包含一个 {{结果变量}}，如 {{物业费总额}}"
    if len(right_vars) > 1:
        return False, f"等号右侧只能有一个 {{结果变量}}，当前有 {len(right_vars)} 个：{', '.join(right_vars)}"

    # 提取左侧变量
    left_vars = re.findall(r'\{\{(.+?)\}\}', left)
    if len(left_vars) == 0:
        return False, "等号左侧至少需要一个 {{变量}} 参与运算"

    # 检查左侧变量是否在有效变量集合中
    invalid_vars = [v.strip() for v in left_vars if v.strip() not in valid_variables]
    if invalid_vars:
        return False, f"以下变量在变量对照表中不存在：{', '.join(invalid_vars)}"

    # 基本表达式安全检查：替换变量占位符后检查语法
    test_expr = left
    for i, v in enumerate(left_vars):
        test_expr = test_expr.replace('{{' + v + '}}', f'_v{i}')
    # 移除空格检查可eval性
    try:
        test_vars = {f'_v{i}': 1.0 for i in range(len(left_vars))}
        eval(test_expr.replace(' ', ''), {"__builtins__": {}}, test_vars)
    except SyntaxError as e:
        return False, f"表达式语法错误：{e}"
    except Exception:
        pass  # 运行时错误（如除零）在 eval 测试时可能出现，但语法没问题

    return True, ""
