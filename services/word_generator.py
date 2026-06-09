"""
Word 生成服务 — 使用 docxtpl 将 DataFrame 行填充到 Word 模板。

将模板渲染逻辑独立出来，便于测试和复用。
"""

import os
import re
import pandas as pd
from docxtpl import DocxTemplate
import config
from models.id_parser import clean_value


def _build_filename(row, selected_columns):
    """根据选中的列构建文件名。支持虚拟列"行号"。"""
    parts = []
    row_idx = getattr(row, 'name', 0)
    for col in selected_columns:
        if col == "行号":
            parts.append(str(row_idx + 1))
        elif col in row.index:
            val = str(row[col]) if pd.notna(row[col]) else ''
            val = re.sub(r'[\\/*?:"<>|]', '_', val)
            val = val.replace('\n', '').replace('\r', '').strip()
            if val == '':
                val = '空'
            parts.append(val)
        else:
            parts.append('缺失列')

    filename = '_'.join(parts) + '.docx'
    if len(filename) > config.MAX_FILENAME_LENGTH:
        filename = filename[:config.MAX_FILENAME_LENGTH] + '...docx'
    return filename


def _truncate_path(full_path):
    """截断过长的文件路径。"""
    if len(full_path) <= config.MAX_PATH_LENGTH:
        return full_path

    dir_part = os.path.dirname(full_path)
    filename = os.path.basename(full_path)
    excess = len(full_path) - config.MAX_PATH_LENGTH
    name_without_ext = filename[:-5]  # remove '.docx'
    if len(name_without_ext) > excess:
        new_name = name_without_ext[:len(name_without_ext) - excess] + '...docx'
        return os.path.join(dir_part, new_name)
    return full_path


def _build_context(row, placeholders):
    """从行数据和占位符列表构建模板上下文字典。"""
    context = {}

    # 占位符字段
    for ph in placeholders:
        context[ph] = clean_value(row.get(ph, ''))

    # 固定映射字段
    has_second = bool(str(row.get('被申请人姓名二', '')).strip())
    hardcoded = {
        '被申请人姓名一': clean_value(row.get('被申请人姓名一', '')),
        '身份证地址一': clean_value(row.get('身份证地址一', '')),
        '性别一': clean_value(row.get('性别一', '')),
        '出生年一': clean_value(row.get('出生年一', '    ')),
        '出生月一': clean_value(row.get('出生月一', '  ')),
        '出生日一': clean_value(row.get('出生日一', '  ')),
        '民族一': clean_value(row.get('民族一', '')),
        '被申请人姓名二': clean_value(row.get('被申请人姓名二', '')),
        '身份证地址二': clean_value(row.get('身份证地址二', '')),
        '性别二': clean_value(row.get('性别二', '')),
        '出生年二': clean_value(row.get('出生年二', '    ')),
        '出生月二': clean_value(row.get('出生月二', '  ')),
        '出生日二': clean_value(row.get('出生日二', '  ')),
        '民族二': clean_value(row.get('民族二', '')),
        '身份证号一': clean_value(row.get('身份证号一', '')),
        '联系电话一': clean_value(row.get('联系电话一', '')),
        '身份证号二': clean_value(row.get('身份证号二', '')),
        '联系电话二': clean_value(row.get('联系电话二', '')),
        '身份证地址': clean_value(row.get('身份证地址', '')),
        '物业服务费': clean_value(row.get('物业服务费', '0')),
        '物业服务费违约金合计': clean_value(row.get('物业服务费违约金合计', '0')),
        '车位管理费': clean_value(row.get('车位管理费', '0')),
        '车位管理费违约金': clean_value(row.get('车位管理费违约金', '0')),
        '公共能耗费': clean_value(row.get('公共能耗费', '0')),
        '公共能耗费违约金': clean_value(row.get('公共能耗费违约金', '0')),
        '室内自用水费': clean_value(row.get('室内自用水费', '0')),
        '欠费合计': clean_value(row.get('欠费合计', '0')),
        '总违约金合计': clean_value(row.get('总违约金合计', '0')),
        '标的总额': clean_value(row.get('标的总额', '0')),
        '物业服务费欠费周期': clean_value(row.get('物业服务费欠费周期', '')),
        '物业服务费欠费天数': clean_value(row.get('物业服务费欠费天数', '0')),
        '不动产所有人': clean_value(row.get('不动产所有人', row.get('业主姓名', ''))),
        '住宅房号': clean_value(row.get('住宅房号', row.get('身份证地址', ''))),
        '建筑面积': clean_value(row.get('建筑面积', '')),
        '车位号': clean_value(row.get('车位号', '')),
        '有第二被告': '是' if has_second else '否',
    }
    context.update(hardcoded)
    return context


def generate_documents(df, selected_indices, word_template_path,
                       placeholders, output_dir, filename_columns,
                       progress_callback=None, cancel_checker=None):
    """批量生成 Word 文档。

    Args:
        df: 处理后的 DataFrame
        selected_indices: 要处理的行索引列表
        word_template_path: Word 模板路径
        placeholders: 模板中的占位符列表
        output_dir: 输出目录
        filename_columns: 用于构建文件名的列名列表
        progress_callback: callable(current, total) — 进度回调
        cancel_checker: callable() -> bool — 取消检查

    Returns:
        (success_count, failed_list, was_cancelled, truncated_list)
        truncated_list 是 [(序号, 原始名, 截断名), ...]
    """
    count = 0
    total = len(selected_indices)
    failed = []
    truncated = []

    for i, idx in enumerate(selected_indices):
        if cancel_checker and cancel_checker():
            return count, failed, True, truncated

        row = df.iloc[idx]

        # 构建文件名
        filename = _build_filename(row, filename_columns)
        full_path = os.path.join(output_dir, filename)
        save_path = _truncate_path(full_path)

        # 检测是否被截断
        if save_path != full_path:
            truncated.append((row.get('序号', idx + 1), filename, os.path.basename(save_path)))

        # 构建上下文并渲染
        context = _build_context(row, placeholders)

        try:
            doc = DocxTemplate(word_template_path)
            doc.render(context)
            doc.save(save_path)
            count += 1
        except Exception as e:
            failed.append((row.get('序号', idx + 1), str(e)))
            continue

        if progress_callback:
            progress_callback(i + 1, total)

    return count, failed, False, truncated
