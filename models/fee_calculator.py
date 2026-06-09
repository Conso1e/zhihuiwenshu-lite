"""
费用计算器模块 — 策略模式

每种模板类型对应一个计算器子类，统一接口 calculate(row) -> dict。
工厂函数 get_calculator() 根据模板类型返回对应实例。
"""

from abc import ABC, abstractmethod
import pandas as pd
import config
from models.id_parser import (
    safe_float, parse_date_range_to_days, extract_start_date, get_property_days,
)


# 计算器返回的字段列表（用于构建错误时的空字典）
RESULT_COLUMNS_PROPERTY = [
    '物业服务费开始时间', '物业服务费', '一年物业服务费', '一日物业服务费',
    '物业服务费违约金一', '物业服务费违约金二', '物业服务费违约金三',
    '物业服务费违约金合计', '车位欠费天数', '车位管理费', '车位管理费违约金',
    '室内自用水费', '欠费合计', '总违约金合计', '标的总额',
    '物业服务费欠费周期', '物业服务费欠费天数', '第一年物业费逾期天数',
]

RESULT_COLUMNS_CUSTOM = [
    '物业服务费', '一年物业服务费', '一日物业服务费', '违约金合计',
    '车位欠费天数', '车位管理费', '车位管理费违约金',
    '公共能耗费', '公共能耗费违约金', '室内自用水费',
    '欠费合计', '总违约金合计', '标的总额', '物业服务费欠费周期',
]


class BaseFeeCalculator(ABC):
    """费用计算器基类"""

    def _empty_result(self, columns):
        """返回空结果字典（计算异常时使用）"""
        return {col: '' for col in columns}

    def _parse_parking_num(self, row):
        """从车位号字符串解析车位数"""
        parking_spaces = row.get('车位号', '')
        if parking_spaces and isinstance(parking_spaces, str):
            num = len([s for s in parking_spaces.split('/') if s.strip()])
            if num > 0:
                return num
        num = safe_float(row.get('车位数', 1))
        return int(num) if num > 0 else 1

    def _parse_parking_days(self, row):
        """从日期区间或数值列获取车位欠费天数"""
        parking_date_range = row.get('车位起止时间', '')
        days = parse_date_range_to_days(parking_date_range) + 1
        if days > 0:
            return days
        try:
            return int(float(row.get('车位欠费天数', 0)))
        except (ValueError, TypeError):
            return 0

    def _calc_property_params(self, row):
        """计算物业费相关的基础参数。
        返回: (unit_price, area, property_period_raw, property_days, property_start_str)
        """
        unit_price = safe_float(row.get('单价'))
        area = safe_float(row.get('建筑面积'))
        property_period_raw = row.get('物业服务费欠费周期', '')
        property_days = parse_date_range_to_days(property_period_raw) + 1
        property_start_str = extract_start_date(property_period_raw)

        if property_days == 0:
            try:
                property_days = int(float(row.get('物业服务费欠费天数', 0)))
            except (ValueError, TypeError):
                property_days = 0

        return unit_price, area, property_period_raw, property_days, property_start_str

    def _calc_property_fee(self, unit_price, area, property_days):
        """计算物业费总额、一年物业费、一日物业费"""
        if property_days > 0:
            one_year_fee = unit_price * area * 12
            one_day_fee = one_year_fee / 365
            total_property = one_day_fee * property_days
        else:
            months = config.DEFAULT_PROPERTY_MONTHS
            days = config.DEFAULT_PROPERTY_EXTRA_DAYS
            days_per_month = config.DEFAULT_DAYS_PER_MONTH
            total_property = unit_price * area * months + unit_price * area / days_per_month * days
            one_year_fee = unit_price * area * 12
            one_day_fee = one_year_fee / 365
        return total_property, one_year_fee, one_day_fee

    def _calc_penalty_days_1(self, row):
        """计算第一年违约金天数"""
        penalty_start_date_raw = row.get('第一年物业费逾期开始时间', '')
        if penalty_start_date_raw and pd.notna(penalty_start_date_raw):
            try:
                date_str = str(penalty_start_date_raw).strip()
                if ' ' in date_str:
                    date_str = date_str.split(' ')[0]
                penalty_start_date = pd.to_datetime(date_str)
                end_date = pd.to_datetime(config.PENALTY_END_DATE)
                return max((end_date - penalty_start_date).days, 0) + 1
            except Exception as e:
                print(f"日期解析失败: {penalty_start_date_raw}, 错误: {e}")
        return config.DEFAULT_PENALTY_DAYS_1

    @abstractmethod
    def validate_columns(self, df):
        """验证并补全 DataFrame 的必要列。返回警告消息列表。"""
        ...

    @abstractmethod
    def calculate(self, row):
        """计算单行费用。返回字段字典。"""
        ...

    def get_result_columns(self):
        """返回此计算器会产生的所有结果列名"""
        return RESULT_COLUMNS_PROPERTY


class OuhaiCalculator(BaseFeeCalculator):
    """瓯海模板计算器"""

    def validate_columns(self, df):
        warnings = []
        for col in ['单价', '建筑面积']:
            if col not in df.columns:
                df[col] = '0'
        for col, msg in [
            ('车位起止时间', "Excel中缺少 '车位起止时间' 列，已使用默认值0，请检查数据！"),
            ('能耗欠费时间', "Excel中缺少 '能耗欠费时间' 列，已使用默认值0，请检查数据！"),
        ]:
            if col not in df.columns:
                df[col] = '0'
                warnings.append(msg)
        for col in ['车位管理费', '公共能耗费']:
            if col not in df.columns:
                df[col] = ''
        return warnings

    def calculate(self, row):
        try:
            unit_price, area, property_period_raw, property_days, property_start_str = \
                self._calc_property_params(row)
            water_usage = safe_float(row.get('室内用水'))
            parking_num = self._parse_parking_num(row)

            total_property, one_year_fee, one_day_fee = \
                self._calc_property_fee(unit_price, area, property_days)

            days_penalty_1 = self._calc_penalty_days_1(row)

            # 瓯海违约金公式
            penalty1 = one_year_fee * config.DAILY_RATE * days_penalty_1
            penalty2 = one_year_fee * config.DAILY_RATE * config.PENALTY_DAYS_2
            penalty3 = (one_year_fee / 365) * config.PENALTY_DAYS_3 * config.DAILY_RATE * config.PENALTY_DAYS_3
            total_property_penalty = penalty1 + penalty2 + penalty3

            # 车位
            parking_total_days = self._parse_parking_days(row)
            total_parking = config.ONE_DAY_PARKING_FEE * parking_total_days * parking_num
            parking_penalty = total_parking * config.DAILY_RATE * config.PENALTY_DAYS_3

            # 水费
            water_fee = water_usage * config.WATER_RATE

            # 合计
            total_due = total_property + total_parking + water_fee
            total_penalty = total_property_penalty + parking_penalty
            total_subject = total_due + total_penalty

            return {
                '物业服务费开始时间': property_start_str,
                '物业服务费': round(total_property, 2),
                '一年物业服务费': round(one_year_fee, 2),
                '一日物业服务费': round(one_day_fee, 2),
                '物业服务费违约金一': round(penalty1, 2),
                '物业服务费违约金二': round(penalty2, 2),
                '物业服务费违约金三': round(penalty3, 2),
                '物业服务费违约金合计': round(total_property_penalty, 2),
                '车位欠费天数': round(parking_total_days, 2),
                '车位管理费': round(total_parking, 2),
                '车位管理费违约金': round(parking_penalty, 2),
                '室内自用水费': round(water_fee, 2),
                '欠费合计': round(total_due, 2),
                '总违约金合计': round(total_penalty, 2),
                '标的总额': round(total_subject, 2),
                '物业服务费欠费周期': property_period_raw,
                '物业服务费欠费天数': round(property_days, 0),
                '第一年物业费逾期天数': round(days_penalty_1, 0),
            }
        except Exception as e:
            print(f"行 {row.name} 计算出错: {e}")
            return self._empty_result(RESULT_COLUMNS_PROPERTY)


class PingyangCalculator(OuhaiCalculator):
    """平阳模板计算器 — 继承瓯海，仅覆盖差异部分"""

    def validate_columns(self, df):
        warnings = []
        for col in ['单价', '建筑面积', '室内用水']:
            if col not in df.columns:
                df[col] = '0'
                if col == '室内用水':
                    warnings.append("Excel中缺少 '室内用水' 列，已使用默认值0，请检查数据！")
        if '公共能耗费' not in df.columns:
            df['公共能耗费'] = ''
        return warnings

    def calculate(self, row):
        try:
            unit_price, area, property_period_raw, property_days, property_start_str = \
                self._calc_property_params(row)
            water_usage = safe_float(row.get('室内用水'))
            parking_num = self._parse_parking_num(row)

            total_property, one_year_fee, one_day_fee = \
                self._calc_property_fee(unit_price, area, property_days)

            days_penalty_1 = self._calc_penalty_days_1(row)

            # 平阳违约金公式（penalty3 使用 overtime_days）
            penalty1 = one_year_fee * config.DAILY_RATE * days_penalty_1
            penalty2 = one_year_fee * config.DAILY_RATE * config.PENALTY_DAYS_2
            penalty3 = (one_year_fee / 365) * config.OVERTIME_DAYS * config.DAILY_RATE * config.PENALTY_DAYS_3
            total_property_penalty = penalty1 + penalty2 + penalty3

            # 车位
            parking_total_days = self._parse_parking_days(row)
            total_parking = config.ONE_DAY_PARKING_FEE * parking_total_days * parking_num
            parking_penalty = total_parking * config.DAILY_RATE * config.PENALTY_DAYS_3

            # 水费
            water_fee = water_usage * config.WATER_RATE

            # 合计
            total_due = total_property + total_parking + water_fee
            total_penalty = total_property_penalty + parking_penalty
            total_subject = total_due + total_penalty

            return {
                '物业服务费开始时间': property_start_str,
                '物业服务费': round(total_property, 2),
                '一年物业服务费': round(one_year_fee, 2),
                '一日物业服务费': round(one_day_fee, 2),
                '物业服务费违约金一': round(penalty1, 2),
                '物业服务费违约金二': round(penalty2, 2),
                '物业服务费违约金三': round(penalty3, 2),
                '物业服务费违约金合计': round(total_property_penalty, 2),
                '车位欠费天数': round(parking_total_days, 2),
                '车位管理费': round(total_parking, 2),
                '车位管理费违约金': round(parking_penalty, 2),
                '室内自用水费': round(water_fee, 2),
                '欠费合计': round(total_due, 2),
                '总违约金合计': round(total_penalty, 2),
                '标的总额': round(total_subject, 2),
                '物业服务费欠费周期': property_period_raw,
                '物业服务费欠费天数': round(property_days, 0),
                '第一年物业费逾期天数': round(days_penalty_1, 0),
            }
        except Exception as e:
            print(f"行 {row.name} 计算出错: {e}")
            return self._empty_result(RESULT_COLUMNS_PROPERTY)


class FormulaCalculator(BaseFeeCalculator):
    """自定义公式模板计算器 — 由用户定义计算公式"""

    def __init__(self, formulas):
        self.formulas = formulas  # 公式字符串列表

    def validate_columns(self, df):
        return []  # 公式引擎自行处理列匹配

    def calculate(self, row):
        return {}  # 公式计算在 excel_processor 中统一处理

    def get_result_columns(self):
        return []  # 结果列动态生成


class GenericCalculator(BaseFeeCalculator):
    """通用模板计算器 — 不计算，仅透传占位符"""

    def validate_columns(self, df):
        return []

    def calculate(self, row):
        return {}

    def get_result_columns(self):
        return []


# ====================== 工厂函数 ======================

def get_calculator(template_type, formulas=None):
    """根据模板类型返回对应的计算器实例。

    Args:
        template_type: "瓯海模板" | "平阳模板" | "通用模板" | "自定义参数模板"
        formulas: list, 仅自定义参数模板需要

    Returns:
        BaseFeeCalculator 子类实例
    """
    if config.LITE_MODE:
        return GenericCalculator()
    if template_type == "预设公式型（瓯海）":
        return OuhaiCalculator()
    elif template_type == "预设公式型（平阳）":
        return PingyangCalculator()
    elif template_type == "自定义公式型":
        return FormulaCalculator(formulas or [])
    else:
        return GenericCalculator()
