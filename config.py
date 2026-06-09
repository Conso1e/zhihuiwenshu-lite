"""
智汇文枢 v4.0 — 全局常量与配置

集中管理所有魔法数字、默认参数和可配置项。
修改此文件中的值会影响整个应用的计算逻辑。
"""

# ====================== 推广精简版开关 ======================
# True = 精简推广版（仅通用填充，不含客户定制公式），False = 完整定制版
LITE_MODE = True

# ====================== 违约金计算常量 ======================
PENALTY_DAYS_2 = 627          # 第二年违约金天数
PENALTY_DAYS_3 = 261          # 第三年违约金天数
OVERTIME_DAYS = 284           # 平阳模板专用的逾期天数
DEFAULT_PENALTY_DAYS_1 = 992  # 第一年违约金默认天数（无确切日期时使用）
DAILY_RATE = 0.0001           # 违约金日利率
PENALTY_END_DATE = "2025-09-30"  # 违约金计算的截止日期

# ====================== 单价常量 ======================
WATER_RATE = 3.55             # 水费单价（元/吨）
ONE_DAY_PARKING_FEE = 1.97    # 车位管理费日单价（元/个/天）

# ====================== 默认欠费参数（无日期区间时回退） ======================
DEFAULT_PROPERTY_MONTHS = 33
DEFAULT_PROPERTY_EXTRA_DAYS = 10
DEFAULT_DAYS_PER_MONTH = 31

# ====================== 自动生成的列名（不在预览列选择中出现） ======================
AUTO_GENERATED_COLUMNS = {
    '被申请人姓名一', '身份证地址一', '性别一', '出生年一', '出生月一', '出生日一',
    '被申请人姓名二', '身份证地址二', '性别二', '出生年二', '出生月二', '出生日二',
    '物业服务费', '一年物业服务费', '一日物业服务费', '第一年逾期时间', '车位欠费天数',
    '物业服务费欠费周期',
    '物业服务费违约金一', '物业服务费违约金二',
    '物业服务费违约金三', '物业服务费违约金合计', '车位管理费', '车位管理费违约金',
    '公共能耗费', '公共能耗费违约金', '室内自用水费', '欠费合计', '总违约金合计',
    '标的总额', '不动产所有人', '住宅房号', '建筑面积', '车位号',
    '联系电话一', '联系电话二', '第一年物业费逾期天数',
}

# ====================== GUI 常量 ======================
WINDOW_TITLE = "智汇文枢 Lite"
WINDOW_SIZE = "1200x850"
SIDEBAR_WIDTH = 190

# 模板类型列表（精简版仅保留通用填充型，完整版包含客户定制公式）
if LITE_MODE:
    TEMPLATE_TYPES = ["通用填充型"]
else:
    TEMPLATE_TYPES = [
        "预设公式型（瓯海）",
        "预设公式型（平阳）",
        "通用填充型",
        "自定义公式型",
    ]

# 预览区默认额外列
DEFAULT_PREVIEW_COLUMNS = ['身份证地址']

# 文件名默认组成列
DEFAULT_FILENAME_COLUMNS = ['行号', '业主姓名', '身份证地址']

# 最大预览额外列数
MAX_PREVIEW_EXTRA_COLUMNS = 5

# 文件名最大长度
MAX_FILENAME_LENGTH = 200
MAX_PATH_LENGTH = 240

# ====================== UI 配色常量 ======================
# 设计系统：Trust & Authority + Swiss Modernism 2.0
# 法律文书工具 — 权威海军蓝 + 信任琥珀金
#
# 深色主题
DARK = {
    'sidebar_bg': "#0A0F1A",          # 极深海军蓝侧边栏
    'sidebar_hover': "#1E293B",        # Slate-800
    'sidebar_text': "#64748B",         # Slate-500
    'sidebar_active': "#F59E0B",       # 琥珀金活跃指示
    'main_bg': "#0F172A",              # Slate-900
    'card_bg': "#1E293B",              # Slate-800 卡片
    'card_border': "#334155",          # Slate-700 边框
    'accent': "#60A5FA",               # Blue-400（深色下提亮）
    'accent_hover': "#93BBFD",         # Blue-300
    'accent_light': "#1E3A5F",         # 深蓝底
    'text_primary': "#F1F5F9",         # Slate-100
    'text_secondary': "#94A3B8",       # Slate-400
    'danger': "#FCA5A5",               # Red-300
    'danger_hover': "#F87171",         # Red-400
    'success': "#6EE7B7",              # Emerald-300
    'success_hover': "#34D399",        # Emerald-400
    'warning': "#FCD34D",              # Amber-300
    'entry_bg': "#0F172A",
    'entry_border': "#475569",         # Slate-600
    'header_bg': "#1E293B",
}
# 浅色主题
LIGHT = {
    'sidebar_bg': "#0F172A",           # 深海军蓝侧边栏（浅色模式也保持深色）
    'sidebar_hover': "#1E293B",        # Slate-800
    'sidebar_text': "#94A3B8",         # Slate-400
    'sidebar_active': "#F59E0B",       # 琥珀金活跃指示
    'main_bg': "#F8FAFC",              # Slate-50 冷灰背景
    'card_bg': "#FFFFFF",              # 白卡片
    'card_border': "#E2E8F0",          # Slate-200 边框
    'accent': "#1E40AF",               # Navy-800 主色
    'accent_hover': "#1E3A8A",         # Navy-900 hover
    'accent_light': "#EFF6FF",         # Blue-50 淡底
    'text_primary': "#0F172A",         # Slate-900 正文
    'text_secondary': "#64748B",       # Slate-500 辅助文字
    'danger': "#DC2626",               # Red-600
    'danger_hover': "#B91C1C",         # Red-700
    'success': "#059669",              # Emerald-600
    'success_hover': "#047857",        # Emerald-700
    'warning': "#D97706",              # Amber-600
    'entry_bg': "#FFFFFF",
    'entry_border': "#CBD5E1",         # Slate-300
    'header_bg': "#F1F5F9",            # Slate-100
}

# 房屋类型 → 默认单价映射
HOUSE_TYPE_PRICE_MAP = {
    '住宅': '2.5',
    '商业': '4',
}
