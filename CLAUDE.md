# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

**自动填充宝 v3.5** — 基于 `customtkinter` 的 Windows 桌面 GUI 应用，根据 Excel 数据自动填充 Word 模板生成物业费诉状文书。整个代码库为单文件架构，约 1700 行，全部位于 `main.py`。

## 环境与运行

```bash
# 激活虚拟环境（主环境）
source venv_new/Scripts/activate

# 直接运行
python main.py

# 打包为 exe（需要 PyInstaller）
pyinstaller 自动填充宝.spec
```

## 核心依赖

- `customtkinter` — GUI 框架
- `pandas` + `openpyxl` — Excel 数据处理
- `docxtpl` (`DocxTemplate`) + `python-docx` — Word 模板渲染与占位符解析
- `jinja2` — `docxtpl` 的模板引擎

## 架构要点

### 单文件结构

所有逻辑在 `main.py` 中，`PaymentApp` 类（继承 `ctk.CTk`）是唯一的顶层窗体和业务控制器。没有 MVC 分层，UI 和业务逻辑紧密耦合。

### 模板类型（4 种）

通过 `self.template_type` (`ctk.StringVar`) 切换，核心计算逻辑在 `_process_excel_worker()` 中按模板类型分支：

| 模板 | 特点 |
|---|---|
| **瓯海模板** | 违约金三段式计算，车位按日期区间，含能耗 |
| **平阳模板** | 与瓯海模板结构相似，计算参数略有差异 |
| **通用模板** | 不计算，仅做占位符填充 |
| **自定义参数模板** | 可配置参数（欠费月数、违约金天数、单价等），通过 `custom_params.json` 持久化 |

### 数据处理管线

```
选择 Word 模板 → extract_placeholders_from_template() 解析 {{ }} 占位符
选择 Excel → pd.read_excel()
点击"处理原始数据" → _process_excel_worker()（后台线程）:
  1. 拆分姓名（"业主姓名"按"、"分割 → 被申请人姓名一/二）
  2. 拆分身份证地址
  3. parse_id_card() 解析身份证号 → 出生年月日、性别
  4. 按模板类型计算物业费/车位费/水费/能耗费/违约金
  5. 生成处理后 Excel（openpyxl，强制文本格式）
  6. 通过 queue.Queue 异步更新进度条
勾选条目 → 选择文件名组成列 → _generate_worker():
  DocxTemplate.render(context) → doc.save()
```

### 关键辅助函数

- `extract_placeholders_from_template(path)` — 用 `python-docx` 扫描段落、表格、页眉页脚中的 `{{ xxx }}` 占位符
- `parse_id_card(id_number)` — 从 18 位身份证提取出生年/月/日和性别
- `parse_date_range_to_days(date_str)` — 解析 `2023-01-01~2024-01-01` 或 `2023-01-01-2024-01-01` 等日期区间
- `TwoWayScrollableFrame` — 自定义双向滚动容器（Canvas 实现），嵌入在预览区

### 预览区 (Line 1323-1433)

`load_preview()` 在 `self.preview_inner`（`TwoWayScrollableFrame` 的内层 frame）中动态创建 CheckBox + Label 网格，每行对应 Excel 中的一条记录。鼠标滚轮事件通过递归绑定传递。

### 线程模型

所有重操作（Excel 处理、Word 生成）在 `threading.Thread` 中运行，通过 `queue.Queue` → `_process_queue()` (每 100ms 轮询) 向主线程发送进度/错误/完成消息。取消操作通过 `self.cancel_requested` 标志实现协作式取消。

### 占位符处理

Word 模板使用 `{{ 变量名 }}` 语法。处理流程：
1. Excel 列名必须与占位符名一致（或通过硬编码映射在 `hardcoded` 字典中补充）
2. `_generate_worker` 中先填充占位符变量，再用 `hardcoded` 字典覆盖/补充固定字段
3. `auto_generated_columns` 集合包含所有程序自动计算生成的列名，不在"可选预览列"中展示

### 配置持久化

`custom_params.json` 存储在项目根目录，启动时自动加载合并到 `self.custom_params`，参数窗口保存时全量写回。仅在"自定义参数模板"模式下可编辑。

### 打包配置

`自动填充宝.spec` — PyInstaller spec，入口 `main.py`，输出名 `自动填充宝`，带 `logo.ico` 图标，windowed 模式（无控制台）。
# 项目：自动填充宝 v3.5

## 技术栈
- Python 3.x
- customtkinter（GUI框架）
- pandas（数据处理）
- python-docx / docxtpl（Word模板填充）

## 项目特点
- 这是一个批量文书生成工具，从Excel读取数据，填充到Word模板
- 需要根据模板类型（瓯海/平阳/通用/自定义）执行不同的计算逻辑

## 编码规范
- 类名使用大驼峰
- 函数名使用小写下划线
- 所有计算需精确到小数点后2位
- 日期输出格式统一为YYYY-MM-DD
