# 测试用例记录表（TestCase Viewer）— 产品需求文档（PRD）

> **项目**：TestCase Viewer  
> **技术栈**：Flask + 原生 HTML/CSS/JS + openpyxl  
> **当前版本**：v1.3.4（2026-07-05）  
> **本 PRD 目标**：完整记录产品从 v1.0 到 v2.x 的版本演进，当前重点为「汇总统计」功能升级  
> **文档版本**：v2.0-draft

---

## 一、版本总览

### 1.1 历史版本（v1.x 已发布）

| 版本 | 日期 | 里程碑名称 | 核心功能点 | 状态 |
|------|------|-----------|-----------|------|
| **v1.0** | 2026-07-03 | 基础执行闭环 | ① Flask + 浏览器前端 ② Excel 读取→HTML 卡片展示→结果回写 ③ 进度条 + 图例 | ✅ 已发布 |
| **v1.1** | 2026-07-03 | 搜索筛选 + 汇总统计 | ① 全文关键词搜索 ② 结果/优先级/模块筛选 ③ 汇总统计页（统计卡片+环形图+分布表） ④ 三个标签页 | ✅ 已发布 |
| **v1.2** | 2026-07-04 | 文件上传与记忆 | ① 上传页 + _active.xlsx 记忆机制 ② 智能 Sheet 选择（_pick_best_sheet） ③ 表头行自动定位 ④ 刷新按钮 ⑤ 完整字段展示 | ✅ 已发布 |
| **v1.3** | 2026-07-05 | 附加字段 + 问题列表 | ① 「测试现象备注」→「实际结果」更名 ② 新增测试人员/BugID/Bug频率/问题时间 ③ 失败自动记录时间 ④ 问题列表标签页 + 庆祝动画 ⑤ UI 美化（2×2 网格布局） | ✅ 已发布 |
| **v1.3.1** | 2026-07-05 | Bug 修复批次 1 | ① B4: loadIssues 死代码导致 JS 解析失败 ② B5: detect_columns 列映射覆盖 ③ B6: result_col 误匹配 Pass/Fail ④ B7: 读回改用 SAVE_COLUMNS 精确匹配 ⑤ B8: _active.xlsx 降级逻辑 ⑥ B9: 读/写 Sheet 不一致 ⑦ B10: 副本文件困惑 | ✅ 已发布 |
| **v1.3.4** | 2026-07-05 | 列映射加固 | ① save_result 按 SAVE_COLUMNS 精确列名写入 ② _pick_best_sheet 过滤规则：核心字段 ≥ 3 才视为用例 Sheet ③ read_testcases/all-status 按 STATE 中 sheet_name + header_row 精确匹配 | ✅ 已发布 |

### 1.2 本次升级版本（v2.x 规划中）

本次升级将一个完整的大需求拆分为 **4 个递进版本**，遵循「先地基、再上层、后增强」原则，每个版本可独立交付、独立验收。

| 版本 | 里程碑名称 | 一句话描述 | 核心功能点 | 依赖 |
|------|-----------|-----------|-----------|------|
| **v2.0** | 多 Sheet 读取与分类 | 重构读取层，支持读取 Excel 全部 Sheet 并自动区分为「测试用例 Sheet」与「非用例 Sheet」 | ① 多 Sheet 全量读取 ② Sheet 分类识别 ③ STATE 多 Sheet 化重构 ④ 全局用例索引 | 无（地基） |
| **v2.1** | 按 Sheet 分组汇总展示 | 在汇总统计页新增「按 Sheet 分组的用例列表」，每条仅展示序号/名称/结果 | ① 分组用例列表 API ② 汇总页新增分组列表区 ③ 空 Sheet/空文件兜底展示 | v2.0 |
| **v2.2** | 用例名称点击跳转 | 在分组列表中点击用例名称，跳转到执行测试视图并定位到该用例 | ① 全局索引映射 ② 跨视图跳转交互 ③ 高亮定位动效 | v2.1 |
| **v2.3** | 非用例 Sheet 信息展示 | 在顶部 Tab 栏新增「参考信息」模块，展示非用例 Sheet 的原始内容 | ① 非用例 Sheet 内容 API ② 新增 Tab 与表格渲染 ③ 空/异常兜底 | v2.0 |

### 版本依赖关系

```
v2.0 (多Sheet读取与分类)
 ├──> v2.1 (按Sheet分组汇总展示)
 │     └──> v2.2 (用例名称点击跳转)
 └──> v2.3 (非用例Sheet信息展示)
```

- v2.0 是所有后续版本的地基，必须最先完成
- v2.1 和 v2.3 都只依赖 v2.0，理论上可并行，但 v2.1 是核心用户价值，优先级更高
- v2.2 依赖 v2.1 的分组列表 UI，必须在 v2.1 之后

### 交付优先级

```
v2.0 → v2.1 → v2.2 → v2.3
```

### 1.3 v1.x 已有 API 清单（v1.3.4 基线）

| API | 方法 | 引入版本 | 说明 | v2.0 后状态 |
|-----|------|---------|------|------------|
| `/api/init` | GET | v1.0 | 初始化状态（文件是否加载、用例总数） | 变更（新增 Sheet 概要） |
| `/api/titles` | GET | v1.1 | 全部用例标题列表（用于下拉跳转） | 不变 |
| `/api/testcase/<index>` | GET | v1.0 | 单条用例详情（含字段映射后的展示数据） | 变更（新增 `_sheet_name`） |
| `/api/all-status` | GET | v1.0 | 所有用例的执行状态数组 | 微调（适配多 Sheet） |
| `/api/search` | POST | v1.1 | 搜索筛选（关键词+结果+优先级+模块+分页） | 不变（自动覆盖全量） |
| `/api/save` | POST | v1.0 | 保存测试结果到 Excel | 微调（Sheet 名从用例取） |
| `/api/summary` | GET | v1.1 | 汇总统计（总数/通过率/优先级分布/模块分布） | 不变（自动覆盖全量） |
| `/api/filter-options` | GET | v1.1 | 筛选下拉选项（优先级/模块列表） | 不变 |
| `/api/reload` | GET | v1.2 | 重新加载 Excel（刷新按钮） | 微调（调用 `read_all_sheets`） |
| `/api/upload` | POST | v1.2 | 上传新 Excel 文件 | 微调（调用 `read_all_sheets`） |
| `/api/check-results` | GET | v1.3 | 检查文件是否已有执行结果（替换确认） | 不变 |
| `/api/issues` | GET | v1.3 | 问题列表（失败/阻塞用例） | 不变 |

### 1.4 v1.x 已有视图与数据模型（v1.3.4 基线）

**现有视图（3 个 Tab）：**

| Tab | 引入版本 | 说明 |
|-----|---------|------|
| 执行测试 | v1.0 | 逐条卡片展示用例，选择结果+填写备注→保存回写 Excel |
| 汇总统计 | v1.1 | 统计卡片+环形进度图+优先级分布表+模块分布表 |
| 问题列表 | v1.3 | 展示失败/阻塞用例，全部通过时庆祝动画 |

**现有数据模型（`app/models/testcase.py`）：**

| 数据类 | 引入版本 | 说明 |
|--------|---------|------|
| `TestCase` | v1.0 | 一条测试用例的领域数据（id/title/module/priority/steps/expected 等） |
| `ExecutionRecord` | v1.3 | 一次执行的结果记录（result/actual_result/tester/bug_id/bug_frequency/issue_time） |
| `SheetMeta` | v1.2 | Excel Sheet 解析元数据（sheet_name/header_row/headers/mapping） |
| `LoadedData` | v1.2 | read_testcases 返回的完整数据包（testcases/mapping/headers/sheet_name/header_row） |

**现有 STATE 全局状态（v1.3.4）：**

```python
STATE = {
    'testcases': [],       # 用例列表（单 Sheet）
    'mapping': {},         # 列映射
    'headers': [],         # 表头
    'filepath': None,
    'filename': '',
    'sheet_name': None,    # 当前 Sheet 名
    'header_row': 1,       # 表头行号
}
```

**现有核心函数：**

| 函数 | 引入版本 | 说明 |
|------|---------|------|
| `read_testcases(filepath)` | v1.0 | 读取 Excel，返回 LoadedData |
| `_pick_best_sheet(wb)` | v1.2 | 从所有 Sheet 中选出最佳用例 Sheet |
| `detect_columns(ws, header_row)` | v1.2 | 自动识别列映射 |
| `_find_header_row(ws)` | v1.2 | 定位表头行 |
| `save_result(...)` | v1.0 | 保存结果到 Excel |
| `find_excel_file()` | v1.2 | 查找 testcases/ 下的 Excel 文件 |

---

## 二、v2.0 — 多 Sheet 读取与分类（地基）

### 2.1 功能描述

当前 `_pick_best_sheet` 只从所有 Sheet 中选出「最佳」的一个进行用例解析，其余 Sheet 全部忽略。本版本重构读取层，实现：

1. **全量读取**：读取 Excel 文件中的所有 Sheet，不再只选一个
2. **分类识别**：对每个 Sheet 自动判定属于「测试用例 Sheet」还是「非用例 Sheet」
   - 判定标准复用现有 `_pick_best_sheet` 的核心字段评分逻辑（核心字段数 ≥ 阈值即为用例 Sheet）
   - 名称含「用例/Case/TestCase」等关键词的 Sheet，放宽阈值（与现有逻辑一致）
   - 不满足阈值的 Sheet 归为「非用例 Sheet」
3. **全局用例索引**：所有用例 Sheet 的用例合并为一个扁平列表，每条用例标记 `_sheet_name`，保持全局 index 连续
4. **保持原始 Sheet 顺序**：按 Excel 中 Sheet 的出现顺序排列

### 2.2 数据模型变更

#### 2.2.1 `SheetMeta` 新增字段

```python
@dataclass
class SheetMeta:
    sheet_name: str = ''
    sheet_type: str = 'testcase'   # 【新增】'testcase' | 'note'
    header_row: int = 1
    headers: List[str] = field(default_factory=list)
    mapping: Dict[str, int] = field(default_factory=dict)
    case_count: int = 0            # 【新增】该 Sheet 的用例数量（note 类型为 0）
```

#### 2.2.2 `LoadedData` 重构为多 Sheet

```python
@dataclass
class LoadedData:
    # 【变更】从单 Sheet 扁平结构 → 多 Sheet 结构
    sheets: List[SheetMeta] = field(default_factory=list)       # 所有 Sheet 元信息（按 Excel 顺序）
    testcases: List[Dict] = field(default_factory=list)         # 所有用例 Sheet 的用例合并扁平列表
                                                                 # 每条 tc 新增 _sheet_name 字段
    testcase_sheets: List[str] = field(default_factory=list)    # 用例 Sheet 名称列表
    note_sheets: List[str] = field(default_factory=list)        # 非用例 Sheet 名称列表
    # 以下字段保留但语义变更为「兼容字段」（取第一个用例 Sheet 的值）
    mapping: Dict[str, int] = field(default_factory=dict)
    headers: List[str] = field(default_factory=list)
    sheet_name: str = ''
    header_row: int = 1
```

#### 2.2.3 STATE 全局状态重构

```python
STATE = {
    # 【新增】多 Sheet 结构
    'sheets': [],                  # List[SheetMeta]，所有 Sheet 元信息
    'testcase_sheets': [],         # List[str]，用例 Sheet 名称
    'note_sheets': [],             # List[str]，非用例 Sheet 名称

    # 【变更】扁平用例列表（每条 tc 带 _sheet_name）
    'testcases': [],               # 所有用例 Sheet 合并，全局 index 连续
    'mapping': {},                 # 兼容字段：第一个用例 Sheet 的 mapping
    'headers': [],                 # 兼容字段：第一个用例 Sheet 的 headers
    'sheet_name': None,            # 兼容字段：第一个用例 Sheet 名称
    'header_row': 1,               # 兼容字段：第一个用例 Sheet 的表头行

    'filepath': None,
    'filename': '',
}
```

#### 2.2.4 用例 dict 新增字段

每条用例 dict 新增 `_sheet_name` 字段，标记该用例所属 Sheet：

```python
tc = {
    '_row': 12,                    # Excel 行号（原有）
    '_sheet_name': '登录用例',      # 【新增】所属 Sheet 名称
    'col_0': '...',
    # ... 其余字段不变
}
```

### 2.3 新增 / 变更 API

#### 2.3.1 新增 `GET /api/sheets`

返回所有 Sheet 的分类信息。

**响应示例**：

```json
{
  "testcase_sheets": [
    { "name": "登录用例", "case_count": 25, "header_row": 2 },
    { "name": "支付用例", "case_count": 18, "header_row": 1 }
  ],
  "note_sheets": [
    { "name": "测试备注", "header_row": 1 },
    { "name": "场景说明", "header_row": 1 }
  ],
  "total_sheets": 4
}
```

#### 2.3.2 变更 `GET /api/init`

新增 `sheets` 概要信息：

```json
{
  "loaded": true,
  "filename": "xxx.xlsx",
  "total": 43,
  "has_active": true,
  "sheet_count": 4,
  "testcase_sheet_count": 2,
  "note_sheet_count": 2
}
```

#### 2.3.3 变更 `GET /api/testcase/<int:index>`

响应中新增 `_sheet_name` 字段，其余不变。

#### 2.3.4 变更 `POST /api/save`

请求体不变。后端保存时从 `STATE['testcases'][index]['_sheet_name']` 获取目标 Sheet 名称（替代原来统一用 `STATE['sheet_name']`），确保结果写回正确的 Sheet。

#### 2.3.5 变更 `GET /api/all-status`

逻辑不变，但需适配多 Sheet：从每条用例的 `_sheet_name` 定位对应 Sheet 读取结果列。

### 2.4 后端核心函数变更

| 函数 | 变更类型 | 说明 |
|------|---------|------|
| `read_testcases(filepath)` | **重构** → `read_all_sheets(filepath)` | 遍历所有 Sheet，分类识别，返回 `LoadedData`（多 Sheet 结构） |
| `_pick_best_sheet(wb)` | **废弃** | 拆分为 `_classify_sheet(ws)` 对单个 Sheet 分类，返回 `SheetMeta` |
| `_classify_sheet(ws)` | **新增** | 对单个 Sheet 评分，判定 `sheet_type` 为 `testcase` 或 `note`，返回 `SheetMeta` |
| `save_result(...)` | **微调** | 参数 `sheet_name` 改为从用例的 `_sheet_name` 获取（调用方传入） |
| `detect_columns` / `_find_header_row` | **不变** | 复用现有逻辑 |

### 2.5 前端变更

本版本前端改动最小，主要是适配：

1. `init()` 读取 `/api/init` 时展示 Sheet 数量信息（topbar 文件名后追加 `(N个Sheet)`）
2. 进度条/统计的 `totalCount` 使用全局用例总数（跨 Sheet 合并）
3. 用例卡片可选展示 `_sheet_name` 标签（小徽标，标识所属 Sheet）

### 2.6 验收标准

| 编号 | 验收项 | 验证方法 |
|------|-------|---------|
| V2.0-1 | 上传含 4 个 Sheet（2 用例 + 2 非用例）的 Excel，`/api/sheets` 正确返回分类 | 接口测试 |
| V2.0-2 | 全局用例列表包含所有用例 Sheet 的用例，`_sheet_name` 字段正确 | 检查 `/api/testcase/0` 响应 |
| V2.0-3 | 单 Sheet 文件（旧场景）仍正常工作，分类为 1 个用例 Sheet | 回归测试 |
| V2.0-4 | 在 Sheet A 的用例上保存结果，结果写入 Sheet A 而非默认 Sheet | 保存后打开 Excel 验证 |
| V2.0-5 | `/api/all-status` 正确返回跨 Sheet 的所有用例状态 | 接口测试 |
| V2.0-6 | Sheet 按原始 Excel 顺序排列 | 检查 `/api/sheets` 响应顺序 |
| V2.0-7 | 全文件无任何用例 Sheet 时，前端显示空状态 | 上传纯非用例文件验证 |

---

## 三、v2.1 — 按 Sheet 分组汇总展示

### 3.1 功能描述

在汇总统计页新增「按 Sheet 分组的用例列表」区域，实现：

1. **按 Sheet 分组**：每个用例 Sheet 作为一个分组，展示该 Sheet 下所有用例
2. **精简字段**：每条用例仅展示 3 个字段：**序号**、**用例名称**、**测试结果**
3. **结果标识**：测试结果用颜色徽标展示（通过=绿/失败=红/阻塞=黄/跳过=灰/未执行=空）
4. **保留原有统计**：顶部统计卡片、环形图、优先级分布表、模块分布表均保留不变
5. **空状态兜底**：
   - 空 Sheet → 该分组内显示「暂无测试用例」
   - 整个文件无用例 → 页面直接显示「暂无测试用例」
   - 某用例缺少结果字段 → 结果列显示为空（不报错）

### 3.2 新增 API

#### `GET /api/summary-by-sheet`

返回按 Sheet 分组的用例精简列表。

**响应示例**：

```json
{
  "total": 43,
  "sheets": [
    {
      "sheet_name": "登录用例",
      "case_count": 25,
      "cases": [
        {
          "global_index": 0,
          "seq": 1,
          "title": "正确账号密码登录",
          "result": "通过"
        },
        {
          "global_index": 1,
          "seq": 2,
          "title": "密码错误登录",
          "result": "失败"
        },
        {
          "global_index": 2,
          "seq": 3,
          "title": "账号为空登录",
          "result": ""
        }
      ]
    },
    {
      "sheet_name": "支付用例",
      "case_count": 18,
      "cases": [
        {
          "global_index": 25,
          "seq": 1,
          "title": "微信支付正常流程",
          "result": "通过"
        }
      ]
    }
  ],
  "empty_sheets": []
}
```

**字段说明**：

| 字段 | 类型 | 说明 |
|------|------|------|
| `global_index` | int | 全局用例索引（用于 v2.2 跳转） |
| `seq` | int | 该 Sheet 内的序号（从 1 开始） |
| `title` | str | 用例名称（取 `title` 字段，无则取 `col_1`/`col_0` 兜底） |
| `result` | str | 测试结果（通过/失败/阻塞/跳过/空字符串=未执行） |

**兜底逻辑**：

- 空 Sheet：`cases` 为空数组，`case_count` 为 0
- 全文件无用例：返回 `{ "total": 0, "sheets": [], "empty": true }`
- 用例无结果字段：`result` 返回空字符串 `""`

### 3.3 数据模型变更

无。本版本复用 v2.0 的多 Sheet 数据结构，仅新增聚合查询接口。

### 3.4 前端变更

#### 3.4.1 汇总统计页布局调整

```
┌─────────────────────────────────────────┐
│  顶部统计卡片（总数/通过/失败/阻塞/跳过）  │  ← 保留不变
├─────────────────────────────────────────┤
│  环形进度图 + 通过率/执行率指标卡         │  ← 保留不变
├─────────────────────────────────────────┤
│  按优先级分布表                          │  ← 保留不变
├─────────────────────────────────────────┤
│  按模块分布表                            │  ← 保留不变
├─────────────────────────────────────────┤
│  【新增】按 Sheet 分组用例列表            │
│  ┌─ Sheet: 登录用例 (25条) ───────────┐  │
│  │ #1  正确账号密码登录    [通过]      │  │
│  │ #2  密码错误登录        [失败]      │  │
│  │ #3  账号为空登录        [  ]        │  │
│  │ ...                                │  │
│  └────────────────────────────────────┘  │
│  ┌─ Sheet: 支付用例 (18条) ───────────┐  │
│  │ #1  微信支付正常流程    [通过]      │  │
│  │ ...                                │  │
│  └────────────────────────────────────┘  │
└─────────────────────────────────────────┘
```

#### 3.4.2 新增前端函数

| 函数 | 职责 |
|------|------|
| `loadSummaryBySheet()` | 调用 `/api/summary-by-sheet`，渲染分组用例列表 |
| `renderSheetGroup(sheetData)` | 渲染单个 Sheet 分组（标题行 + 用例列表） |
| `renderCaseRow(caseData)` | 渲染单条用例行（序号 + 名称 + 结果徽标） |
| `getResultBadgeClass(result)` | 根据结果返回 CSS 类名（pass/fail/block/skip/none） |

#### 3.4.3 新增 CSS 样式

- `.sheet-group`：Sheet 分组容器（卡片样式，带 Sheet 名称标题栏）
- `.sheet-group-header`：分组标题栏（Sheet 名 + 用例数）
- `.case-row`：单条用例行（flex 布局：序号 + 名称 + 结果）
- `.case-seq`：序号样式（monospace 字体）
- `.case-title`：用例名称样式（可点击，v2.2 启用）
- `.case-result-badge`：结果徽标（复用现有 `.result-badge-*` 样式）

### 3.5 验收标准

| 编号 | 验收项 | 验证方法 |
|------|-------|---------|
| V2.1-1 | 汇总页底部出现「按 Sheet 分组用例列表」区域 | UI 检查 |
| V2.1-2 | 每个 Sheet 分组标题显示 Sheet 名 + 用例数 | UI 检查 |
| V2.1-3 | 每条用例仅显示序号、名称、结果三项 | UI 检查 |
| V2.1-4 | 结果用对应颜色徽标展示（通过绿/失败红/阻塞黄/跳过灰） | UI 检查 |
| V2.1-5 | Sheet 按原始 Excel 顺序排列 | 多 Sheet 文件验证 |
| V2.1-6 | 空 Sheet 分组显示「暂无测试用例」 | 上传含空 Sheet 的文件 |
| V2.1-7 | 整个文件无用例时，汇总页显示「暂无测试用例」 | 上传纯非用例文件 |
| V2.1-8 | 用例缺少结果字段时，结果列显示为空（不报错） | 上传无结果列的文件 |
| V2.1-9 | 原有统计卡片、环形图、分布表均正常显示 | 回归测试 |

---

## 四、v2.2 — 用例名称点击跳转

### 4.1 功能描述

在 v2.1 的分组用例列表中，点击任意用例的**名称**，自动：

1. 切换到「执行测试」视图
2. 加载该用例的完整内容到卡片中
3. 滚动定位到该用例卡片
4. 短暂高亮提示（动画 1-2 秒后消失）

跳转依据是 v2.1 接口返回的 `global_index`（全局用例索引），执行测试视图直接用该索引调用 `/api/testcase/<global_index>` 加载。

### 4.2 新增 API

无。本版本纯前端交互，复用 v2.1 返回的 `global_index` 和现有的 `/api/testcase/<index>` 接口。

### 4.3 数据模型变更

无。

### 4.4 前端变更

#### 4.4.1 交互逻辑

```javascript
// 点击用例名称 → 跳转到执行测试视图
function jumpToCase(globalIndex) {
    switchView('execute');          // 切换视图
    loadTestCase(globalIndex);      // 加载该用例
    highlightCard();                // 高亮提示
}
```

#### 4.4.2 新增/修改函数

| 函数 | 职责 |
|------|------|
| `jumpToCase(globalIndex)` | 【新增】点击跳转入口：切视图 + 加载用例 + 高亮 |
| `highlightCard()` | 【新增】给当前用例卡片添加 1-2 秒高亮动画 class |
| `renderCaseRow(caseData)` | 【修改】给用例名称添加 `onclick` 和 `cursor:pointer` 样式 |

#### 4.4.3 新增 CSS

```css
.case-title-clickable {
    cursor: pointer;
    color: var(--primary);
    transition: color .2s;
}
.case-title-clickable:hover {
    color: var(--primary-dark);
    text-decoration: underline;
}
.card-highlight {
    animation: highlightPulse 1.5s ease;
}
@keyframes highlightPulse {
    0%   { box-shadow: 0 0 0 0 rgba(37,99,235,.4); }
    50%  { box-shadow: 0 0 0 8px rgba(37,99,235,.15); }
    100% { box-shadow: 0 0 0 0 rgba(37,99,235,0); }
}
```

### 4.5 验收标准

| 编号 | 验收项 | 验证方法 |
|------|-------|---------|
| V2.2-1 | 鼠标悬停用例名称时显示手型指针和下划线 | UI 检查 |
| V2.2-2 | 点击用例名称后切换到执行测试视图 | 操作验证 |
| V2.2-3 | 卡片内容为被点击的用例（非第一条） | 检查卡片标题与点击项一致 |
| V2.2-4 | 卡片有短暂高亮动画后消失 | UI 检查 |
| V2.2-5 | 跳转后底部导航栏的「上一条/下一条」基于新位置工作 | 连续点击下一条验证 |
| V2.2-6 | 跳转后下拉跳转框选中被点击的用例 | 检查下拉框选中项 |
| V2.2-7 | 有未保存修改时点击跳转，弹出未保存确认弹窗 | 修改后不保存直接点跳转 |

---

## 五、v2.3 — 非用例 Sheet 信息展示

### 5.1 功能描述

在顶部 Tab 栏新增「参考信息」模块，展示所有非用例 Sheet（测试备注、场景说明、测试结果总览等）的原始内容：

1. **新增 Tab**：顶部导航栏在「问题列表」后新增「参考信息」Tab
2. **按 Sheet 分页**：左侧列出所有非用例 Sheet 名称，点击切换查看
3. **原始表格展示**：以 HTML 表格形式展示 Sheet 的原始行列内容（不做字段映射，原样展示）
4. **样式保留**：尽量保留 Excel 中的视觉层次（表头加粗、首行标题等）
5. **空兜底**：无非用例 Sheet 时，该 Tab 隐藏或显示「暂无参考信息」

### 5.2 新增 API

#### `GET /api/note-sheets`

返回所有非用例 Sheet 的原始内容。

**响应示例**：

```json
{
  "sheets": [
    {
      "name": "测试备注",
      "header_row": 1,
      "rows": [
        ["备注类型", "备注内容", "关联用例"],
        ["环境说明", "测试环境：iOS 17.2", "-"],
        ["数据准备", "需预置 3 个测试账号", "登录用例#1-3"]
      ]
    },
    {
      "name": "场景说明",
      "header_row": 1,
      "rows": [
        ["场景编号", "场景描述", "优先级"],
        ["S01", "完整购物流程", "P0"]
      ]
    }
  ],
  "total": 2
}
```

**兜底逻辑**：

- 无非用例 Sheet：返回 `{ "sheets": [], "total": 0 }`
- Sheet 内容为空：`rows` 返回空数组
- 单元格为 None：返回空字符串 `""`

#### `GET /api/note-sheet/<sheet_name>`

返回单个非用例 Sheet 的内容（按需加载，减少首屏数据量）。

**响应示例**：

```json
{
  "name": "测试备注",
  "header_row": 1,
  "rows": [
    ["备注类型", "备注内容", "关联用例"],
    ["环境说明", "测试环境：iOS 17.2", "-"]
  ]
}
```

### 5.3 数据模型变更

无。非用例 Sheet 内容在 v2.0 读取时已分类，本版本仅新增内容查询接口。可选优化：在 `SheetMeta` 中增加 `raw_rows` 字段缓存非用例 Sheet 的原始行数据，避免每次请求重新读 Excel。

```python
@dataclass
class SheetMeta:
    # ... 原有字段
    raw_rows: List[List[str]] = field(default_factory=list)  # 【可选新增】非用例 Sheet 的原始行数据
```

### 5.4 前端变更

#### 5.4.1 顶部 Tab 栏调整

```html
<!-- 现有 -->
<button class="view-tab" id="tabExecute">▶ 执行测试</button>
<button class="view-tab" id="tabSummary">📊 汇总统计</button>
<button class="view-tab" id="tabIssues">📝 问题列表</button>
<!-- 新增 -->
<button class="view-tab" id="tabNotes">📋 参考信息</button>
```

#### 5.4.2 参考信息视图布局

```
┌──────────────────────────────────────────┐
│  侧边栏            │  内容区              │
│  ┌──────────────┐  │  ┌────────────────┐ │
│  │ 测试备注  ●  │  │  │ 表格标题        │ │
│  │ 场景说明     │  │  │ ┌──┬──┬──┐     │ │
│  │ 结果总览     │  │  │ │  │  │  │     │ │
│  └──────────────┘  │  │ └──┴──┴──┘     │ │
│                    │  └────────────────┘ │
└──────────────────────────────────────────┘
```

#### 5.4.3 新增前端函数

| 函数 | 职责 |
|------|------|
| `loadNoteSheets()` | 调用 `/api/note-sheets` 获取 Sheet 列表，渲染侧边栏 |
| `loadNoteSheet(sheetName)` | 调用 `/api/note-sheet/<name>` 加载单个 Sheet 内容 |
| `renderNoteTable(rows)` | 将二维数组渲染为 HTML 表格 |

#### 5.4.4 新增 CSS

- `.note-layout`：左右分栏布局（侧边栏 + 内容区）
- `.note-sidebar`：侧边栏（Sheet 列表）
- `.note-sidebar-item`：Sheet 列表项（选中态高亮）
- `.note-content`：内容区
- `.note-table`：原始数据表格（表头加粗、斑马纹）

### 5.5 验收标准

| 编号 | 验收项 | 验证方法 |
|------|-------|---------|
| V2.3-1 | 顶部 Tab 栏出现「参考信息」按钮 | UI 检查 |
| V2.3-2 | 点击 Tab 进入参考信息视图 | 操作验证 |
| V2.3-3 | 侧边栏列出所有非用例 Sheet 名称 | UI 检查 |
| V2.3-4 | 点击侧边栏 Sheet 名，内容区切换为该 Sheet 内容 | 操作验证 |
| V2.3-5 | 内容以表格形式展示原始行列数据 | UI 检查 |
| V2.3-6 | 表头行加粗显示 | UI 检查 |
| V2.3-7 | 无非用例 Sheet 时，Tab 隐藏或视图显示「暂无参考信息」 | 上传纯用例文件验证 |
| V2.3-8 | 空 Sheet 显示「该 Sheet 暂无内容」 | 上传含空非用例 Sheet 的文件 |
| V2.3-9 | 单元格为空的单元格正常渲染为空白（不报错） | 检查含空单元格的 Sheet |

---

## 六、全局边界条件与向后兼容

### 6.1 向后兼容说明

| 场景 | 兼容策略 |
|------|---------|
| **单 Sheet 文件**（旧场景） | v2.0 读取层将其分类为 1 个用例 Sheet，所有现有功能（执行测试/搜索/保存/汇总）行为不变 |
| **`/api/testcase/<index>`** | 响应新增 `_sheet_name` 字段，原有字段全部保留，前端旧代码不读该字段不影响功能 |
| **`/api/save`** | 请求体不变，后端自动从用例数据中取 `_sheet_name`，调用方无感 |
| **`/api/summary`** | 接口不变，统计范围自动扩展为所有用例 Sheet 的合并数据 |
| **`/api/search`** | 接口不变，搜索范围为所有用例 Sheet 的合并数据 |
| **STATE 兼容字段** | `mapping`/`headers`/`sheet_name`/`header_row` 保留，取第一个用例 Sheet 的值，确保旧代码兼容 |
| **Repository 抽象层** | `TestCaseRepository.load()` 返回值从五元组扩展为 `LoadedData`，`ExcelTestCaseRepository` 和 `InMemoryTestCaseRepository` 同步更新 |

### 6.2 边界条件处理

| 边界场景 | 处理方式 | 涉及版本 |
|---------|---------|---------|
| Excel 文件完全为空（0 个 Sheet） | 前端显示「未读取到用例数据」空状态，与现有行为一致 | v2.0 |
| 所有 Sheet 都是非用例 Sheet | `/api/sheets` 的 `testcase_sheets` 为空，前端执行测试视图显示「暂无测试用例」 | v2.0 |
| 某个用例 Sheet 为空（只有表头无数据行） | 该 Sheet 的 `case_count` 为 0，分组列表中显示「暂无测试用例」 | v2.1 |
| 用例缺少结果列 | `result` 返回空字符串，前端结果徽标不渲染（显示空白） | v2.1 |
| 用例缺少标题字段 | `title` 取 `col_1` → `col_0` → `'(无标题)'` 兜底 | v2.1 |
| 非用例 Sheet 为空 | `rows` 返回空数组，前端显示「该 Sheet 暂无内容」 | v2.3 |
| Sheet 名称含特殊字符（`/`、`&` 等） | URL 编码处理，`/api/note-sheet/<name>` 使用 `quote/unquote` | v2.3 |
| Sheet 名称重复（极少见） | 以 Excel 原始顺序 + 索引后缀区分（如 `Sheet1`、`Sheet1(2)`） | v2.0 |
| 超大文件（1000+ 用例） | `/api/summary-by-sheet` 一次性返回全量，前端虚拟滚动（可选优化，非必须） | v2.1 |
| 保存时 Excel 被占用（权限错误） | 沿用现有 `PermissionError` 处理，提示「请关闭 Excel 后重试」 | v2.0 |

### 6.3 测试用例 Sheet 分类阈值规则（v2.0 核心逻辑）

复用现有 `_pick_best_sheet` 的评分逻辑，拆分为单 Sheet 分类：

| 条件 | 判定结果 |
|------|---------|
| Sheet 名称含「用例/Case/TestCase」等关键词 **且** 核心字段 ≥ 2 | 用例 Sheet |
| Sheet 名称不含关键词 **但** 核心字段 ≥ 3 | 用例 Sheet |
| 核心字段 < 阈值 | 非用例 Sheet |
| Sheet 数据行 < 2（空 Sheet） | 非用例 Sheet |

**核心字段集合**（沿用现有定义）：

```
{id, title, steps, expected, precondition, purpose, priority, module, category, result_col, remark_col}
```

### 6.4 性能考量

| 关注点 | 策略 |
|-------|------|
| 多 Sheet 读取性能 | `read_all_sheets` 一次性遍历所有 Sheet，每个 Sheet 最多读 100 行用于分类（沿用现有 `max_row=min(ws.max_row, 100)` 限制） |
| 非用例 Sheet 内容加载 | v2.3 提供 `/api/note-sheet/<name>` 按需加载单 Sheet，避免首屏拉取全部非用例数据 |
| 内存占用 | 用例数据扁平化合并后内存占用与用例总数成正比，单文件万条用例级别无压力 |

---

## 七、附录

### 7.1 全量 API 清单（v1.x + v2.x）

> v1.x 已有 API 的详细信息见 [1.3 节](#13-v1x-已有-api-清单v134-基线)

| API | 方法 | 引入版本 | 说明 | v2.0 后状态 |
|-----|------|---------|------|------------|
| `/api/init` | GET | v1.0 | 初始化状态 | 变更（新增 Sheet 概要） |
| `/api/titles` | GET | v1.1 | 全部用例标题列表 | 不变 |
| `/api/testcase/<index>` | GET | v1.0 | 单条用例详情 | 变更（新增 `_sheet_name`） |
| `/api/all-status` | GET | v1.0 | 所有用例状态 | 微调（适配多 Sheet） |
| `/api/search` | POST | v1.1 | 搜索筛选 | 不变（自动覆盖全量） |
| `/api/save` | POST | v1.0 | 保存结果 | 微调（Sheet 名从用例取） |
| `/api/summary` | GET | v1.1 | 汇总统计 | 不变（自动覆盖全量） |
| `/api/filter-options` | GET | v1.1 | 筛选选项 | 不变 |
| `/api/reload` | GET | v1.2 | 重新加载 | 微调（调用 `read_all_sheets`） |
| `/api/upload` | POST | v1.2 | 上传文件 | 微调（调用 `read_all_sheets`） |
| `/api/check-results` | GET | v1.3 | 检查已有结果 | 不变 |
| `/api/issues` | GET | v1.3 | 问题列表 | 不变 |
| `/api/sheets` | GET | **v2.0** | 所有 Sheet 分类信息 | — |
| `/api/summary-by-sheet` | GET | **v2.1** | 按 Sheet 分组的用例精简列表 | — |
| `/api/note-sheets` | GET | **v2.3** | 所有非用例 Sheet 列表 | — |
| `/api/note-sheet/<name>` | GET | **v2.3** | 单个非用例 Sheet 内容 | — |

### 7.2 版本验收检查清单

每个版本发布前需确认：

- [ ] 新增 API 单元测试通过
- [ ] 现有 API 回归测试通过（向后兼容）
- [ ] 多 Sheet 文件端到端验证通过
- [ ] 单 Sheet 文件回归验证通过（兼容旧场景）
- [ ] 空文件/空 Sheet 兜底逻辑验证通过
- [ ] `InMemoryTestCaseRepository` 测试夹具同步更新
- [ ] 前端暗色模式适配验证
- [ ] 移动端响应式布局验证

---

**文档结束**
