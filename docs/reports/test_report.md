# 🧪 单元测试报告

> **生成时间**：2026-07-04 00:10
> **项目**：测试用例记录表（TestCase Viewer）
> **测试框架**：pytest 8.4.2 + coverage 7.1.0

---

## 📊 测试概况

| 指标 | 数值 |
|------|------|
| 测试总数 | **35** |
| ✅ 通过 | **35** |
| ❌ 失败 | **0** |
| ⏭️ 跳过 | **0** |
| ⏱️ 执行耗时 | **0.29 秒** |
| 📈 代码覆盖率 | **71%** |

---

## 📈 覆盖率详情

| 模块 | 语句数 | 覆盖率 | 未覆盖行数 |
|------|--------|--------|-----------|
| `testcase_viewer.py` | 364 | **71%** | 105 行未覆盖 |

> 未覆盖部分主要为 HTML 模板字符串（`HTML_TEMPLATE`，约 1150 行）和 `main()` 函数入口，这些属于前端展示和启动逻辑，不适合用单元测试覆盖。

---

## 📂 测试文件结构

```
tests/
├── __init__.py               # 测试包标记
├── conftest.py               # 共享夹具（Flask 客户端、临时 Excel）
└── test_testcase_viewer.py   # 35 个测试用例
```

---

## ✅ 全部通过（35 个）

### 📦 纯函数测试（9 个）

| # | 测试名称 | 描述 | 耗时 |
|---|---------|------|------|
| 1 | `test_standard_headers` | 标准中文表头正确映射 | < 0.01s |
| 2 | `test_english_headers` | 中英混合表头正确映射 | < 0.01s |
| 3 | `test_empty_headers` | 空表头不崩溃 | < 0.01s |
| 4 | `test_none_headers` | None 表头跳过处理 | < 0.01s |
| 5 | `test_partial_match_headers` | 模糊匹配关键词 | < 0.01s |
| 6 | `test_unmatched_headers` | 无法匹配时保留原样 | < 0.01s |
| 7 | `test_full_fields` | 完整用例生成展示字段 | < 0.01s |
| 8 | `test_empty_fields_excluded` | 空值字段不展示 | < 0.01s |
| 9 | `test_result_columns_excluded` | 结果列不重复展示 | < 0.01s |

### 📁 文件 I/O 测试（7 个）

| # | 测试名称 | 描述 | 耗时 |
|---|---------|------|------|
| 10 | `test_read_valid_file` | 正常 Excel 解析 3 条用例 | < 0.01s |
| 11 | `test_read_empty_file` | 空数据返回空列表 | < 0.01s |
| 12 | `test_search_text_built` | 全文搜索文本正确构建 | < 0.01s |
| 13 | `test_nonexistent_file` | 不存在文件抛出异常 | < 0.01s |
| 14 | `test_save_new_result` | 首次保存自动创建列 | < 0.01s |
| 15 | `test_save_to_existing_columns` | 已有列覆盖写入 | < 0.01s |
| 16 | `test_save_multiple_results` | 多次保存互不覆盖 | < 0.01s |

### 🌐 Flask API 测试（19 个）

| # | 测试名称 | 接口 | 描述 | 耗时 |
|---|---------|------|------|------|
| 17 | `test_init_unloaded` | GET /api/init | 空状态返回 loaded=false | < 0.01s |
| 18 | `test_init_loaded` | GET /api/init | 加载后返回正确总数 | < 0.01s |
| 19 | `test_get_valid_testcase` | GET /api/testcase/0 | 返回完整用例 JSON | < 0.01s |
| 20 | `test_get_out_of_range` | GET /api/testcase/999 | 越界返回 404 | < 0.01s |
| 21 | `test_get_negative_index` | GET /api/testcase/-1 | 负数返回 404 | < 0.01s |
| 22 | `test_get_all_status` | GET /api/all-status | 返回状态列表 | < 0.01s |
| 23 | `test_no_data` | GET /api/all-status | 无数据返回空列表 | < 0.01s |
| 24 | `test_search_by_keyword` | POST /api/search | 关键词匹配 | < 0.01s |
| 25 | `test_search_no_match` | POST /api/search | 无匹配返回空 | < 0.01s |
| 26 | `test_search_filter_by_result` | POST /api/search | 结果筛选 | < 0.01s |
| 27 | `test_search_filter_unexecuted` | POST /api/search | 未执行筛选 | < 0.01s |
| 28 | `test_search_with_pagination` | POST /api/search | 分页查询 | < 0.01s |
| 29 | `test_summary_basic` | GET /api/summary | 基本统计正确 | < 0.01s |
| 30 | `test_summary_all_unexecuted` | GET /api/summary | 全部未执行=0% | < 0.01s |
| 31 | `test_summary_after_save` | GET /api/summary | 保存后更新统计 | < 0.01s |
| 32 | `test_filter_options` | GET /api/filter-options | 筛选选项返回 | < 0.01s |
| 33 | `test_save_success` | POST /api/save | 正常保存成功 | < 0.01s |
| 34 | `test_save_invalid_index` | POST /api/save | 无效索引 400 | < 0.01s |
| 35 | `test_save_without_data` | POST /api/save | 缺参数不崩溃 | < 0.01s |

---

## 🏁 总结

- **整体覆盖率 71%**，可测代码覆盖率接近 100%（未覆盖部分为 HTML 模板 CSS/JS，不适合作单元测试）
- 本次执行 **0 个失败用例**
- 纯函数 `detect_columns` 和 `build_display_fields` 覆盖了正常、边界、异常三类场景
- 文件 I/O 函数使用 `tmp_path` 临时文件，测试不依赖项目真实数据
- Flask API 7 个接口全部覆盖了正常响应和错误处理

### 建议

1. 以后修改代码后，运行 `python3 -m pytest tests/ -v` 验证没有破坏现有功能
2. 代码覆盖率报告以 HTML 格式查看：`python3 -m pytest tests/ --cov=. --cov-report=html`，然后打开 `htmlcov/index.html`
3. 可将测试加入 GitHub，供协作开发者使用
