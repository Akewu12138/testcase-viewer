---
name: self-testing
description: 写测试、跑测试、检查覆盖率——代码改动后自动验证不回归（vibe-coding 五层金字塔 第3层 核心）
agent_created: true
triggers:
  - 用户写完代码后要求"测试一下"或"检查一下"
  - 代码改动后
  - 用户说"帮我写测试"
  - 修改了 testcase_viewer.py 或 app/ 下的文件后
---

# 自测 Skill（vibe-coding 五层金字塔 第3层）

> 目标：改代码不回归——测试红了就是改坏了，测试绿了才安全。

## 测试架构

本项目测试分层（当前状态）：

| 层 | 工具 | 文件 | 状态 |
|---|------|------|------|
| 后端单元测试 | pytest | `tests/test_testcase_viewer.py` | ✅ 43 个用例 |
| Repository 接口测试 | pytest | `tests/test_in_memory_repo.py` | ✅ 9 个用例 |
| 前端语法检查 | node --check | `app/static/js/app.js` | ✅ 零依赖 |
| E2E 测试 | — | — | ❌ 未配置 |

## 工作流程

### 改代码后（必须执行）

```bash
make test
```

**预期输出**：
```
52 passed in 0.38s
TOTAL 646 196 70%
```

**如果测试红了**：
1. 看错误信息，确认是哪个测试失败
2. 判断是"代码改坏了"还是"测试需要更新"：
   - **代码改坏了**：修代码，重新跑测试
   - **测试需要更新**（函数签名变了但功能正确）：先用 git diff 确认改动范围，再更新测试
3. 修复后重新 `make test`，确认全绿

**如果测试全绿**：
1. 确认覆盖率没有下降（当前 70%）
2. 可以提交了

### 写新功能前（推荐）

1. 先看现有测试覆盖了哪些路径（参考 `tests/test_testcase_viewer.py`）
2. 新功能关键路径要加测试

### 写测试的原则

1. **用 Flask test client**：`app.test_client()` 模拟 HTTP 请求
2. **用 `populated_state` fixture**：模拟已加载 Excel 的状态
3. **用 `app_client` fixture**：创建测试客户端
4. **用 `monkeypatch` mock 外部依赖**：如 `find_excel_file`、`read_memory` 等
5. **测试正常路径 + 异常路径**：
   - 正常：输入正确 → 返回 200 + 正确数据
   - 异常：输入错误 → 返回 400/404/500 + 错误信息

### 测试模板

```python
class Test{功能名}:
    """测试 {功能}"""

    def test_{正常场景}(self, app_client, populated_state):
        """{描述} → 返回正确"""
        resp = app_client.{get/post}('{路径}')
        assert resp.status_code == 200
        data = resp.get_json()
        assert {关键字段} == {期望值}

    def test_{异常场景}(self, app_client):
        """{描述} → 返回错误"""
        resp = app_client.{get/post}('{路径}')
        assert resp.status_code in (400, 404)
```

## 当前项目测试要点

### Python 环境

- 用 `/usr/bin/python3`（有 pytest 8.4.2）
- Makefile 里 `PYTHON ?= /usr/bin/python3`
- managed Python 3.13 网络不通，不能用

### 测试注意事项

1. **导入副作用**：`testcase_viewer.py` 顶层会执行 `_ensure_deps()` + 创建 Flask app，测试里必须 mock `webbrowser.open`
2. **read_testcases 返回 5 个值**：`testcases, mapping, headers, sheet_name, header_row`
3. **save_result 签名**：`save_result(filepath, row_number, header_row_idx, sheet_name, result, actual_result, tester, bug_id, bug_frequency, issue_time)`
4. **STATE 是全局字典**：测试用 `populated_state` fixture 管理
5. **testcases/xlsx 文件可能被测试修改**：注意隔离测试数据

## 关键规则

1. **改代码后必跑 `make test`**：不过不提交
2. **测试红了先判断是代码改坏了还是测试要更新**：不要盲目修测试掩盖 bug
3. **新功能关键路径要加测试**：覆盖率只升不降
4. **前端 JS 改动后跑 `node --check`**：零依赖抓语法错误
5. **Mock 外部依赖时注意返回值类型**：如 `find_excel_file` 需要用 `monkeypatch` 替换为返回 Path 对象的函数
