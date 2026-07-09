---
name: code-standards
description: 编写代码时自动检查编码规范——Conventional Commits、make lint、pre-commit hook、代码风格统一
agent_created: true
triggers:
  - 用户要求编写或修改代码时
  - 用户说"帮我写一个XX功能"时
  - 修改文件后准备提交前
---

# 编码规范 Skill（vibe-coding 五层金字塔 第2层）

> 目标：让代码长得一样、提交记录可追溯、防呆卡在提交前。

## 工作流程

### 写代码前

1. **确认要求**：用户要改什么？哪个文件？
2. **看现有代码风格**：缩进（4空格 / 2空格）、命名（驼峰 / 下划线）、注释风格
3. **遵循现有风格**：不和项目现有代码风格冲突

### 写代码中

1. **Python 代码**：
   - 缩进 4 空格（EditorConfig 已设置）
   - 函数/变量用 `snake_case`
   - 类名用 `PascalCase`
   - 函数加简短注释（用中文说明做什么）
   - 复杂逻辑加行内注释

2. **JS 代码**：
   - 缩进 4 空格
   - 变量/函数用 `camelCase`
   - 每个语句末尾加分号 `;`
   - 字符串用单引号 `'`

3. **HTML/CSS**：
   - 缩进 4 空格
   - class 名用 `kebab-case`（如 `case-card`）
   - CSS 属性按逻辑分组（布局 → 尺寸 → 外观 → 动效）

### 提交前（必须执行，不通过不提交）

按顺序执行以下检查：

| 步骤 | 命令 | 作用 |
|------|------|------|
| 1. JS 语法 | `node --check app/static/js/app.js` | 抓 JS 语法错误（B4 那种重复 catch 块） |
| 2. Python 语法 | `python3 -m py_compile testcase_viewer.py` | 抓 Python 语法错误 |
| 3. 跑测试 | `make test` | 确认 52 passed 全绿，代码改动不回归 |

**任何一步失败都阻止提交。**

### 提交信息格式（Conventional Commits）

```
<type>: <简短描述>

- 改动 1
- 改动 2
```

| type | 含义 | 示例 |
|------|------|------|
| `feat` | 新功能 | `feat: 加批量导入功能` |
| `fix` | 修 bug | `fix: 搜索关键词高亮不生效` |
| `style` | 样式改动 | `style: 按钮圆角统一 12px` |
| `refactor` | 重构（功能不变） | `refactor: 拆 testcase_viewer 到 services/` |
| `docs` | 文档 | `docs: 更新 PRD` |
| `chore` | 杂项（配置/构建等） | `chore: 复制 skills 到项目` |
| `test` | 测试 | `test: 补 API 路由测试` |

## 关键规则

1. **不改变现有代码风格**：每个项目有自己的风格，入乡随俗
2. **提交前必跑检查**：不通过不提交（pre-commit hook 自动拦截）
3. **小步提交**：一次提交只做一件事，方便回退
4. **改动后必验证**：跑 `make test` 确认全绿
