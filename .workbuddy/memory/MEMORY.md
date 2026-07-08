# 项目记忆 · TestCase Viewer

## 设计系统（Clarity）
- 2026-07-08 制定，命名为「Clarity · 清晰」
- 配色：Slate 冷灰中性色 + Indigo 主色（#4f46e5），区别于通用蓝 #2563eb
- 架构：令牌三层分层（原始色阶 → 语义令牌 → 组件），换肤只改变量
- 字号基准 14px，8 点栅格（4px 基准）
- 圆角收窄至 6-12px，保持工具感（原 16px 过软）
- 阴影用 slate-900 冷灰调而非纯黑
- 令牌文件：app/static/css/design-tokens.css
- 预览页：docs/design-preview.html（自包含，可独立打开查看效果）

## 框架选型
- 方案 A 优先：原生 CSS + design-tokens.css，无构建依赖
- 理由：项目「双击即用」部署（start.command/start.bat），不引入 Node 构建链路
- 组件激增（页面≥10）时再考虑 Tailwind CSS

## 项目约定
- 技术栈：Flask + 原生 HTML/CSS/JS，轻量工具型应用
- 端口 8765，数据目录 testcases/，读首个 .xlsx
- 中文注释
- 三个视图：执行测试 / 汇总统计 / 问题列表（原搜索筛选已合并）
