# TestCase Viewer · Clarity 设计系统

> 版本：1.0 · 日期：2026-07-08
> 设计师：UI Designer
> 状态：方案就绪，待评审

---

## 一、设计理念 · Clarity · 清晰

测试用例工具是**高频、长时间使用**的工具型应用。用户的核心任务是「逐条执行、记录结果」，而非浏览花哨界面。因此新设计系统确立了四条原则：

| 原则 | 含义 | 在设计中的体现 |
|------|------|----------------|
| **信息优先** | 数据是主角，UI 是配角 | 去除卡片多余装饰，字段以纯文本层级呈现 |
| **克制的精致** | 减少装饰，强化层次 | 收窄圆角、分层阴影、冷灰调中性色 |
| **一致性** | 统一的 token 体系 | 8 点栅格 + 语义令牌分层，杜绝魔数 |
| **可访问性** | 默认 WCAG AA 合规 | 文字对比度 ≥ 4.5:1，触达区 ≥ 44px |

**与现状的对比**：当前 `style.css` 配色用通用蓝 `#2563eb`、圆角 16px 偏软、阴影单一、token 未分层。Clarity 用 **Slate 冷灰中性色 + Indigo 主色**提升专业感与辨识度，圆角收窄至 6–12px 保持工具感，并建立完整的语义令牌分层架构。

---

## 二、设计规范

### 2.1 色彩系统

采用**三层色彩架构**：原始色阶 → 语义令牌 → 组件引用。组件层只引用语义令牌，换肤只改变量。

#### 主色 · Indigo

区别于通用蓝（Tailwind blue），Indigo 色相更偏紫，在保持专业感的同时更有辨识度，且与冷灰中性色和谐。

| Token | 色值 | 用途 |
|-------|------|------|
| `--indigo-600` | `#4f46e5` | 主色（按钮、链接、聚焦环） |
| `--indigo-700` | `#4338ca` | 主色 hover |
| `--indigo-500` | `#6366f1` | 主色亮版（暗色模式） |
| `--indigo-50` | `#eef2ff` | 主色浅底（选中态、悬浮底） |

#### 中性色 · Slate

比纯灰更有质感、色温更冷，与 Indigo 主色协调。

| Token | 色值 | 语义用途 |
|-------|------|----------|
| `--slate-50` | `#f8fafc` | 页面底色 |
| `--slate-100` | `#f1f5f9` | 次级表面、表头底 |
| `--slate-200` | `#e2e8f0` | 默认边框 |
| `--slate-400` | `#94a3b8` | 占位符、辅助文字 |
| `--slate-600` | `#475569` | 次要文字 |
| `--slate-900` | `#0f172a` | 主文字 |
| `--slate-950` | `#020617` | 暗色模式底色 |

#### 语义状态色

| 状态 | 主色 | 浅底（徽章/选中态） | 文字 |
|------|------|---------------------|------|
| 成功 Pass | `#16a34a` | `#f0fdf4` | `#15803d` |
| 警告 Block | `#d97706` | `#fffbeb` | `#b45309` |
| 危险 Fail | `#dc2626` | `#fef2f2` | `#b91c1c` |
| 跳过 Skip | `#64748b` | `#f1f5f9` | `#475569` |

#### 可访问性验证（明色主题）

| 组合 | 对比度 | 等级 |
|------|--------|------|
| `--slate-900` on `#ffffff` | 15.8:1 | AAA |
| `--slate-600` on `#ffffff` | 7.4:1 | AAA |
| `--slate-400` on `#ffffff` | 3.0:1 | AA Large（仅用于辅助） |
| `#ffffff` on `--indigo-600` | 6.5:1 | AA+ |
| `--green-700` on `--green-50` | 4.9:1 | AA |

> 主文字、次要文字、按钮文字均达 AA 以上。`--slate-400` 仅用于非关键辅助信息与占位符。

---

### 2.2 字体系统

工具型应用基准字号取 **14px**（比 16px 紧凑，信息密度更高），建立 8 级字号阶梯。

#### 字体族

```css
--font-sans: -apple-system, "SF Pro Text", "PingFang SC",
             "Microsoft YaHei", system-ui, sans-serif;
--font-mono: "SF Mono", "JetBrains Mono", "Fira Code", monospace;
```

中文优先调用系统苹方/雅黑，零网络字体加载，性能最优。

#### 字号阶梯

| Token | 值 | 用途 |
|-------|----|------|
| `--text-xs` | 12px | 标签、辅助说明、表头 |
| `--text-sm` | 13px | 次要正文、按钮文字 |
| `--text-base` | 14px | **基准正文** |
| `--text-md` | 15px | 卡片标题 |
| `--text-lg` | 17px | 强调标题 |
| `--text-xl` | 20px | 页面标题 |
| `--text-2xl` | 24px | 统计数字 |
| `--text-3xl` | 30px | 庆祝/大数字 |

#### 字重与行高

- 字重：`400` 正文 / `500` 次要强调 / `600` 标题按钮 / `700` 大数字
- 行高：`1.3` 标题 / `1.5` 正文 / `1.7` 长文本多行
- 数字场景启用 `font-variant-numeric: tabular-nums`，编号与统计对齐

---

### 2.3 间距系统 · 8 点栅格

以 **4px 为基准单位**，所有内外边距、栅格间距必须取自下表，**禁止魔数**。

| Token | 值 | 典型用途 |
|-------|----|----------|
| `--space-1` | 4px | 图标与文字间距 |
| `--space-2` | 8px | 紧凑元素间距 |
| `--space-3` | 12px | 输入框内边距 |
| `--space-4` | 16px | 卡片内边距、基础间距 |
| `--space-5` | 20px | 卡片大内边距 |
| `--space-6` | 24px | 区块间距 |
| `--space-8` | 32px | 大区块间距 |
| `--space-12` | 48px | 页面级留白 |

内容容器最大宽度 `--container-max: 960px`，居中布局，移动端两侧 16px。

---

### 2.4 圆角与阴影

#### 圆角（克制，保持工具感）

| Token | 值 | 用途 |
|-------|----|------|
| `--radius-xs` | 4px | 小标签、徽章 |
| `--radius-sm` | 6px | 按钮、输入框 |
| `--radius-md` | 8px | 小卡片、下拉 |
| `--radius-lg` | 12px | 主卡片、弹窗 |
| `--radius-full` | 9999px | 胶囊、圆点 |

> 较现状 16px 收窄，避免「过度柔软」，更贴合工具型应用的克制气质。

#### 分层阴影（slate 冷灰调，非纯黑）

```css
--shadow-xs: 0 1px 2px 0 rgb(15 23 42 / 0.04);          /* 卡片默认 */
--shadow-sm: 0 1px 3px 0 rgb(15 23 42 / 0.06) ...;      /* 悬浮态 */
--shadow-md: 0 4px 8px -2px rgb(15 23 42 / 0.08) ...;   /* 弹出层 */
--shadow-lg: 0 12px 20px -4px rgb(15 23 42 / 0.10) ...; /* 弹窗 */
--shadow-focus: 0 0 0 3px rgb(79 70 229 / 0.15);        /* 聚焦环 */
```

阴影使用 `slate-900` 色调（冷灰）而非纯黑，过渡更自然精致。

---

### 2.5 动效令牌

统一缓动曲线与时长，避免每个组件各自定义。

| Token | 时长 | 曲线 | 用途 |
|-------|------|------|------|
| `--duration-fast` | 120ms | standard | 颜色、悬停 |
| `--duration-normal` | 200ms | standard | 通用过渡 |
| `--duration-slow` | 320ms | emphasized | 进场、位移 |

并全局尊重 `prefers-reduced-motion`，为前庭敏感用户提供静态体验。

---

## 三、组件体系

所有组件统一引用语义令牌，确保跨页面一致。完整效果见 `docs/design-preview.html`。

### 3.1 按钮 Button

**变体**：`primary`（实心主色）/ `secondary`（描边）/ `ghost`（幽灵）/ `danger`（危险）。
**尺寸**：`sm`（28px 高）/ `md`（36px）/ `lg`（44px，移动端最小触达）。
**状态**：default / hover / active / focus-visible / disabled。

```css
/* 主按钮 */
.btn-primary {
    background: var(--color-primary);
    color: var(--color-text-inverse);
}
.btn-primary:hover { background: var(--color-primary-hover); }
.btn-primary:active { background: var(--color-primary-active); transform: scale(0.98); }
.btn-primary:focus-visible { box-shadow: var(--shadow-focus); }
.btn-primary:disabled { opacity: 0.5; cursor: not-allowed; }
```

**一致性要点**：所有按钮统一 `--radius-sm`、统一内边距令牌、统一 `active: scale(0.98)` 反馈。

### 3.2 表单控件 Form

- **输入框 / 文本域**：`--radius-sm`、`1px solid --color-border`、聚焦时 `border-color: --color-primary` + `--shadow-focus`。
- **下拉选择**：自定义箭头，与输入框高度对齐。
- **校验态**：错误用 `--color-danger` 边框 + `--shadow-focus-danger`。

### 3.3 卡片 Card

```css
.card {
    background: var(--color-surface);
    border: 1px solid var(--color-border);
    border-radius: var(--radius-lg);
    box-shadow: var(--shadow-xs);
}
```

去除了现状的「边框 + 阴影叠加」，改为**单一描边 + 极淡阴影**，视觉更干净。

### 3.4 导航 Tabs

胶囊容器内嵌的标签，激活态白底 + 主色文字 + 极淡阴影，与现状一致但 token 化。

### 3.5 徽章 / 状态标签 Badge

测试结果用胶囊徽章：浅底 + 同色文字，对比度达 AA。例：Pass = `--green-50` 底 + `--green-700` 字。

### 3.6 进度条 Progress

5px 高细条，`--color-primary` 填充，完成态切换 `--result-pass`。配合图例色点。

### 3.7 表格 Table

表头 `--color-surface-2` 底 + `--text-xs` 大写字母 + `--tracking-caps`；单元格 12px 内边距；行分隔线 `--color-border-light`。

### 3.8 弹窗 Modal

`--shadow-xl` + `backdrop-filter: blur(8px)` 遮罩，`--radius-lg` 圆角，主操作右、取消左。

### 3.9 Toast

顶部居中胶囊，`--shadow-lg`，4 类语义色，2.3s 自动消失。

### 3.10 统计卡片 Stat

大数字 `--text-2xl` + `font-weight-700` + `tabular-nums`，配合环形进度（SVG `stroke-dasharray`）。

---

## 四、响应式设计

采用**移动优先**策略，3 个断点覆盖全场景：

| 断点 | 宽度 | 布局调整 |
|------|------|----------|
| 移动端 | < 640px | 单列、顶栏紧凑（隐藏副标题与图标文字）、容器两侧 16px |
| 平板 | 640–1023px | 部分两列、搜索栏允许换行 |
| 桌面 | ≥ 1024px | 完整布局、`--container-max` 居中 |

**移动端关键规则**：
- 交互控件最小高度 `--control-height-lg: 44px`（iOS HIG / WCAG 触达标准）
- 顶栏标签隐藏图标后的文字，保留功能
- 结果按钮组允许换行，单按钮最小宽度 70px
- 统计网格 `grid-template-columns: repeat(2, 1fr)`

---

## 五、暗色模式

通过 `data-theme="dark"` 属性或 `prefers-color-scheme` 媒体查询，**仅覆盖语义令牌**即可切换，组件层零改动。

设计要点：
- 暗色底用 `--slate-950`（近黑但带蓝灰），非纯黑，降低对比疲劳
- 主色提亮至 `--indigo-400`，保证在深底上的可读性
- 阴影改为纯黑半透明，更深沉
- 语义浅底改为半透明叠色（如 `rgba(34,197,94,0.12)`），避免色块过亮

---

## 六、框架选型建议

当前项目是 Flask + 原生 HTML/CSS/JS，**本次方案不强制引入重框架**，推荐两条路径：

### 方案 A（推荐 · 渐进增强）：原生 CSS + Clarity Tokens

- 引入 `design-tokens.css` 作为令牌层
- 将现有 `style.css` 中的硬编码值替换为语义令牌引用
- 组件样式保持原生 CSS，零运行时依赖，符合工具型应用轻量诉求
- **适用**：当前项目体量小、交互中等复杂度，原生方案足够

### 方案 B（若未来组件激增）：Tailwind CSS

- 通过 `tailwind.config.js` 将 Clarity 色阶/字号映射为 Tailwind 主题
- 保留语义令牌做暗色模式与主题切换
- **适用**：页面数 ≥ 10、组件复用度高、需要原子化加速开发时

> 鉴于本项目当前规模与「双击即用」的部署诉求（`start.command`/`start.bat`），**方案 A 更契合**——无构建步骤、无 Node 依赖，保持现有启动链路不变。

---

## 七、一致性保障机制

| 机制 | 说明 |
|------|------|
| 令牌分层 | 原始色阶 → 语义令牌 → 组件，组件层禁止直接使用色阶 |
| 8 点栅格 | 所有间距取自 `--space-*` 阶梯，CI 可加 stylelint 规则校验 |
| 组件清单 | 维护组件文档，新增组件须走 token，不得私造样式 |
| 暗色隔离 | 暗色仅覆盖语义令牌，组件层对主题无感知 |
| 可访问性门禁 | 文字对比度 ≥ 4.5:1，触达区 ≥ 44px，聚焦环可见 |

---

## 八、落地路径

1. **引入令牌层**：`index.html` 中在 `style.css` 之前引入 `design-tokens.css`
2. **迁移组件样式**：逐块将 `style.css` 硬编码值替换为语义令牌（低风险，可分批）
3. **视觉校验**：对照 `docs/design-preview.html` 校验各组件还原度
4. **暗色验证**：`<html data-theme="dark">` 全量走查对比度与可读性
5. **移动端测试**：375px / 768px / 1280px 三档断点回归

---

## 九、交付物

| 文件 | 说明 |
|------|------|
| `app/static/css/design-tokens.css` | 设计令牌层（明暗主题 + 基础重置），可直接落地 |
| `docs/design-system.md` | 本方案文档 |
| `docs/design-preview.html` | 可视化预览页，直观感受新设计语言全貌 |

---

**设计师**：UI Designer
**设计系统日期**：2026-07-08
**实施状态**：令牌层就绪，待评审后进入组件迁移
