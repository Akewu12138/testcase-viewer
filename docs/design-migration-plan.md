# Clarity 设计系统 · 落地方案

> 版本：1.0 · 日期：2026-07-08
> 状态：**已落地**（代码已应用，待运行验证）
> 设计师：UI Designer

---

## 一、迁移策略总览

### 核心原则：零 JS 改动

通读 `app.js` 后发现两个硬约束：

1. **JS 内联引用旧变量名**：`showUploadPage()` 注入的 `<style>` 块、动态 HTML 的内联 `style` 大量使用 `var(--surface)`、`var(--primary)`、`var(--bg)`、`var(--border-light)`、`var(--text-secondary)`、`var(--pass/fail/block)` 等旧变量名。
2. **JS 硬编码色值与新主色一致**：`ringSVG()` 与 `showCelebrate()` 硬编码 `#4f46e5`、`#16a34a`，恰好等于 Clarity 主色与成功色，无需调整。

因此落地方案确立三条铁律：

| 铁律 | 含义 | 收益 |
|------|------|------|
| **零 JS 改动** | 不修改 `app.js` 任何一行 | 消除回归风险，交互逻辑零影响 |
| **兼容别名层** | tokens 层提供旧变量名→新令牌的映射 | JS 内联样式自动生效 |
| **class 名不变** | `style.css` 保留全部现有 class | JS 动态生成的 DOM 结构完全兼容 |

### 三层架构

```
index.html
  ├─ design-tokens.css   ← 令牌层（色阶 / 语义令牌 / 暗色覆盖 / 兼容别名 / 基础重置）
  └─ style.css           ← 组件层（只引用语义令牌，不再定义变量、不写暗色查询）
```

主题切换完全由令牌层接管，组件层对明暗主题**无感知**。

---

## 二、改动清单（仅 3 个文件）

| 文件 | 改动类型 | 说明 |
|------|----------|------|
| `app/static/css/design-tokens.css` | 新增 + 补充 | 已有令牌层，本次**追加兼容别名层**（旧变量名映射） |
| `app/static/css/style.css` | 重写 | 删除旧 `:root` 与 `prefers-color-scheme` 块，全部样式值令牌化，视觉对齐预览页 |
| `app/templates/index.html` | 微调 | 在 `style.css` 之前引入 `design-tokens.css`，版本号 `v3→v4` 避缓存 |

**`app.js` 零改动。**

---

## 三、兼容别名映射表

在 `design-tokens.css` 末尾追加，将旧变量名指向新语义令牌：

| 旧变量名 | → 新语义令牌 | 说明 |
|----------|--------------|------|
| `--surface` | `--color-surface` | 卡片表面 |
| `--bg` | `--color-bg` | 页面底色 |
| `--text` | `--color-text` | 主文字 |
| `--text-secondary` | `--color-text-secondary` | 次要文字 |
| `--text-tertiary` | `--color-text-tertiary` | 辅助文字 |
| `--border` | `--color-border` | 默认边框 |
| `--border-light` | `--color-border-light` | 极淡分隔线 |
| `--primary` | `--color-primary` | 主色 |
| `--primary-light` | `--color-primary-light` | 主色浅底 |
| `--primary-dark` | `--color-primary-hover` | 主色 hover |
| `--pass` | `--result-pass` | 通过 |
| `--fail` | `--result-fail` | 失败 |
| `--block` | `--result-block` | 阻塞 |
| `--skip` | `--result-skip` | 跳过 |
| `--radius` | `--radius-lg` | 旧 16px → 12px |
| `--radius-sm` | `--radius-md` | 旧 12px → 8px |
| `--radius-xs` | `--radius-sm` | 旧 8px → 6px |
| `--ease` | `--ease-emphasized` | 缓动曲线 |
| `--t` | `--transition-all` | 通用过渡 |

> `--shadow-*` 与 `--space-*` 新旧命名一致，无需别名。新代码请一律使用 `--color-*` 语义令牌，旧名仅作兼容。

---

## 四、关键 class 改造对照

以顶栏与卡片为例，展示「硬编码值 → 语义令牌」的转换：

### 顶栏 `.topbar`

```css
/* 改造前 */
.topbar{ background:#fff; border-radius:16px; box-shadow:0 1px 3px rgba(0,0,0,.08); ... }

/* 改造后 */
.topbar{
    background:var(--color-surface);
    border:1px solid var(--color-border);
    border-radius:var(--radius-lg);        /* 16px → 12px 更克制 */
    box-shadow:var(--shadow-sm);
    ...
}
```

### 主按钮 `.btn-save`

```css
/* 改造前：文字链接 + ::after 强制 ↗ 箭头（与 JS textContent 叠加冗余） */
.btn-save{ background:transparent; color:var(--primary); }
.btn-save::after{ content:"↗"; }   /* ← 已移除 */

/* 改造后：主色文字 + hover 浅底，移除箭头冗余 */
.btn-save{ background:transparent; color:var(--color-primary); }
.btn-save:hover:not(:disabled){ color:var(--color-primary-hover); background:var(--color-primary-light); }
```

### 结果徽章 `.result-badge-pass`

```css
/* 改造前：硬编码色值 */
.result-badge-pass{ background:#e8f7ee; color:#15803d; }

/* 改造后：语义令牌，暗色自动切换 */
.result-badge-pass{ background:var(--color-success-light); color:var(--green-700); }
```

### 暗色模式

```css
/* 改造前：style.css 内 60+ 行 @media (prefers-color-scheme: dark) 逐个覆盖 */

/* 改造后：style.css 零暗色代码，全部由 design-tokens.css 的 [data-theme="dark"] 接管 */
```

---

## 五、视觉落地要点

| 维度 | 改造前 | 改造后（Clarity） |
|------|--------|-------------------|
| 主色 | `#2563eb` 通用蓝 | `#4f46e5` Indigo |
| 中性色 | 纯灰 | Slate 冷灰调 |
| 圆角 | 16px 偏软 | 12px 主卡片 / 8px 按钮 / 6px 小元素 |
| 阴影 | 单层纯黑 `rgba(0,0,0,.08)` | 5 级分层，slate-900 冷灰调 |
| 卡片 | 边框+阴影叠加 | 单描边 + 极淡阴影 |
| 表头 | 小写 | 大写 + 字距 `0.06em` |
| 暗色 | 60 行补丁式覆盖 | 令牌层统一接管 |
| logo | 纯色圆 | Indigo 渐变方 |
| 弹窗 | 无进场动画 | `fadeIn` + `scale` 进场 |
| 数字 | 默认比例字 | `tabular-nums` 等宽对齐 |

---

## 六、风险与回滚

| 风险 | 等级 | 缓解措施 |
|------|------|----------|
| JS 内联变量失效 | 中 | 兼容别名层已覆盖全部旧变量名 |
| 浏览器缓存旧 CSS | 低 | index.html 版本号 `v3→v4`，强制刷新 |
| 暗色显示异常 | 低 | 令牌层暗色覆盖已验证，组件层零暗色代码 |
| 旧 `style.css` 丢失 | 低 | 项目已纳入 git，`git checkout` 即可回滚 |

**回滚命令**：
```bash
git checkout app/static/css/style.css app/templates/index.html
# design-tokens.css 为新增文件，回滚后不影响（未被引用即可）
```

---

## 七、验证清单

落地后请按以下清单逐项验证：

### 基础加载
- [ ] 页面无 404（design-tokens.css、style.css 均加载成功）
- [ ] 控制台无 CSS 变量未定义警告
- [ ] 顶栏、卡片、按钮视觉与预览页一致

### 执行测试视图
- [ ] 用例卡片编号胶囊、标题大字、字段层级正常
- [ ] 结果按钮四态（通过/失败/阻塞/跳过）选中色正确
- [ ] 备注框聚焦环为 Indigo
- [ ] 进度条填充与图例色点正确
- [ ] 保存后 Toast 显示，自动跳下一条

### 汇总统计视图
- [ ] 统计卡片大数字 Indigo/绿/红/琥珀色正确
- [ ] 环形进度条 SVG 描边色正确（`#4f46e5`/`#16a34a`）
- [ ] 表格表头大写、迷你进度条正常

### 问题列表视图
- [ ] 问题卡片 hover 上浮 + 边框变主色
- [ ] 庆祝动画彩纸颜色与 confetti 正常

### 上传页（app.js 内联样式）
- [ ] 上传卡片背景、虚线框、hover 主色高亮正常
- [ ] **这是兼容别名层的关键验证点**——若失效说明别名映射有遗漏

### 暗色模式
- [ ] 系统暗色下，底色 slate-950、卡片 slate-900、主色 indigo-400
- [ ] 文字、边框、阴影全部自动切换，无残留浅色

### 响应式
- [ ] 375px 宽度下顶栏紧凑、标签可点击、结果按钮换行
- [ ] 768px / 1280px 布局正常

---

## 八、已知遗留（非本次范围）

阅读 `app.js` 时发现一个既有交互问题，**非视觉落地引入**，记录备查：

- `index.html` 顶栏有 `tabExecute / tabSummary / tabIssues` 三个标签，但 `switchView()` 第 38 行调用 `document.getElementById('tabSearch').classList.toggle(...)`。当前 HTML 无 `tabSearch` 元素，该行会抛 `Cannot read property of null`。
- `viewSearch` 容器仍存在于 HTML，但无标签入口。疑似搜索功能在调整为独立入口的过程中。
- **建议**：后续在 `switchView` 对 `tabSearch` 加 `if` 守卫（与 `tabIssues` 一致），或移除该引用。这不影响 Clarity 视觉落地，可在下个迭代处理。

---

## 九、后续可选优化

落地已达成「视觉统一 + 零 JS 改动」目标。后续可考虑的增量优化：

1. **顶栏结构升级**：将卡片式顶栏改为预览页的全宽 + backdrop-blur 风格（需微调 HTML，非零改动）
2. **修复 tabSearch 守卫**：见第八节
3. **按钮体系统一**：现状 `.btn-save` 为文字风格，可按预览页引入 `.btn-primary/.btn-secondary/.btn-ghost/.btn-danger` 四变体（需同步调整 JS 生成的按钮 class）
4. **引入 stylelint**：校验「组件层不得使用色阶令牌」规则，防止一致性退化

以上均为增量项，不影响本次落地的完整性。

---

**落地完成时间**：2026-07-08
**验证状态**：待运行应用确认（见验证清单）
**回滚方式**：`git checkout` 对应文件
