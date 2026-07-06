# Frontend Style Guide

## Description

当前项目的前端设计规范助手。当用户提出任何与前端 UI、样式、布局、颜色、设计、美化、效果、优化、风格相关的请求时，自动加载本 Skill。

## Usage

本 Skill 用于确保项目前端代码和视觉效果统一。在给出方案或修改前端代码之前，必须先读取项目根目录下的 `FRONTEND_STYLE_GUIDE.md` 设计规范文件。

## Triggers

当用户消息包含以下任意关键词时触发：

- 前端
- UI
- 样式
- 布局
- 颜色
- 配色
- 设计
- 美化
- 效果
- 优化
- 风格
- 界面
- 组件
- 按钮
- 卡片
- 表单
- 表格
- 主题
- 暗色模式
- 响应式

## Instructions

1. **读取规范文件**：在分析需求或修改任何前端代码之前，先检查项目根目录（`/Users/yidaimiyaokangjilou/Downloads/测试用例显示工具/`）下是否存在 `FRONTEND_STYLE_GUIDE.md`。
2. **自动创建默认规范**：如果文件不存在，从本 Skill 的 `assets/FRONTEND_STYLE_GUIDE.template.md` 复制一份到项目根目录，命名为 `FRONTEND_STYLE_GUIDE.md`，然后读取它。
3. **结合规范分析需求**：读取规范后，根据其中的色彩、字体、间距、组件、动画等约定，向用户说明你的设计思路和视觉效果方案。
4. **等待用户确认**：在编写任何前端代码之前，先与用户讨论并确认方案。
5. **按规范编写代码**：用户确认后，再生成或修改前端代码，确保符合 `FRONTEND_STYLE_GUIDE.md` 中的约定。

## Notes

- 若用户明确要求忽略规范或临时突破规范，可在说明原因后执行，但建议提醒用户同步更新 `FRONTEND_STYLE_GUIDE.md`。
- 本 Skill 仅对当前项目生效，规范文件位于项目根目录，可与团队共享。
