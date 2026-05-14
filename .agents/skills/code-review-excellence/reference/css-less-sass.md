# CSS / Less / Sass Review Guide

CSS 及预处理器代码审查指南，覆盖性能、可维护性、响应式设计和浏览器兼容性。

## CSS 变量 vs 硬编码

### 应该使用变量的场景

```css
/* ❌ 硬编码 - 难以维护 */
.button {
  background: #3b82f6;
  border-radius: 8px;
}
.card {
  border: 1px solid #3b82f6;
  border-radius: 8px;
}

/* ✅ 使用 CSS 变量 */
:root {
  --color-primary: #3b82f6;
  --radius-md: 8px;
}
.button {
  background: var(--color-primary);
  border-radius: var(--radius-md);
}
.card {
  border: 1px solid var(--color-primary);
  border-radius: var(--radius-md);
}
```

### 变量命名规范

```css
/* 推荐的变量分类 */
:root {
  /* 颜色 */
  --color-primary: #3b82f6;
  --color-primary-hover: #2563eb;
  --color-text: #1f2937;
  --color-text-muted: #6b7280;
  --color-bg: #ffffff;
  --color-border: #e5e7eb;

  /* 间距 */
  --spacing-xs: 4px;
  --spacing-sm: 8px;
  --spacing-md: 16px;
  --spacing-lg: 24px;
  --spacing-xl: 32px;

  /* 字体 */
  --font-size-sm: 14px;
  --font-size-base: 16px;
  --font-size-lg: 18px;
  --font-weight-normal: 400;
  --font-weight-bold: 700;

  /* 圆角 */
  --radius-sm: 4px;
  --radius-md: 8px;
  --radius-lg: 12px;
  --radius-full: 9999px;

  /* 阴影 */
  --shadow-sm: 0 1px 2px rgba(0, 0, 0, 0.05);
  --shadow-md: 0 4px 6px rgba(0, 0, 0, 0.1);

  /* 过渡 */
  --transition-fast: 150ms ease;
  --transition-normal: 300ms ease;
}
```

### 变量作用域建议

```css
/* ✅ 组件级变量 - 减少全局污染 */
.card {
  --card-padding: var(--spacing-md);
  --card-radius: var(--radius-md);

  padding: var(--card-padding);
  border-radius: var(--card-radius);
}

/* ⚠️ 避免频繁用 JS 动态修改变量 - 影响性能 */
```

### 审查清单

- [ ] 颜色值是否使用变量？
- [ ] 间距是否来自设计系统？
- [ ] 重复值是否提取为变量？
- [ ] 变量命名是否语义化？

---

## !important 使用规范

### 何时可以使用

```css
/* ✅ 工具类 - 明确需要覆盖 */
.hidden { display: none !important; }
.sr-only { position: absolute !important; }

/* ✅ 覆盖第三方库样式（无法修改源码时） */
.third-party-modal {
  z-index: 9999 !important;
}

/* ✅ 打印样式 */
@media print {
  .no-print { display: none !important; }
}
```

### 何时禁止使用

```css
/* ❌ 解决特异性问题 - 应该重构选择器 */
.button {
  background: blue !important;  /* 为什么需要 !important? */
}

/* ❌ 覆盖自己写的样式 */
.card { padding: 20px; }
.card { padding: 30px !important; }  /* 直接修改原规则 */

/* ❌ 在组件样式中 */
.my-component .title {
  font-size: 24px !important;  /* 破坏组件封装 */
}
```

### 替代方案

```css
/* 问题：需要覆盖 .btn 的样式 */

/* ❌ 使用 !important */
.my-btn {
  background: red !important;
}

/* ✅ 提高特异性 */
button.my-btn {
  background: red;
}

/* ✅ 使用更具体的选择器 */
.container .my-btn {
  background: red;
}

/* ✅ 使用 :where() 降低被覆盖样式的特异性 */
:where(.btn) {
  background: blue;  /* 特异性为 0 */
}
.my-btn {
  background: red;   /* 可以正常覆盖 */
}
```

### 审查问题

```markdown
🔴 [blocking] "发现 15 处 !important，请说明每处的必要性"
🟡 [important] "这个 !important 可以通过调整选择器特异性来解决"
💡 [suggestion] "考虑使用 CSS Layers (@layer) 来管理样式优先级"
```

---

## 性能考虑

### 🔴 高危性能问题

#### 1. `transition: all` 问题

```css
/* ❌ 性能杀手 - 浏览器检查所有可动画属性 */
.button {
  transition: all 0.3s ease;
}

/* ✅ 明确指定属性 */
.button {
  transition: background-color 0.3s ease, transform 0.3s ease;
}

/* ✅ 多属性时使用变量 */
.button {
  --transition-duration: 0.3s;
  transition:
    background-color var(--transition-duration) ease,
    box-shadow var(--transition-duration) ease,
    transform var(--transition-duration) ease;
}
```

#### 2. box-shadow 动画

```css
/* ❌ 每帧触发重绘 - 严重影响性能 */
.card {
  box-shadow: 0 2px 4px rgba(0,0,0,0.1);
  transition: box-shadow 0.3s ease;
}
.card:hover {
  box-shadow: 0 8px 16px rgba(0,0,0,0.2);
}

/* ✅ 使用伪元素 + opacity */
.card {
  position: relative;
}
.card::after {
  content: '';
  position: absolute;
  inset: 0;
  box-shadow: 0 8px 16px rgba(0,0,0,0.2);
  opacity: 0;
  transition: opacity 0.3s ease;
  pointer-events: none;
  border-radius: inherit;
}
.card:hover::after {
  opacity: 1;
}
```

#### 3. 触发布局（Reflow）的属性

```css
/* ❌ 动画这些属性会触发布局重计算 */
.bad-animation {
  transition: width 0.3s, height 0.3s, top 0.3s, left 0.3s, margin 0.3s;
}

/* ✅ 只动画 transform 和 opacity（仅触发合成） */
.good-animation {
  transition: transform 0.3s, opacity 0.3s;
}

/* 位移用 translate 代替 top/left */
.move {
  transform: translateX(100px);  /* ✅ */
  /* left: 100px; */             /* ❌ */
}

/* 缩放用 scale 代替 width/height */
.grow {
  transform: scale(1.1);  /* ✅ */
  /* width: 110%; */      /* ❌ */
}
```

### 🟡 中等性能问题

#### 复杂选择器

```css
/* ❌ 过深的嵌套 - 选择器匹配慢 */
.page .container .content .article .section .paragraph span {
  color: red;
}

/* ✅ 扁平化 */
.article-text {
  color: red;
}

/* ❌ 通配符选择器 */
* { box-sizing: border-box; }           /* 影响所有元素 */
[class*="icon-"] { display: inline; }   /* 属性选择器较慢 */

/* ✅ 限制范围 */
.icon-box * { box-sizing: border-box; }
```

#### 大量阴影和滤镜

```css
/* ⚠️ 复杂阴影影响渲染性能 */
.heavy-shadow {
  box-shadow:
    0 1px 2px rgba(0,0,0,0.1),
    0 2px 4px rgba(0,0,0,0.1),
    0 4px 8px rgba(0,0,0,0.1),
    0 8px 16px rgba(0,0,0,0.1),
    0 16px 32px rgba(0,0,0,0.1);  /* 5 层阴影 */
}

/* ⚠️ 滤镜消耗 GPU */
.blur-heavy {
  filter: blur(20px) brightness(1.2) contrast(1.1);
  backdrop-filter: blur(10px);  /* 更消耗性能 */
}
```

### 性能优化建议

```css
/* 使用 will-change 提示浏览器（谨慎使用） */
.animated-element {
  will-change: transform, opacity;
}

/* 动画完成后移除 will-change */
.animated-element.idle {
  will-change: auto;
}

/* 使用 contain 限制重绘范围 */
.card {
  contain: layout paint;  /* 告诉浏览器内部变化不影响外部 */
}
```

### 审查清单

- [ ] 是否使用 `transition: all`？
- [ ] 是否动画 width/height/top/left？
- [ ] box-shadow 是否被动画？
- [ ] 选择器嵌套是否超过 3 层？
- [ ] 是否有不必要的 `will-change`？

---

## 响应式设计检查点

### Mobile First 原则

```css
/* ✅ Mobile First - 基础样式针对移动端 */
.container {
  padding: 16px;
  display: flex;
  flex-direction: column;
}

/* 逐步增强 */
@media (min-width: 768px) {
  .container {
    padding: 24px;
    flex-direction: row;
  }
}

@media (min-width: 1024px) {
  .container {
    padding: 32px;
    max-width: 1200px;
    margin: 0 auto;
  }
}

/* ❌ Desktop First - 需要覆盖更多样式 */
.container {
  max-width: 1200px;
  padding: 32px;
  flex-direction: row;
}

@media (max-width: 1023px) {
  .container {
    padding: 24px;
  }
}

@media (max-width: 767px) {
  .container {
    padding: 16px;
    flex-direction: column;
    max-width: none;
  }
}
```

### 断点建议

```css
/* 推荐断点（基于内容而非设备） */
:root {
  --breakpoint-sm: 640px;   /* 大手机 */
  --breakpoint-md: 768px;   /* 平板竖屏 */
  --breakpoint-lg: 1024px;  /* 平板横屏/小笔记本 */
  --breakpoint-xl: 1280px;  /* 桌面 */
  --breakpoint-2xl: 1536px; /* 大桌面 */
}

/* 使用示例 */
@media (min-width: 768px) { /* md */ }
@media (min-width: 1024px) { /* lg */ }
```

### 响应式审查清单

- [ ] 是否采用 Mobile First？
- [ ] 断点是否基于内容断裂点而非设备？
- [ ] 是否避免断点重叠？
- [ ] 文字是否使用相对单位（rem/em）？
- [ ] 触摸目标是否足够大（≥44px）？
- [ ] 是否测试了横竖屏切换？

### 常见问题

```css
/* ❌ 固定宽度 */
.container {
  width: 1200px;
}

/* ✅ 最大宽度 + 弹性 */
.container {
  width: 100%;
  max-width: 1200px;
  padding-inline: 16px;
}

/* ❌ 固定高度的文本容器 */
.text-box {
  height: 100px;  /* 文字可能溢出 */
}

/* ✅ 最小高度 */
.text-box {
  min-height: 100px;
}

/* ❌ 小触摸目标 */
.small-button {
  padding: 4px 8px;  /* 太小，难以点击 */
}

/* ✅ 足够的触摸区域 */
.touch-button {
  min-height: 44px;
  min-width: 44px;
  padding: 12px 16px;
}
```

---

## 浏览器兼容性

### 需要检查的特性

| 特性 | 兼容性 | 建议 |
|------|--------|------|
| CSS Grid | 现代浏览器 ✅ | IE 需要 Autoprefixer + 测试 |
| Flexbox | 广泛支持 ✅ | 旧版需要前缀 |
| CSS Variables | 现代浏览器 ✅ | IE 不支持，需要回退 |
| `gap` (flexbox) | 较新 ⚠️ | Safari 14.1+ |
| `:has()` | 较新 ⚠️ | Firefox 121+ |
| `container queries` | 较新 ⚠️ | 2023 年后的浏览器 |
| `@layer` | 较新 ⚠️ | 检查目标浏览器 |

### 回退策略

```css
/* CSS 变量回退 */
.button {
  background: #3b82f6;              /* 回退值 */
  background: var(--color-primary); /* 现代浏览器 */
}

/* Flexbox gap 回退 */
.flex-container {
  display: flex;
  gap: 16px;
}
/* 旧浏览器回退 */
.flex-container > * + * {
  margin-left: 16px;
}

/* Grid 回退 */
.grid {
  display: flex;
  flex-wrap: wrap;
}
@supports (display: grid) {
  .grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
  }
}
```

### Autoprefixer 配置

```javascript
// postcss.config.js
module.exports = {
  plugins: [
    require('autoprefixer')({
      // 根据 browserslist 配置
      grid: 'autoplace',  // 启用 Grid 前缀（IE 支持）
      flexbox: 'no-2009', // 只用现代 flexbox 语法
    }),
  ],
};

// package.json
{
  "browserslist": [
    "> 1%",
    "last 2 versions",
    "not dead",
    "not ie 11"  // 根据项目需求
  ]
}
```

### 审查清单

- [ ] 是否检查了 [Can I Use](https://caniuse.com)？
- [ ] 新特性是否有回退方案？
- [ ] 是否配置了 Autoprefixer？
- [ ] browserslist 是否符合项目要求？
- [ ] 是否在目标浏览器中测试？

---

## Less / Sass 特定问题

### 嵌套深度

```scss
/* ❌ 过深嵌套 - 编译后选择器过长 */
.page {
  .container {
    .content {
      .article {
        .title {
          color: red;  // 编译为 .page .container .content .article .title
        }
      }
    }
  }
}

/* ✅ 最多 3 层 */
.article {
  &__title {
    color: red;
  }

  &__content {
    p { margin-bottom: 1em; }
  }
}
```

### Mixin vs Extend vs 变量

```scss
/* 变量 - 用于单个值 */
$primary-color: #3b82f6;

/* Mixin - 用于可配置的代码块 */
@mixin button-variant($bg, $text) {
  background: $bg;
  color: $text;
  &:hover {
    background: darken($bg, 10%);
  }
}

/* Extend - 用于共享相同样式（谨慎使用） */
%visually-hidden {
  position: absolute;
  width: 1px;
  height: 1px;
  overflow: hidden;
  clip: rect(0, 0, 0, 0);
}

.sr-only {
  @extend %visually-hidden;
}

/* ⚠️ @extend 的问题 */
// 可能产生意外的选择器组合
// 不能在 @media 中使用
// 优先使用 mixin
```

### 审查清单

- [ ] 嵌套是否超过 3 层？
- [ ] 是否滥用 @extend？
- [ ] Mixin 是否过于复杂？
- [ ] 编译后的 CSS 大小是否合理？

---

## 快速审查清单

### 🔴 必须修复

```markdown
□ transition: all
□ 动画 width/height/top/left/margin
□ 大量 !important
□ 硬编码的颜色/间距重复 >3 次
□ 选择器嵌套 >4 层
```

### 🟡 建议修复

```markdown
□ 缺少响应式处理
□ 使用 Desktop First
□ 复杂 box-shadow 被动画
□ 缺少浏览器兼容回退
□ CSS 变量作用域过大
```

### 🟢 优化建议

```markdown
□ 可以使用 CSS Grid 简化布局
□ 可以使用 CSS 变量提取重复值
□ 可以使用 @layer 管理优先级
□ 可以添加 contain 优化性能
```

---

## 工具推荐

| 工具 | 用途 |
|------|------|
| [Stylelint](https://stylelint.io/) | CSS 代码检查 |
| [PurgeCSS](https://purgecss.com/) | 移除未使用 CSS |
| [Autoprefixer](https://autoprefixer.github.io/) | 自动添加前缀 |
| [CSS Stats](https://cssstats.com/) | 分析 CSS 统计 |
| [Can I Use](https://caniuse.com/) | 浏览器兼容性查询 |

---

## 参考资源

- [CSS Performance Optimization - MDN](https://developer.mozilla.org/en-US/docs/Learn_web_development/Extensions/Performance/CSS)
- [What a CSS Code Review Might Look Like - CSS-Tricks](https://css-tricks.com/what-a-css-code-review-might-look-like/)
- [How to Animate Box-Shadow - Tobias Ahlin](https://tobiasahlin.com/blog/how-to-animate-box-shadow/)
- [Media Query Fundamentals - MDN](https://developer.mozilla.org/en-US/docs/Learn_web_development/Core/CSS_layout/Media_queries)
- [Autoprefixer - GitHub](https://github.com/postcss/autoprefixer)
