# Contributing to AI Code Review Guide

Thank you for your interest in contributing! This document provides guidelines for contributing to this Claude Code Skill project.

## Claude Code Skill 开发规范

本项目是一个 Claude Code Skill，贡献者需要遵循以下规范。

### 目录结构

```
ai-code-review-guide/
├── SKILL.md                    # 必需：主文件（始终加载）
├── README.md                   # 项目说明文档
├── CONTRIBUTING.md             # 贡献指南（本文件）
├── LICENSE                     # 许可证
├── reference/                  # 按需加载的详细指南
│   ├── react.md
│   ├── vue.md
│   ├── rust.md
│   ├── typescript.md
│   ├── python.md
│   ├── c.md
│   ├── cpp.md
│   ├── common-bugs-checklist.md
│   ├── security-review-guide.md
│   └── code-review-best-practices.md
├── assets/                     # 模板和快速参考
│   ├── review-checklist.md
│   └── pr-review-template.md
└── scripts/                    # 工具脚本
    └── pr-analyzer.py
```

### Frontmatter 规范

SKILL.md 必须包含 YAML frontmatter：

```yaml
---
name: skill-name
description: |
  功能描述。触发条件说明。
  Use when [具体使用场景]。
allowed-tools: ["Read", "Grep", "Glob"]  # 可选：限制工具访问
---
```

#### 必需字段

| 字段 | 说明 | 约束 |
|------|------|------|
| `name` | Skill 标识符 | 小写字母、数字、连字符；最多 64 字符 |
| `description` | 功能和激活条件 | 最多 1024 字符；必须包含 "Use when" |

#### 可选字段

| 字段 | 说明 | 示例 |
|------|------|------|
| `allowed-tools` | 限制工具访问 | `["Read", "Grep", "Glob"]` |

### 命名约定

**Skill 名称规则**：
- 仅使用小写字母、数字和连字符（kebab-case）
- 最多 64 个字符
- 避免下划线或大写字母

```
✅ 正确：code-review-excellence, typescript-advanced-types
❌ 错误：CodeReview, code_review, TYPESCRIPT
```

**文件命名规则**：
- reference 文件使用小写：`react.md`, `vue.md`
- 多词文件使用连字符：`common-bugs-checklist.md`

### Description 写法规范

Description 必须包含两部分：

1. **功能陈述**：具体说明 Skill 能做什么
2. **触发条件**：以 "Use when" 开头，说明何时激活

```yaml
# ✅ 正确示例
description: |
  Provides comprehensive code review guidance for React 19, Vue 3, Rust,
  TypeScript, Java, Python, and C/C++.
  Helps catch bugs, improve code quality, and give constructive feedback.
  Use when reviewing pull requests, conducting PR reviews, establishing
  review standards, or mentoring developers through code reviews.

# ❌ 错误示例（太模糊，缺少触发条件）
description: |
  Helps with code review.
```

### Progressive Disclosure（渐进式披露）

Claude 只在需要时加载支持文件，不会一次性加载所有内容。

#### 文件职责划分

| 文件 | 加载时机 | 内容 |
|------|----------|------|
| `SKILL.md` | 始终加载 | 核心原则、快速索引、何时使用 |
| `reference/*.md` | 按需加载 | 语言/框架的详细指南 |
| `assets/*.md` | 明确需要时 | 模板、清单 |
| `scripts/*.py` | 明确指引时 | 工具脚本 |

#### 内容组织原则

**SKILL.md**（~200 行以内）：
- 简述：2-3 句话说明用途
- 核心原则和方法论
- 语言/框架索引表（链接到 reference/）
- 何时使用此 Skill

**reference/*.md**（详细内容）：
- 完整的代码示例
- 所有最佳实践
- Review Checklist
- 边界情况和陷阱

### 文件引用规范

在 SKILL.md 中引用其他文件时：

```markdown
# ✅ 正确：使用 Markdown 链接格式
| **React** | [React Guide](reference/react.md) | Hooks, React 19, RSC |
| **Vue 3** | [Vue Guide](reference/vue.md) | Composition API |

详见 [React Guide](reference/react.md) 获取完整指南。

# ❌ 错误：使用代码块格式
参考 `reference/react.md` 文件。
```

**路径规则**：
- 使用相对路径（相对于 Skill 目录）
- 使用正斜杠 `/`，不使用反斜杠
- 不需要 `./` 前缀

---

## 贡献类型

### 添加新语言支持

1. 在 `reference/` 目录创建新文件（如 `go.md`）
2. 遵循以下结构：

```markdown
# [Language] Code Review Guide

> 简短描述，一句话说明覆盖内容。

## 目录
- [主题1](#主题1)
- [主题2](#主题2)
- [Review Checklist](#review-checklist)

---

## 主题1

### 子主题

```[language]
// ❌ Bad pattern - 说明为什么不好
bad_code_example()

// ✅ Good pattern - 说明为什么好
good_code_example()
```

---

## Review Checklist

### 类别1
- [ ] 检查项 1
- [ ] 检查项 2
```

3. 在 `SKILL.md` 的索引表中添加链接
4. 更新 `README.md` 的统计信息

### 添加框架模式

1. 确保引用官方文档
2. 包含版本号（如 "React 19", "Vue 3.5+"）
3. 提供可运行的代码示例
4. 添加对应的 checklist 项

### 改进现有内容

- 修复拼写或语法错误
- 更新过时的模式（注明版本变化）
- 添加边界情况示例
- 改进代码示例的清晰度

---

## 代码示例规范

### 格式要求

```markdown
// ❌ 问题描述 - 解释为什么这样做不好
problematic_code()

// ✅ 推荐做法 - 解释为什么这样做更好
recommended_code()
```

### 质量标准

- 示例应基于真实场景，避免人为构造
- 同时展示问题和解决方案
- 保持示例简洁聚焦
- 包含必要的上下文（import 语句等）

---

## 提交流程

### Issue 报告

- 使用 GitHub Issues 报告问题或建议
- 提供清晰的描述和示例
- 标注相关的语言/框架

### Pull Request 流程

1. Fork 仓库
2. 创建功能分支：`git checkout -b feature/add-go-support`
3. 进行修改
4. 提交（见下文 commit 格式）
5. 推送到 fork：`git push origin feature/add-go-support`
6. 创建 Pull Request

### Commit 消息格式

```
类型: 简短描述

详细说明（如需要）

- 具体变更 1
- 具体变更 2
```

**类型**：
- `feat`: 新功能或新内容
- `fix`: 修复错误
- `docs`: 仅文档变更
- `refactor`: 重构（不改变功能）
- `chore`: 维护性工作

**示例**：
```
feat: 添加 Go 语言代码审查指南

- 新增 reference/go.md
- 覆盖错误处理、并发、接口设计
- 更新 SKILL.md 索引表
```

---

## Skill 设计原则

### 单一职责

每个 Skill 专注一个核心能力。本 Skill 专注于**代码审查**，不应扩展到：
- 代码生成
- 项目初始化
- 部署配置

### 版本管理

- 在 reference 文件中标注框架/语言版本
- 更新时在 commit 中说明版本变化
- 过时内容应更新而非删除（除非完全废弃）

### 内容质量

- 所有建议应有依据（官方文档、最佳实践）
- 避免主观偏好（如代码风格），专注于客观问题
- 优先覆盖常见陷阱和安全问题

---

## 常见问题

### Q: 如何测试我的更改？

将修改后的 Skill 复制到 `~/.claude/skills/` 目录，然后在 Claude Code 中测试：
```bash
cp -r ai-code-review-guide ~/.claude/skills/code-review-excellence
```

### Q: 我应该更新 SKILL.md 还是 reference 文件？

- **SKILL.md**：只修改索引表或核心原则
- **reference/*.md**：添加/更新具体的语言或框架内容

### Q: 如何处理过时的内容？

1. 标注版本变化（如 "React 18 → React 19"）
2. 保留旧版本内容（如果仍有用户使用）
3. 在 checklist 中更新相关项

---

## 问题咨询

如有任何问题，欢迎在 GitHub Issues 中提问。
