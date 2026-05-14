<div align="center">

<h1>&#128269; Code Review Skill</h1>

<p>
  <strong>A comprehensive, modular code review skill for Claude Code</strong><br/>
  <strong>面向 Claude Code 的全面模块化代码审查技能</strong>
</p>

<p>
  <a href="https://github.com/awesome-skills/code-review-skill/blob/main/LICENSE">
    <img src="https://img.shields.io/badge/License-MIT-22c55e?style=flat-square" alt="License: MIT"/>
  </a>
  <img src="https://img.shields.io/badge/Claude_Code-Skill-7c3aed?style=flat-square&logo=anthropic&logoColor=white" alt="Claude Code Skill"/>
  <img src="https://img.shields.io/badge/Total_Lines-9%2C500%2B-3b82f6?style=flat-square" alt="9500+ lines"/>
  <img src="https://img.shields.io/badge/Languages-11%2B-f59e0b?style=flat-square" alt="11+ languages"/>
  <img src="https://img.shields.io/badge/PRs-Welcome-ec4899?style=flat-square" alt="PRs Welcome"/>
</p>

<p>
  <a href="#english">English</a>
  &middot;
  <a href="#chinese">中文</a>
  &middot;
  <a href="./CONTRIBUTING.md">Contributing</a>
</p>

</div>

---

<a name="english"></a>

## English

### What is this?

**Code Review Skill** is a production-ready skill for [Claude Code](https://claude.ai/code) that transforms AI-assisted code review from vague suggestions into a **structured, consistent, and expert-level** process.

It covers **11+ languages and frameworks** with over **9,500 lines** of carefully curated review guidelines — loaded progressively to minimize context window usage.

---

### &#10024; Key Features

- **Progressive Disclosure** — Core skill is ~190 lines; language guides (~200–1,000 lines each) load only when needed.
- **Four-Phase Review Process** — Structured workflow from understanding scope to delivering clear feedback.
- **Severity Labeling** — Every finding is categorized: `blocking` · `important` · `nit` · `suggestion` · `learning` · `praise`
- **Security-First** — Dedicated security checklists per language ecosystem.
- **Collaborative Tone** — Questions over commands, suggestions over mandates.
- **Automation Awareness** — Clearly separates what human review should catch vs. what linters handle.

---

### &#127760; Supported Languages & Frameworks

<table>
  <thead>
    <tr>
      <th>Category</th>
      <th>Technology</th>
      <th>Guide</th>
      <th>Lines</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td rowspan="4"><strong>Frontend</strong></td>
      <td>&#9883;&#65039; React 19 / Next.js / TanStack Query v5</td>
      <td><code>reference/react.md</code></td>
      <td>~870</td>
    </tr>
    <tr>
      <td>&#128154; Vue 3.5 + Composition API</td>
      <td><code>reference/vue.md</code></td>
      <td>~920</td>
    </tr>
    <tr>
      <td>&#127912; CSS / Less / Sass</td>
      <td><code>reference/css-less-sass.md</code></td>
      <td>~660</td>
    </tr>
    <tr>
      <td>&#128311; TypeScript</td>
      <td><code>reference/typescript.md</code></td>
      <td>~540</td>
    </tr>
    <tr>
      <td rowspan="4"><strong>Backend</strong></td>
      <td>&#9749; Java 17/21 + Spring Boot 3</td>
      <td><code>reference/java.md</code></td>
      <td>~800</td>
    </tr>
    <tr>
      <td>&#128013; Python</td>
      <td><code>reference/python.md</code></td>
      <td>~1,070</td>
    </tr>
    <tr>
      <td>&#128057; Go</td>
      <td><code>reference/go.md</code></td>
      <td>~990</td>
    </tr>
    <tr>
      <td>&#129408; Rust</td>
      <td><code>reference/rust.md</code></td>
      <td>~840</td>
    </tr>
    <tr>
      <td rowspan="3"><strong>Systems</strong></td>
      <td>&#9881;&#65039; C</td>
      <td><code>reference/c.md</code></td>
      <td>~210</td>
    </tr>
    <tr>
      <td>&#128297; C++</td>
      <td><code>reference/cpp.md</code></td>
      <td>~300</td>
    </tr>
    <tr>
      <td>&#128421;&#65039; Qt Framework</td>
      <td><code>reference/qt.md</code></td>
      <td>~190</td>
    </tr>
    <tr>
      <td rowspan="2"><strong>Architecture</strong></td>
      <td>&#127963;&#65039; Architecture Design Review</td>
      <td><code>reference/architecture-review-guide.md</code></td>
      <td>~470</td>
    </tr>
    <tr>
      <td>&#9889; Performance Review</td>
      <td><code>reference/performance-review-guide.md</code></td>
      <td>~750</td>
    </tr>
  </tbody>
</table>

---

### &#128260; The Four-Phase Review Process

```
Phase 1 - Context Gathering
  Understand PR scope, linked issues, and intent
                    |
                    v
Phase 2 - High-Level Review
  Architecture - Performance impact - Test strategy
                    |
                    v
Phase 3 - Line-by-Line Analysis
  Logic - Security - Maintainability - Edge cases
                    |
                    v
Phase 4 - Summary & Decision
  Structured feedback - Approval status - Action items
```

---

### &#127991;&#65039; Severity Labels

| Label | Meaning |
|-------|---------|
| &#128308; `blocking` | Must be fixed before merge |
| &#128992; `important` | Should be fixed; may block depending on context |
| &#128993; `nit` | Minor style or preference issue |
| &#128309; `suggestion` | Optional improvement worth considering |
| &#128218; `learning` | Educational note for the author |
| &#127775; `praise` | Explicitly highlight great work |

---

### &#128193; Repository Structure

```
code-review-skill/
|
+-- SKILL.md                              # Core skill - loaded on activation (~190 lines)
+-- README.md
+-- LICENSE
+-- CONTRIBUTING.md
|
+-- reference/                            # On-demand language guides
|   +-- react.md                          # React 19 / Next.js / TanStack Query v5
|   +-- vue.md                            # Vue 3.5 Composition API
|   +-- rust.md                           # Rust ownership, async/await, unsafe
|   +-- typescript.md                     # TypeScript strict mode, generics, ESLint
|   +-- java.md                           # Java 17/21 & Spring Boot 3
|   +-- python.md                         # Python async, typing, pytest
|   +-- go.md                             # Go goroutines, channels, context, interfaces
|   +-- c.md                              # C memory safety, UB, error handling
|   +-- cpp.md                            # C++ RAII, move semantics, exception safety
|   +-- qt.md                             # Qt object model, signals/slots, GUI perf
|   +-- css-less-sass.md                  # CSS/Less/Sass variables, responsive design
|   +-- architecture-review-guide.md      # SOLID, anti-patterns, coupling/cohesion
|   +-- performance-review-guide.md       # Core Web Vitals, N+1, memory leaks
|   +-- security-review-guide.md          # Security checklist (all languages)
|   +-- common-bugs-checklist.md          # Language-specific bug patterns
|   +-- code-review-best-practices.md     # Communication & process guidelines
|
+-- assets/
|   +-- review-checklist.md               # Quick reference checklist
|   +-- pr-review-template.md             # PR review comment template
|
+-- scripts/
    +-- pr-analyzer.py                    # PR complexity analyzer
```

---

### &#128640; Installation

**Clone to your Claude Code skills directory:**

```bash
# macOS / Linux
git clone https://github.com/awesome-skills/code-review-skill.git \
  ~/.claude/skills/code-review-skill

# Windows (PowerShell)
git clone https://github.com/awesome-skills/code-review-skill.git `
  "$env:USERPROFILE\.claude\skills\code-review-skill"
```

**Or add to an existing plugin:**

```bash
cp -r code-review-skill ~/.claude/plugins/your-plugin/skills/code-review/
```

---

### &#128161; Usage

Once installed, activate the skill in your Claude Code session:

```
Use code-review-skill to review this PR
```

Or create a custom slash command in `.claude/commands/`:

```markdown
<!-- .claude/commands/review.md -->
Use code-review-skill to perform a thorough review of the changes in this PR.
Focus on: security, performance, and maintainability.
```

**Example prompts:**

| Prompt | What happens |
|--------|-------------|
| `Review this React component` | Loads `react.md` - checks hooks, Server Components, Suspense patterns |
| `Review this Java PR` | Loads `java.md` - checks virtual threads, JPA, Spring Boot 3 patterns |
| `Security review of this Go service` | Loads `go.md` + `security-review-guide.md` |
| `Architecture review` | Loads `architecture-review-guide.md` - SOLID, anti-patterns, coupling |
| `Performance review` | Loads `performance-review-guide.md` - Web Vitals, N+1, complexity |

---

### &#128300; Highlights by Language

<details>
<summary><strong>&#9883;&#65039; React 19</strong></summary>

- `useActionState` - Unified form state management
- `useFormStatus` - Access parent form status without prop drilling
- `useOptimistic` - Optimistic UI updates with automatic rollback
- Server Components & Server Actions patterns (Next.js 15+)
- Suspense boundary design, Error Boundary integration, streaming SSR
- `use()` Hook for consuming Promises

</details>

<details>
<summary><strong>&#9749; Java & Spring Boot 3</strong></summary>

- **Java 17/21**: Records, Pattern Matching for Switch, Text Blocks, Sealed Classes
- **Virtual Threads** (Project Loom): High-throughput I/O patterns
- **Spring Boot 3**: Constructor injection, `@ConfigurationProperties`, `ProblemDetail`
- **JPA Performance**: Solving N+1, correct `equals`/`hashCode` on Entities

</details>

<details>
<summary><strong>&#129408; Rust</strong></summary>

- Ownership patterns and common pitfalls
- `unsafe` code review requirements (mandatory `SAFETY` comments)
- Async/await - avoiding blocking in async context, cancellation safety
- Error handling: `thiserror` for libraries, `anyhow` for applications

</details>

<details>
<summary><strong>&#128057; Go</strong></summary>

- Goroutine lifecycle management and leak prevention
- Channel patterns, select usage
- `context.Context` propagation
- Interface design (accept interfaces, return structs)
- Error wrapping with `%w`

</details>

<details>
<summary><strong>&#9881;&#65039; C / C++</strong></summary>

- **C**: Pointer/buffer safety, undefined behavior, resource cleanup, integer overflow
- **C++**: RAII ownership, Rule of 0/3/5, move semantics, exception safety, `noexcept`
- **Qt**: Object parent/child memory model, thread-safe signal/slot connections, GUI performance

</details>

---

### &#129309; Contributing

Contributions are welcome! See [CONTRIBUTING.md](./CONTRIBUTING.md) for guidelines.

**Ideas:**
- New language guides (C#, Swift, Kotlin, Ruby, PHP...)
- Framework-specific guides (Django, Laravel, NestJS...)
- Additional checklists and templates
- Translations of core documentation

---

### &#128196; License

MIT &copy; [awesome-skills](https://github.com/awesome-skills)

---

<a name="chinese"></a>

## 中文

### 这是什么？

**Code Review Skill** 是专为 [Claude Code](https://claude.ai/code) 打造的生产级代码审查技能，将 AI 辅助的代码审查从模糊建议转变为**结构化、一致且专业级**的流程。

覆盖 **11+ 种语言和框架**，拥有超过 **9,500 行**精心整理的代码审查指南——按需加载，最大程度减少上下文占用。

---

### &#10024; 核心特性

- **渐进式加载** — 核心技能仅 ~190 行，各语言指南（每份 200–1,000 行）仅在需要时才加载。
- **四阶段审查流程** — 从理解 PR 范围到输出清晰反馈，每一步都有规可循。
- **严重性标记** — 每条发现均分级：`blocking` · `important` · `nit` · `suggestion` · `learning` · `praise`
- **安全优先** — 每个语言生态均配备专属安全检查清单。
- **协作式语气** — 以提问替代命令，以建议替代指令。
- **自动化感知** — 明确区分人工审查应关注的内容与 linter 自动处理的内容。

---

### &#127760; 支持的语言与框架

| 分类 | 技术栈 | 指南文件 | 行数 |
|------|--------|----------|------|
| **前端** | &#9883;&#65039; React 19 / Next.js / TanStack Query v5 | `reference/react.md` | ~870 |
| | &#128154; Vue 3.5 Composition API | `reference/vue.md` | ~920 |
| | &#127912; CSS / Less / Sass | `reference/css-less-sass.md` | ~660 |
| | &#128311; TypeScript | `reference/typescript.md` | ~540 |
| **后端** | &#9749; Java 17/21 + Spring Boot 3 | `reference/java.md` | ~800 |
| | &#128013; Python | `reference/python.md` | ~1,070 |
| | &#128057; Go | `reference/go.md` | ~990 |
| | &#129408; Rust | `reference/rust.md` | ~840 |
| **系统级** | &#9881;&#65039; C | `reference/c.md` | ~210 |
| | &#128297; C++ | `reference/cpp.md` | ~300 |
| | &#128421;&#65039; Qt 框架 | `reference/qt.md` | ~190 |
| **架构** | &#127963;&#65039; 架构设计审查 | `reference/architecture-review-guide.md` | ~470 |
| | &#9889; 性能审查 | `reference/performance-review-guide.md` | ~750 |

---

### &#128260; 四阶段审查流程

```
阶段一 - 上下文收集
  理解 PR 范围、关联 Issue 和实现意图
                    |
                    v
阶段二 - 高层级审查
  架构设计 - 性能影响 - 测试策略
                    |
                    v
阶段三 - 逐行深度分析
  逻辑正确性 - 安全漏洞 - 可维护性 - 边界情况
                    |
                    v
阶段四 - 总结与决策
  结构化反馈 - 审批状态 - 后续行动项
```

---

### &#127991;&#65039; 严重性标记说明

| 标记 | 含义 |
|------|------|
| &#128308; `blocking` | 合并前必须修复 |
| &#128992; `important` | 应当修复，视情况可能阻塞合并 |
| &#128993; `nit` | 风格或偏好上的小问题 |
| &#128309; `suggestion` | 值得考虑的可选优化 |
| &#128218; `learning` | 给作者的教育性说明 |
| &#127775; `praise` | 明确表扬优秀代码 |

---

### &#128193; 仓库结构

```
code-review-skill/
|
+-- SKILL.md                              # 核心技能，激活时加载（~190 行）
+-- README.md
+-- LICENSE
+-- CONTRIBUTING.md
|
+-- reference/                            # 按需加载的语言指南
|   +-- react.md                          # React 19 / Next.js / TanStack Query v5
|   +-- vue.md                            # Vue 3.5 组合式 API
|   +-- rust.md                           # Rust 所有权、async/await、unsafe
|   +-- typescript.md                     # TypeScript strict 模式、泛型、ESLint
|   +-- java.md                           # Java 17/21 & Spring Boot 3
|   +-- python.md                         # Python async、类型注解、pytest
|   +-- go.md                             # Go goroutine、channel、context、接口
|   +-- c.md                              # C 内存安全、UB、错误处理
|   +-- cpp.md                            # C++ RAII、移动语义、异常安全
|   +-- qt.md                             # Qt 对象模型、信号/槽、GUI 性能
|   +-- css-less-sass.md                  # CSS/Less/Sass 变量、响应式设计
|   +-- architecture-review-guide.md      # SOLID、反模式、耦合度分析
|   +-- performance-review-guide.md       # Core Web Vitals、N+1、内存泄漏
|   +-- security-review-guide.md          # 安全审查清单（全语言通用）
|   +-- common-bugs-checklist.md          # 各语言常见 Bug 模式
|   +-- code-review-best-practices.md     # 沟通与流程最佳实践
|
+-- assets/
|   +-- review-checklist.md               # 快速参考清单
|   +-- pr-review-template.md             # PR 审查评论模板
|
+-- scripts/
    +-- pr-analyzer.py                    # PR 复杂度分析工具
```

---

### &#128640; 安装方法

**克隆到 Claude Code skills 目录：**

```bash
# macOS / Linux
git clone https://github.com/awesome-skills/code-review-skill.git \
  ~/.claude/skills/code-review-skill

# Windows（PowerShell）
git clone https://github.com/awesome-skills/code-review-skill.git `
  "$env:USERPROFILE\.claude\skills\code-review-skill"
```

**或添加到现有插件：**

```bash
cp -r code-review-skill ~/.claude/plugins/your-plugin/skills/code-review/
```

---

### &#128161; 使用方式

安装后，在 Claude Code 会话中激活技能：

```
Use code-review-skill to review this PR
```

或在 `.claude/commands/` 中创建自定义斜杠命令：

```markdown
<!-- .claude/commands/review.md -->
使用 code-review-skill 对这次 PR 的变更进行全面审查。
重点关注：安全性、性能和可维护性。
```

**示例提示词：**

| 提示词 | 效果 |
|--------|------|
| `审查这个 React 组件` | 加载 `react.md`，检查 Hooks、Server Components、Suspense |
| `审查这个 Java PR` | 加载 `java.md`，检查虚拟线程、JPA、Spring Boot 3 |
| `对这个 Go 服务进行安全审查` | 加载 `go.md` + `security-review-guide.md` |
| `架构审查` | 加载 `architecture-review-guide.md`，检查 SOLID 与反模式 |
| `性能审查` | 加载 `performance-review-guide.md`，分析 Web Vitals、N+1 等 |

---

### &#128300; 各语言核心内容

<details>
<summary><strong>&#9883;&#65039; React 19</strong></summary>

- `useActionState` — 统一的表单状态管理
- `useFormStatus` — 无需 props 透传即可访问父表单状态
- `useOptimistic` — 带自动回滚的乐观 UI 更新
- Server Components & Server Actions（Next.js 15+）
- Suspense 边界设计、Error Boundary 集成、流式 SSR
- `use()` Hook 消费 Promise

</details>

<details>
<summary><strong>&#9749; Java & Spring Boot 3</strong></summary>

- **Java 17/21**：Records、Switch 模式匹配、文本块、Sealed Classes
- **虚拟线程**（Project Loom）：高吞吐量 I/O 模式
- **Spring Boot 3**：构造器注入、`@ConfigurationProperties`、`ProblemDetail`
- **JPA 性能**：解决 N+1、Entity 正确的 `equals`/`hashCode` 实现

</details>

<details>
<summary><strong>&#129408; Rust</strong></summary>

- 所有权模式与常见陷阱
- `unsafe` 代码审查要求（必须有 `SAFETY` 注释）
- Async/await — 避免在异步上下文中阻塞，取消安全性
- 错误处理：库用 `thiserror`，应用用 `anyhow`

</details>

<details>
<summary><strong>&#128057; Go</strong></summary>

- Goroutine 生命周期管理与泄漏预防
- Channel 模式、select 用法
- `context.Context` 传播规范
- 接口设计原则（接受接口，返回结构体）
- 错误包装：使用 `%w`

</details>

<details>
<summary><strong>&#9881;&#65039; C / C++</strong></summary>

- **C**：指针/缓冲区安全、未定义行为、资源清理、整数溢出
- **C++**：RAII 所有权、Rule of 0/3/5、移动语义、异常安全、`noexcept`
- **Qt**：父子内存模型、线程安全的信号/槽连接、GUI 性能优化

</details>

---

### &#129309; 参与贡献

欢迎贡献！请查阅 [CONTRIBUTING.md](./CONTRIBUTING.md) 了解规范。

**可贡献方向：**
- 新增语言指南（C#、Swift、Kotlin、Ruby、PHP...）
- 框架专属指南（Django、Laravel、NestJS...）
- 补充检查清单和审查模板
- 核心文档的多语言翻译

---

### &#128196; 开源协议

MIT &copy; [awesome-skills](https://github.com/awesome-skills)

---

<div align="center">
  Made with &#10084;&#65039; for developers who care about code quality
</div>
