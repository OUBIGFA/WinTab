# Antigravity Kit 架构

> **版本 5.0** - 全面的 AI 代理能力扩展工具包

---

## 📋 概览

Antigravity Kit 是一个模块化系统，由以下部分组成：
- **16 个专家代理** - 基于角色的 AI 角色
- **40 项技能** - 面向特定领域的知识模块
- **11 个工作流** - 斜杠命令流程

---

## 🏗️ 目录结构

```
.agent/
├── ARCHITECTURE.md          # 本文件
├── agents/                  # 16 个专家代理
├── skills/                  # 40 项技能
├── workflows/               # 11 个斜杠命令
├── rules/                   # 全局规则
└── .shared/                 # 共享资源
```

---

## 🤖 代理（16）

面向不同领域的专家 AI 角色。

| 代理 | 方向 | 使用技能 |
|-------|-------|-------------|
| `orchestrator` | 多代理协同 | parallel-agents, behavioral-modes |
| `project-planner` | 发现与任务规划 | brainstorming, plan-writing, architecture |
| `frontend-specialist` | Web UI/UX | frontend-design, react-patterns, tailwind-patterns |
| `backend-specialist` | API 与业务逻辑 | api-patterns, nodejs-best-practices, database-design |
| `database-architect` | 模式与 SQL | database-design, prisma-expert |
| `mobile-developer` | iOS, Android, RN | mobile-design |
| `game-developer` | 游戏逻辑与机制 | game-development |
| `devops-engineer` | CI/CD, Docker | deployment-procedures, docker-expert |
| `security-auditor` | 安全合规 | vulnerability-scanner, red-team-tactics |
| `penetration-tester` | 进攻性安全 | red-team-tactics |
| `test-engineer` | 测试策略 | testing-patterns, tdd-workflow, webapp-testing |
| `debugger` | 根因分析 | systematic-debugging |
| `performance-optimizer` | 性能速度、Web Vitals | performance-profiling |
| `seo-specialist` | 排名与可见性 | seo-fundamentals, geo-fundamentals |
| `documentation-writer` | 手册与文档 | documentation-templates |
| `explorer-agent` | 代码库分析 | - |

---

## 🧠 技能（40）

面向特定领域的知识模块。技能会根据任务上下文按需加载。

### 前端与 UI
| 技能 | 描述 |
|-------|-------------|
| `react-patterns` | React Hooks、状态与性能 |
| `nextjs-best-practices` | App Router、Server Components |
| `tailwind-patterns` | Tailwind CSS v4 工具类 |
| `frontend-design` | UI/UX 模式与设计系统 |
| `ui-ux-pro-max` | 50 种风格、21 套配色、50 款字体 |

### 后端与 API
| 技能 | 描述 |
|-------|-------------|
| `api-patterns` | REST、GraphQL、tRPC |
| `nestjs-expert` | NestJS 模块、DI、装饰器 |
| `nodejs-best-practices` | Node.js 异步与模块 |
| `python-patterns` | Python 规范、FastAPI |

### 数据库
| 技能 | 描述 |
|-------|-------------|
| `database-design` | 模式设计与优化 |
| `prisma-expert` | Prisma ORM 与迁移 |

### TypeScript/JavaScript
| 技能 | 描述 |
|-------|-------------|
| `typescript-expert` | 类型级编程与性能 |

### 云与基础设施
| 技能 | 描述 |
|-------|-------------|
| `docker-expert` | 容器化、Compose |
| `deployment-procedures` | CI/CD 与部署流程 |
| `server-management` | 基础设施管理 |

### 测试与质量
| 技能 | 描述 |
|-------|-------------|
| `testing-patterns` | Jest、Vitest 与策略 |
| `webapp-testing` | E2E、Playwright |
| `tdd-workflow` | 测试驱动开发 |
| `code-review-checklist` | 代码评审标准 |
| `lint-and-validate` | Lint 与校验 |

### 安全
| 技能 | 描述 |
|-------|-------------|
| `vulnerability-scanner` | 安全审计、OWASP |
| `red-team-tactics` | 进攻性安全 |

### 架构与规划
| 技能 | 描述 |
|-------|-------------|
| `app-builder` | 全栈应用脚手架 |
| `architecture` | 系统设计模式 |
| `plan-writing` | 任务规划与拆分 |
| `brainstorming` | 苏格拉底式提问 |

### 移动端
| 技能 | 描述 |
|-------|-------------|
| `mobile-design` | 移动端 UI/UX 模式 |

### 游戏开发
| 技能 | 描述 |
|-------|-------------|
| `game-development` | 游戏逻辑与机制 |

### SEO 与增长
| 技能 | 描述 |
|-------|-------------|
| `seo-fundamentals` | SEO、E-E-A-T、核心 Web 指标 |
| `geo-fundamentals` | 生成式 AI 优化 |

### Shell/CLI
| 技能 | 描述 |
|-------|-------------|
| `bash-linux` | Linux 命令与脚本 |
| `powershell-windows` | Windows PowerShell |

### 其他
| 技能 | 描述 |
|-------|-------------|
| `clean-code` | 编码规范（全局） |
| `behavioral-modes` | 代理角色 |
| `parallel-agents` | 多代理模式 |
| `mcp-builder` | 模型上下文协议 |
| `documentation-templates` | 文档格式 |
| `i18n-localization` | 国际化 |
| `performance-profiling` | Web Vitals 与优化 |
| `systematic-debugging` | 故障排查 |

---

## 🔄 工作流（11）

斜杠命令流程。使用 `/command` 调用。

| 命令 | 描述 |
|---------|-------------|
| `/brainstorm` | 苏格拉底式探索 |
| `/create` | 创建新功能 |
| `/debug` | 调试问题 |
| `/deploy` | 部署应用 |
| `/enhance` | 改进现有代码 |
| `/orchestrate` | 多代理协同 |
| `/plan` | 任务拆分 |
| `/preview` | 预览改动 |
| `/status` | 检查项目状态 |
| `/test` | 运行测试 |
| `/ui-ux-pro-max` | 使用 50 种风格设计 |

---

## 🎯 技能加载协议

```
用户请求 → 技能描述匹配 → 加载 SKILL.md
                                            ↓
                                    读取 references/
                                            ↓
                                    读取 scripts/
```

### 技能结构

```
skill-name/
├── SKILL.md           # （必需）元数据与说明
├── scripts/           # （可选）Python/Bash 脚本
├── references/        # （可选）模板、文档
└── assets/            # （可选）图片、Logo
```

### 增强技能（含 scripts/references）

| 技能 | 文件数 | 覆盖内容 |
|-------|-------|----------|
| `typescript-expert` | 5 | 工具类型、tsconfig、速查表 |
| `ui-ux-pro-max` | 27 | 50 种风格、21 套配色、50 款字体 |
| `app-builder` | 20 | 全栈应用脚手架 |

---

## 📊 统计

| 指标 | 数值 |
|--------|-------|
| **代理总数** | 16 |
| **技能总数** | 40 |
| **工作流总数** | 11 |
| **覆盖率** | ~90% Web/移动开发 |

---

## 🔗 速查

| 需求 | 代理 | 技能 |
|------|-------|--------|
| Web 应用 | `frontend-specialist` | react-patterns, nextjs-best-practices |
| API | `backend-specialist` | api-patterns, nodejs-best-practices |
| 移动端 | `mobile-developer` | mobile-design |
| 数据库 | `database-architect` | database-design, prisma-expert |
| 安全 | `security-auditor` | vulnerability-scanner |
| 测试 | `test-engineer` | testing-patterns, webapp-testing |
| 调试 | `debugger` | systematic-debugging |
| 规划 | `project-planner` | brainstorming, plan-writing |
