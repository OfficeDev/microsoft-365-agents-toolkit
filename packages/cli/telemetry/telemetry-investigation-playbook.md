# Telemetry 事件排查经验手册

本手册用于快速回答三个问题：

1. 这个 telemetry event 是什么时候发出的？
2. 有没有对应 CLI 命令？
3. 如果是自动触发，它属于哪个阶段？

## 一、推荐排查流程

### 步骤 1：先在 Kusto 找 Top 事件

先看过去 90 天 TopN，确认优先级。

示例思路：
- 过滤 eventTimestamp >= ago(90d)
- 按 name 聚合 sum(coalesce(itemCount, 1))
- 取 top 10

### 步骤 2：确认事件名前缀规则

先确认是否统一前缀上报（例如 teamsfx-cli/）。

检查位置：
- packages/cli/src/constants.ts
- packages/cli/src/telemetry/telemetryReporter.ts

目的：避免误判为多个不同事件源。

### 步骤 3：判定是命令型还是自动型

看 CLI 命令引擎是否统一发 start/finish：
- packages/cli/src/commands/engine.ts

判断原则：
- 若来自 context.command.telemetry.event 与 event-start，通常是命令型。
- 若来自 fx-core 的工具类、网络层、问答层，多为自动型阶段事件。

### 步骤 4：用事件名反查源码定义和发射点

优先搜索：
- packages/cli/src/telemetry/cliTelemetryEvents.ts
- packages/fx-core/src/common/telemetry.ts

再搜索真实发射点：
- sendTelemetryEvent(...)
- sendTelemetryErrorEvent(...)

### 步骤 5：按分层定位阶段

常见阶段与入口：
- 命令阶段：packages/cli/src/commands/engine.ts
- 网络依赖阶段：packages/fx-core/src/common/wrappedAxiosClient.ts
- 项目元数据解析阶段：packages/fx-core/src/component/utils/metadataUtil.ts
- 项目类型检查阶段：packages/fx-core/src/core/FxCore.ts
- 交互问答阶段：packages/fx-core/src/ui/visitor.ts
- 升级迁移检查阶段：packages/fx-core/src/core/middleware/projectMigratorV3.ts

## 二、事件到阶段的快速判断规则

- 以 -start 结尾：通常是“开始执行”。
- 命令名事件与命令 model 一致：多为命令型（例如 install）。
- api 类事件：优先看 wrappedAxiosClient 的 getEventName 分类逻辑。
- askquestion：高概率来自 interactive 问答树。
- project-migrator-*：高概率来自版本检查或迁移中间件。

## 三、常见误区

- 误区 1：把所有事件都当作命令。
  说明：很多是自动流程事件，会被多个命令共用。

- 误区 2：忽略统一前缀，误以为多套埋点。
  说明：可能只是同一 reporter 统一加前缀。

- 误区 3：只看成功事件，不看 start 和 error。
  说明：排查失败率时要同时看 start、success、error。

## 四、可复用查询建议

建议固定准备 4 类查询：

1. TopN 总量查询：看热点。
2. 事件按天趋势：看异常波峰。
3. 事件按 success/error 分布：看稳定性。
4. 事件按 command-name 或 interactive 分组：看来源场景。

## 五、产出模板（建议）

每次排查输出建议包含：

1. 统计口径：时间窗、计数方式、过滤条件。
2. TopN 列表：事件名与数量。
3. 映射表：事件 -> 命令/自动阶段 -> 关键代码位置。
4. 后续建议：需要新增的埋点或拆分维度。

## 六、团队实践建议

- 为高频自动事件补充 command-name、stage 等属性，减少后续排查成本。
- 对关键命令统一维护 start/success/error 三段式事件。
- 保持 event 命名风格统一，避免历史遗留别名长期共存。
