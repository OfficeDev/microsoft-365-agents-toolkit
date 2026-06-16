// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * ## Hybrid Dependency Injection Pattern for VS Code Extension
 *
 * ### 原则
 *
 * - ✅ **readonly npm 包** (fs-extra, @microsoft/teamsfx-core 等)
 *   → 集中在 `src/common/npmPackageDeps.ts`
 *
 * - ✅ **内部模块** (fileSystemUtils, appDefinitionUtils 等)
 *   → 直接 import 调用，测试中用 `vi.mock()` mock
 *
 * ### 设计目标
 *
 * 1. 最小化 deps 对象 - 只包装外部 npm 包
 * 2. 源代码清晰 - 易于理解依赖关系
 * 3. 测试灵活 - 内部模块用 vi.mock()，npm 包用 spyOn
 * 4. 减少冗余 - npm 包的 deps 集中管理，不重复定义
 *
 * ### 示例：envTreeUtils.ts
 *
 * #### 源代码
 * ```typescript
 * import { getProvisionResultJson } from "./fileSystemUtils";  // ← 内部模块：直接 import
 * import { getV3TeamsAppId } from "./appDefinitionUtils";      // ← 内部模块：直接 import
 * import { workspaceUri } from "../globalVariables";           // ← 全局变量：直接 import
 * import { fsAdapter, envParseAdapter } from "../common/npmPackageDeps";  // ← npm 包 deps
 * import path from "path";
 *
 * export async function getM365TenantFromEnv(env: string) {
 *   const projectPath = workspaceUri?.fsPath || "";
 *   const envFile = path.resolve(projectPath, "env", `.env.${env}`);
 *
 *   // ← 只有 npm 包调用才用 adapter
 *   if (await fsAdapter.pathExists(envFile)) {
 *     const content = fsAdapter.readFileSync(envFile, "utf-8");
 *     const envData = envParseAdapter.deserializeDotenv(content);
 *     return envData.obj["TEAMS_APP_TENANT_ID"];
 *   }
 *   return undefined;
 * }
 * ```
 *
 * #### 测试代码
 * ```typescript
 * import * as fileSystemUtils from "../../src/utils/fileSystemUtils";
 * import { fsAdapter, envParseAdapter } from "../../src/common/npmPackageDeps";
 *
 * // ✅ Mock 内部模块
 * vi.mock("../../src/utils/fileSystemUtils");
 *
 * describe("getM365TenantFromEnv", () => {
 *   it("returns m365 tenantId successfully", async () => {
 *     // 内部模块：用 vi.mocked()
 *     vi.mocked(fileSystemUtils.getProvisionResultJson).mockResolvedValue(...);
 *
 *     // npm 包：用 spyOn
 *     vi.spyOn(fsAdapter, "pathExists").mockResolvedValue(true);
 *     vi.spyOn(envParseAdapter, "deserializeDotenv").mockReturnValue(...);
 *
 *     const result = await envTreeUtils.getM365TenantFromEnv("test");
 *     chai.expect(result).equal("fakeTenantId");
 *   });
 * });
 * ```
 *
 * ### 迁移步骤
 *
 * #### 第 1 步：识别 npm 包
 * 在要重构的模块中，找出所有 npm 包的使用：
 * ```typescript
 * import fs from "fs-extra";               // ← npm 包
 * import { dotenvUtil } from "@microsoft/teamsfx-core/...";  // ← npm 包
 * import { getProvisionResultJson } from "./fileSystemUtils"; // ← 内部模块
 * ```
 *
 * #### 第 2 步：在 npmPackageDeps.ts 中定义 adapter
 * ```typescript
 * export const fsAdapter = {
 *   pathExists: (filePath: string) => fs.pathExists(filePath),
 *   readFileSync: (filePath: string, encoding: BufferEncoding) =>
 *     fs.readFileSync(filePath, encoding),
 *   // ... 其他方法
 * };
 * ```
 *
 * #### 第 3 步：更新源代码
 * ```typescript
 * // 移除：import fs from "fs-extra";
 * // 添加：import { fsAdapter } from "../common/npmPackageDeps";
 *
 * // 改：fs.pathExists(path)
 * // 成：fsAdapter.pathExists(path)
 * ```
 *
 * #### 第 4 步：更新测试代码
 * ```typescript
 * // 添加 vi.mock() 来 mock 内部模块
 * vi.mock("../../src/utils/fileSystemUtils");
 *
 * // 改：vi.spyOn(envTreeUtilsDeps, "pathExists")
 * // 成：vi.spyOn(fsAdapter, "pathExists")
 * ```
 *
 * ### 常见陷阱
 *
 * ❌ **错误**：每个模块都定义一份 npm 包 deps
 * ```typescript
 * // src/utils/envTreeUtils.ts
 * export const envTreeUtilsDeps = {
 *   pathExists: (filePath) => fs.pathExists(filePath),
 * };
 *
 * // src/utils/mcpUtils.ts
 * export const mcpUtilsDeps = {
 *   pathExists: (filePath) => fs.pathExists(filePath),  // ← 重复了！
 * };
 * ```
 *
 * ✅ **正确**：集中在 npmPackageDeps.ts
 * ```typescript
 * // src/common/npmPackageDeps.ts
 * export const fsAdapter = {
 *   pathExists: (filePath) => fs.pathExists(filePath),
 * };
 *
 * // src/utils/envTreeUtils.ts
 * import { fsAdapter } from "../common/npmPackageDeps";
 *
 * // src/utils/mcpUtils.ts
 * import { fsAdapter } from "../common/npmPackageDeps";  // ← 复用同一个
 * ```
 *
 * ### 验证迁移成功
 *
 * ✅ 源代码
 * - [ ] 没有 `import fs from "fs-extra"` 等 npm 包
 * - [ ] 只有 `import { fsAdapter, ... } from "../common/npmPackageDeps"`
 * - [ ] 内部模块直接 import：`import { helper } from "./xxx"`
 *
 * ✅ 测试代码
 * - [ ] 有 `vi.mock("../../src/utils/xxx")` 来 mock 内部模块
 * - [ ] 使用 `vi.mocked(...).mockReturnValue()` 来 mock 内部模块
 * - [ ] 使用 `vi.spyOn(fsAdapter, ...)` 来 mock npm 包
 * - [ ] 没有 `envTreeUtilsDeps` 对象
 *
 * ✅ 测试通过
 * - [ ] `npm run test:unit` 该模块的测试全部通过
 * - [ ] `npm run build` 无编译错误
 *
 */

export const migrationGuide = "See comments above";
