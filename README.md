# 游戏需求测试文档技能 V5

`game-requirement-test-docs-v5` 根据游戏需求和通用测试思路设计测试用例，生成配套的 XMind 导图和 Excel AI 实现检查表。它以服务端配置、条件判断、请求处理、状态及资产结果的核对为主，并为纯客户端表现附上简要检查点。适用于测试设计和后续验收交接。

**本技能只设计和导出文档。**它不会检查游戏源码或配置的实际实现，不会操作游戏，也不会把未执行的检查预填为通过。

## 交付内容

| 文件 | 用途 |
| --- | --- |
| `<需求名称>-测试用例.xmind` | 按业务模块整理测试点，用简短描述呈现场景和预期；完整检查项、执行设计与证据保存在备注中。 |
| `<需求名称>-AI验收表.xlsx` | 供后续 AI 核对实现、人工安排测试。包含检查目标、具体检查项、证据、待确认问题及最短验收路径；检查结果留待实际验收回填。 |

Excel 包含七个工作表，默认打开“AI实现检查”。其中 F/G 列用于回填总状态和实现说明，H 列预填设计阶段编写的最短验收路径。导出目录的 `.support` 保存设计基线、验证报告和辅助预览。

## 输入与设计方式

- **需求资料**：以用户选定的需求为准，核对正文、相关图片、批注和修订；资料缺失或含义不明时记录待确认问题，不凭标题补规则。
- **通用测试思路**：优先使用用户指定文件，其次使用工作区的 `outputs/**/通用游戏测试思路.xlsx`，最后使用随包快照。先判断条目是否适用，再据此规划用例。
- **历史导图**：只参考层次、样式和表达，不沿用历史需求的业务规则。随包提供样本和四类格式模板；模板内容也不能直接当作当前需求。
- **测试方法**：按需求选择直接验证、等价类、边界值、决策表、状态迁移或风险经验，记录选择理由，不为凑数量增加用例。
- **检查重点**：服务端用例写清配置和逻辑核对点；纯客户端布局、文字、动画等归入 UI 简查。每例提供“必要前提 → 游戏内操作 → 关键观察结果”的最短验收路径，但路径不代表已执行。
- **待确认事项**：分别写明已知事实、缺口、具体问题及影响。来源证据、设计推导和建议保持区分，不编造配置表、字段、接口、函数、游戏入口或 GM 命令。

## 安装与调用

仓库同时提供[解压后的技能目录](game-requirement-test-docs-v5-5.0.0/game-requirement-test-docs-v5/)和[发行 ZIP](game-requirement-test-docs-v5-5.0.0.zip)。在仓库根目录运行安装脚本，可将整个技能复制到默认的 Codex `skills` 目录；目标已有同名技能时，脚本会拒绝覆盖：

```powershell
python game-requirement-test-docs-v5-5.0.0/game-requirement-test-docs-v5/scripts/install.py
```

也可以使用 `--skills-dir <技能父目录>` 指定安装位置。安装后，在 Codex 中调用 `$game-requirement-test-docs-v5`，并提供需求文件；如有指定的通用测试思路或历史导图，也一并提供。

文档导出采用 `plan.json` 和 `evidence.json`。新 V5 设计使用严格校验：

```text
python <技能目录>/scripts/build_documents.py --plan <工作目录>/plan.json --evidence <工作目录>/evidence.json --out <新输出目录> --require-v5
```

输出目录必须是新的目录，避免覆盖既有文档及历史回填。旧计划可使用兼容模式；旧导图转换、需求变更基线和稳定用例 ID 的规则见[兼容与变更说明](game-requirement-test-docs-v5-5.0.0/game-requirement-test-docs-v5/references/change-and-migration.md)。V5 与原版、V2、V3、V4 独立共存。

## 依赖与验证

生成 Excel 和导图需要 `xlsxwriter`、`openpyxl`、`Pillow`；读取 PDF 另需 `pypdf`。维护技能后可运行：

```text
python <技能目录>/scripts/self_test.py --work-dir <新的测试目录>
```

结构自测只能检查文档和数据结构，不能替代原生软件兼容性检查、需求语义复核或真实游戏验收。详细规则见[技能说明](game-requirement-test-docs-v5-5.0.0/game-requirement-test-docs-v5/SKILL.md)、[计划契约](game-requirement-test-docs-v5-5.0.0/game-requirement-test-docs-v5/references/plan-contract.md)及[验收约定](game-requirement-test-docs-v5-5.0.0/game-requirement-test-docs-v5/references/acceptance-contract.md)。
