# plan.json（schema_version=2，V3设计与V5展示扩展）

AI读真实需求后填写UTF-8 JSON；脚本不设计业务规则。必需顶层：schema_version为2、title、requirement_version、source_review、rules、coverage_items、cases。questions可为空数组；scope_notes/unknowns/failed_attempts只写本轮事实。

## 字段

- rules[]：id稳定且唯一；statement独立规则；kind为需求验证/边界派生/补充建议/待确认；evidence[]真实需求证据；rationale支持/推导/建议理由；questions[]Q。待确认性质与非空Q同时存在。
- coverage_items[]：id、rule_id、method、target、rationale、disposition。disposition为planned/deferred/excluded；后两种必须reason，不能挂执行用例。每条规则至少一项，planned至少一条用例。
- cases[]：id如TC-001、path[]原业务父路径（V5另用display_path展示）、title、coverage_ids[]、preconditions[]、test_data[]、steps[]、postconditions[]、impact、priority（P0/P1/P2）、priority_reason、verification、kind、evidence[]、ideas[]、questions[]。无适用思路时ideas为空并逐条写ideas_reason，scope_notes说明范围。
- test_data[]每项字符串name/value/basis，basis说明来源或构造方式；无需额外数据用空数组。
- steps[]每项字符串action/expected/observe，observe写观察位置和时点。
- cases[].xmind_points：可选的非空字符串数组。新设计逐条提供；省略时兼容旧详细展示并提示补写，空数组/空白项/非字符串均拒绝。属于展示信息，不参与业务设计变更指纹。
- verification含字符串method、非空requirements[]、可为空limitations[]。这是准备要求，不是实际方式或证据。
- questions[]含id如Q01、title/gap/ask/evidence[]。用例须包含关联规则全部Q与需求证据，可附更多。

方法枚举：直接验证、等价类、边界值、决策表、状态迁移、风险经验。中间四种必需model，下列字段值均为非空字符串，可包含组合或明确的待确认描述：

| 方法 | model字段 |
|---|---|
| 等价类 | partition、representative |
| 边界值 | threshold、unit、step、selected_values |
| 决策表 | conditions、outcome |
| 状态迁移 | from_state、event、to_state |

脚本派生action/expected/rule_ids供两格式统一显示，不手工维护两套。下划线开头保留给脚本。未执行不得填status/actual/current_code/result/acceptance_status。

V2 2.3.0以来完整action/expected存入Excel检查详情及XMind备注；主表需求简述复用xmind_points，缺少时用title，不截断完整预期。verification及执行准备仍用于后续人工测试设计，不要求其他AI完成实机才可判实现检查通过。设计模型与业务指纹不因表格布局改变。

## 合成最小示例

仅示范结构；R000001和COM-001必须替换为本次真实编号，不能把示例规则作为需求。

```json
{
  "schema_version": 2,
  "title": "合成按钮需求",
  "requirement_version": "合成v1",
  "source_review": "一段文字，无图片/修订；全文已读",
  "rules": [{"id": "RULE-001", "statement": "点击说明按钮显示指定说明", "kind": "需求验证", "evidence": ["R000001"], "rationale": "原文明示", "questions": []}],
  "coverage_items": [{"id": "COV-001", "rule_id": "RULE-001", "method": "直接验证", "target": "打开指定说明", "rationale": "单一交互", "disposition": "planned"}],
  "cases": [{
    "id": "TC-001", "path": ["说明界面"], "title": "点击说明按钮", "xmind_points": ["显示需求指定的说明内容"],
    "coverage_ids": ["COV-001"], "preconditions": ["进入含说明按钮的界面"], "test_data": [],
    "steps": [{"action": "点击说明按钮", "expected": "显示指定说明", "observe": "打开后的文字与原文对照"}],
    "postconditions": ["关闭说明返回原界面"], "impact": "核对返回后可操作",
    "priority": "P1", "priority_reason": "功能交互",
    "verification": {"method": "运行实测", "requirements": ["可进入界面的测试角色，当前未确认是否具备"], "limitations": []},
    "kind": "需求验证", "evidence": ["R000001"], "ideas": ["COM-001"], "questions": []
  }],
  "questions": [], "scope_notes": ["合成示例"], "unknowns": [], "failed_attempts": []
}
```

## XMind简述编写

先完成规则、覆盖、详细用例，再按每条用例的实际业务意思编写简述。通常1条，确有独立观察内容时2～3条；不要把详细步骤逐条搬到导图。标题已说明场景时，简述直接说明结果，避免重复。固定的“条件/操作”“预期/判定”“影响核对”不是简述节点。

- 保留决定判定的数字、阈值及等号方向、时间、对象、作用范围、上限、例外和概率；省去重复准备、观察位置和通用提醒。多步骤只保留理解场景所需的关键先后关系。
- 不机械删除“可能”等词；概率结果不能写成必然。缺规则时简述写清要确认什么，Q编号由导出器放在测试点标题中，不再增加待确认子分支。
- 简述与标题合读仍须准确；过长时先检查场景是否可独立拆分，不为字数截断原文。脚本对超过3条或单条超过80字仅给审阅提示，不自动改写。
- 前置、数据及构造依据、逐步操作/预期/观察点、清理、准备要求与证据必须保留原字段；导出器写入测试点备注及Excel。简述不是执行判定的替代品。

例如已获原文支持的“满血时使用”可写 `["提示‘当前生命值已满’，不执行回血"]`；“复活次数”待确认时可写 `["确认次数限制按每轮还是每期计算"]`。这些只是表达示例，不是任何新需求的业务依据。

## 覆盖口径

报告已登记规则数、覆盖项总数、planned/deferred/excluded数及待确认planned数，均为设计计数。planned含待确认项不能称为可直接验收。脚本不保证检出遗漏规则、错误预期、无意义重复或错误分区，仍须内容审查。

若全无可设计用例，先交付缺口说明，不用占位用例强行导出。source_review如实记录未读图片、版本冲突，不为过校验虚填全文已读。

## V3新设计字段与严格校验

格式仍为schema_version=2。新增可选字段保证旧计划可读，但V3新设计必须填写，使用--require-v3校验：

- cases[].review_focus：对象，target只能是“服务端”或“客户端简查”，checks是非空字符串数组，每项写具体业务核对目标。未知属性不代替必需项。格式错误即拒绝，不静默降级。
- questions[].impact：非空字符串，一句话说明问题影响哪些条件或结果。旧问题缺字段不补写；新问题必须填写。
- 不改steps、verification等原字段契约。checks补充静态核对路径，steps保留具体条件、检查动作与预期；requirements写需取得的材料和证据类型。客户端样例按实际需要保留运行观察。
- 旧计划省略--require-v3可导出；报告列出缺少目标或影响说明的条目，不猜归属。旧图--allow-legacy与--require-v3不可同时使用。
- review_focus与问题impact参与设计指纹，改核对项或影响说明会触发关联用例复核。xmind_points不参与指纹。旧字段完全不变且未增加新字段时，基线指纹保持一致。
- 无V5展示配置的兼容导出仅对展示排序，在模块内将客户端简查移至末尾独立分支；不修改输入path或ID。Excel详情中的导图路径使用实际展示路径。

服务端检查字段示例（仅合成表达，必须替换为本次真实规则）：

```json
{"review_focus":{"target":"服务端","checks":["积分阈值为100，边界比较包含100","未领取且积分达到100时发放1份奖励并保存领取状态","同一角色再次领取不增加奖励"]}}
```

旧章节的按钮示例展示兼容格式。将它作为V3新设计使用时，补review_focus：target为客户端简查、checks为“核对按钮打开的说明与需求一致”；不能把纯按钮显示重命名成服务端逻辑。

## 待确认写法

一个Q只问一个可回答的问题，同一问题多个用例引用，不反复复制。title直接提问；gap先说原文明确什么、缺什么；ask用具体场景说明要选择的规则；impact说影响哪个业务判断。示例数字源于需求，或明确标注“假设示例”，不能从例子创造正式规则。原文引用留证据，不机械截断。

示例（假设需求只写累计满50元）：
- title：充值金额是否跨天累计？
- gap：需求写累计满50元可领取，没有写跨天是否清零。
- ask：假设示例：昨天充30元、今天充20元，是否可以领取50元档奖励？
- impact：影响累计达标判断和奖励领取。

避免只写“累计维度/结算口径待确认”，也不要给没有依据的默认答案。问题页仍展示原文/缺口、需确认、影响说明与关联用例；测试点简述只问需要确认的事情，Q编号由导出器显示。

“源码未取得”是验收受阻原因，“只找到部分处理”是部分验收，“没有实机环境”是准备限制；这些情况不能自动生成需求Q。“未找到实现”必须说明实际搜索范围，不能断言没有实现。

## 3.0.1 最短验收路径

cases[].acceptance_path是非空字符串，写入Excel主表H“最短验收路径”，替代原“后续测试重点”。3.0.1新设计必须逐例填写，--require-v3会拒绝缺失或空白；兼容旧计划允许省略，H留空且报告cases_without_acceptance_path，不编造路径。其他地方不维护第二份相同短路径，完整步骤仍在原字段中；XMind短描述不变。

先完成详细用例，再写“必要前提 → 游戏内操作 → 关键观察结果”，通常一行、必要时两行。保留最少但足够的前置、关键触发和观察；需要独立判断的不同场景仍拆例。优先已有角色与正常入口，不强行省掉准备数据、跨日等待或记录操作前资产。未确认入口写“从已确认的活动入口进入（具体路径待补）”；没有资料不能发明菜单、GM命令或加速能力。状态始终按真实检查证据回填。

路径必须对应本条测试点：纯配置检查可写游戏内抽查效果，但注明不能替代配置逐项核对；重复回调、并发等游戏内无法直接触发时，写明“游戏内无法直接触发；需准备……”，可附可观察的玩家侧结果，不能把连点当作重复回调证明。缺环境是准备限制，不自动生成需求Q。

待确认场景只写准备、触发、记录什么，以及“待Q编号确认后判定”，不预设清零、补发等未定结果。短路径由设计者填写，后续AI可按已证实入口与条件修正；禁止把未做过的步骤写成已完成记录。

示例仅用于写法，机制、数据均须按本次需求替换：
- 普通：需准备积分100且未领奖角色 → 打开活动领取 → 检查奖励增加1份、状态变为已领取。
- 消耗不足：需准备比单次消耗少1个的材料 → 尝试操作 → 检查操作失败、材料与奖励不变。
- 跨日待确认：需准备昨天60积分且未领奖角色 → 跨天后增加40积分 → 记录积分和领奖状态，待Q01确认后判定。
- 无直接游戏路径：游戏内无法直接触发重复支付回调；需回调工具和测试订单，重放后在游戏内检查奖励仅到账一次（不能用连点代替）。

acceptance_path是从详细设计派生的阅读指引，与xmind_points一样不纳入业务设计指纹。改变业务条件、步骤或预期时必须同步更新原详细字段并由其触发变更复核，不能只改短路径掩盖设计变化。旧计划指纹不变；历史工作簿不迁移、不覆盖。


## V5 展示字段与严格校验

新V5设计使用`--require-v5`，格式仍为schema_version=2。它同时检查V3必填内容及xmind_points，并要求：

- 顶层template_profile：活动、玩法、养成、英雄、通用之一；按本次需求选，不依赖桌面目录。
- cases[].display_group：入口、红点、主功能、特殊点验证、异常状态、配置检查、UI检查之一。
- cases[].display_path：非空字符串数组，保存实际导图父路径。主功能以具体业务名称开头，如["积分领奖","领取条件"]；其余以分类名称开头，如["异常状态","重复请求"]。不使用主功能1、次级模块、功能点1等占位词。
- target为客户端简查的用例必须归入UI检查，如["UI检查","积分领奖"]。服务端返回资格、金额、奖励、时间等数据仍归实际业务分类，不与纯画面合并。
- 可选template_decisions：分类名到取舍理由的字典，新设计记录各分类适用或不适用的理由；未提供时辅助报告提示复核，不凭空补造理由。

同一用例只有一个主要归属。按分类顺序、主功能首次出现顺序和路径树顺序输出，原ID、path和详细设计不修改。主表模块和总览分类使用display_path[0]；详情保存实际完整导图路径，主表与详情仍按ID互跳。深层分组折叠，测试点展开后读简述。

template_profile、template_decisions、display_group、display_path均不参与业务设计指纹。禁止只改展示摘要或路径掩盖真实条件/预期变化；实际设计改动仍更新详细字段。所有用例都缺展示字段且无template_profile时保留旧组织并提示；部分填写时拒绝导出，避免混用分类。旧导图--allow-legacy保真转换不接受V5展示字段，不能自动套模板。详见[模板展示规则](template-presentation.md)。
