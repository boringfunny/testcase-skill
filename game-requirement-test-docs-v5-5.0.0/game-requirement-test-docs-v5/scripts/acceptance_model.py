"""Implementation review contract; evidence completeness is not a correctness oracle."""
STATUSES = ['【未开始】','【部分验收】','【验收受阻】','【待确认】','【未通过】','【已通过】','【不适用】']
LEGACY_STATUSES = ['【待实测】','【已验收】','【其他见备注】']
LAYER_RESULTS = ['未验证','部分验证','符合','不符','待确认','不适用']
REASONS = ['处理链路未查清','检查点未完成','源码或配置缺失','账号或权限缺失','工具能力未确认','外部依赖阻断','规则未明确','范围不适用']
SUMMARY_HEADERS = ['模块','总数','未填写']+[s.strip('【】') for s in STATUSES]+['旧待实测','旧已验收','旧其他见备注','反馈已填写','无效状态']
SUMMARY_WIDTHS = [160,70,85]+[90]*7+[95,95,120,115,95]
INSTRUCTIONS = [
    '用途：AI检查实现，为人工测试提供依据。【已通过】表示关键条件和结果有代码或配置证据支持，不要求实机；不代表游戏已实测通过。',
    '填写：主表回填状态和1～2句实现说明；最短验收路径已由设计提供，可按已确认条件修正。详情写证据、检查范围、三层结果；共用方法/版本/人员/时间在下方填一次，逐条例外覆盖。',
    '判定：链路不完整→部分验收；材料不足→验收受阻；规则不明→待确认；明确不符→未通过。搜索命中不能直接通过，没找到不能直接断言未实现。',
    '检查结果初始为空，Q预设待确认。路径只说明如何测，不代表已执行；未采用层填未验证，静态通过无需先完成路径。旧状态逐条复核，不自动迁移。',
    '反馈提示只检查缺填与明显矛盾；证据须具体到文件/行号/配置项及版本。回填后重算汇总；待确认批准写详情，新增Q同步问题页及主表问题编号。',
]


def formula(r, refs=None):
    """Logical columns: D status, K finding, O questions; I/L/N/T/U/V/R/X/M/Y/Z/AA detail.

    refs optionally maps logical columns to Excel expressions (ID lookups/shared baseline).
    Without refs, simple cell names support independent synthetic evaluation.
    """
    refs=refs or {}
    def cell(c): return refs.get(c,f'{c}{r}')
    def eq(c,v): return f'{cell(c)}="{v}"'
    def ne(c,v): return f'{cell(c)}<>"{v}"'
    def OR(*xs): return 'OR('+','.join(xs)+')'
    def AND(*xs): return 'AND('+','.join(xs)+')'
    def NOT(x): return f'IF({x},"否","是")="是"'
    def state(s): return eq('D','【'+s+'】')
    def missing(*cols): return OR(*(eq(c,'') for c in cols))
    layers=['Y','Z','AA']
    failure=OR(*(eq(c,'不符') for c in layers))
    uncertain=OR(*(eq(c,'待确认') for c in layers))
    partial=OR(*(eq(c,'部分验证') for c in layers))
    found=OR(*(eq(c,'符合') for c in layers))
    progress=OR(found,partial,failure)
    guards=[
        (eq('D',''),'未填写状态'),
        (AND(*(ne('D',s) for s in STATUSES+LEGACY_STATUSES)),'状态值无效'),
        (OR(*(eq('D',s) for s in LEGACY_STATUSES)),'旧状态待逐条复核'),
        (OR(*(AND(ne(c,''),*(ne(c,v) for v in LAYER_RESULTS)) for c in layers)),'分层结果值无效'),
        (AND(failure,ne('D','【未通过】')),'状态冲突：已记录不符'),
        (AND(state('未开始'),progress),'状态冲突：已有检查结果'),
        (state('未开始'),'尚未开始'),
        (missing(*layers),'待补三层结果（未采用填未验证）'),
        (missing('I','T','U','V','N'),'待补检查方式/范围/基线'),
        (missing('K','L'),'待补实现说明/定位证据'),
        (AND(state('已通过'),uncertain),'状态冲突：规则仍待确认'),
        (AND(state('已通过'),ne('O',''),eq('R','')),'待补规则批准记录'),
        (AND(state('已通过'),OR(partial,NOT(found))),'待补关键处理的完整依据'),
        (AND(state('未通过'),NOT(failure)),'待补不符层及差异'),
        (AND(state('部分验收'),NOT(progress)),'待补已检查内容'),
        (AND(state('不适用'),OR(*(ne(c,'不适用') for c in layers))),'状态冲突：整体不适用须说明各层'),
        (AND(OR(*(state(s) for s in ['部分验收','验收受阻','待确认','不适用'])),missing('X','M')),'待补具体缺口/原因/下一步'),
    ]
    result='"反馈已填写"'
    for condition,message in reversed(guards):result=f'IF({condition},"{message}",{result})'
    return '='+result
