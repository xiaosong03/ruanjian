# -*- coding: utf-8 -*-
"""
软件测试用例批量生成工具（字段驱动版）
- 界面识别 Word 模板表格中的全部字段，用户填写公共部分（每张表相同）
- 未填写的字段从 Excel(.xlsx) 读取，按第一行表头匹配模板字段，一行 = 一个用例
- 界面优先：界面有值的字段用界面值；界面未填的才用 Excel
- 输出一个新的 Word，包含 N 张深拷贝的模板表
"""

import base64
import copy
import io
import json
import os
import sys

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

from _embedded_template import TEMPLATE_BASE64


# ---------------------------------------------------------------------------
# 内嵌模板
# ---------------------------------------------------------------------------
def _load_template():
    """从内嵌的 base64 模板解码为可读的字节流（不写临时文件）。"""
    return io.BytesIO(base64.b64decode(TEMPLATE_BASE64))


# ---------------------------------------------------------------------------
# 模板配置记忆（一次性校正复用，用户无需改代码/重打包）
# ---------------------------------------------------------------------------
# 配置以模板路径为键，存的是用户对某模板的手工校正：
#   detail_col_roles : { 明细列表头文本 : 角色(step/期望/准则/实际/ignore) }
#   excel_col_map    : { Excel表头文本 : 目标(模板字段label 或 明细角色 或 ignore) }
# 公共字段的自动识别始终不需要任何配置；这些配置只用于"自动判不准"的列。
def _config_base_dir():
    if getattr(sys, 'frozen', False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return base if os.access(base, os.W_OK) else os.path.expanduser('~')


CONFIG_PATH = os.path.join(_config_base_dir(), '测试用例生成器配置.json')


def template_key(template_source):
    """模板的配置键。None/空串=内嵌默认模板。"""
    if template_source is None or str(template_source).strip() == '':
        return '__embedded__'
    return os.path.normpath(str(template_source))


def _read_config_all():
    try:
        with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
            data = json.load(f)
        if isinstance(data, dict):
            return data
    except Exception:
        pass
    return {}


def load_template_config(template_source):
    """读取某模板的配置，返回 (key, cfg)。
    cfg = {'detail_col_roles': {...}, 'excel_col_map': {...}}，读写使用同一配置对象。"""
    key = template_key(template_source)
    tpl = _read_config_all().get('templates', {}).get(key, {})
    if not isinstance(tpl, dict):
        tpl = {}
    cfg = {'detail_col_roles': dict(tpl.get('detail_col_roles') or {}),
           'excel_col_map': dict(tpl.get('excel_col_map') or {})}
    return key, cfg


def _write_config_all(allc):
    try:
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            json.dump(allc, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def save_template_config(template_source, new_fields):
    """将 new_fields(部分字段) 合并写回某模板的配置。返回配置键。"""
    key, cfg = load_template_config(template_source)
    for k, v in new_fields.items():
        if isinstance(v, dict):
            cfg.setdefault(k, {}).update(dict(v))
        else:
            cfg[k] = v
    allc = _read_config_all()
    allc.setdefault('templates', {})[key] = cfg
    _write_config_all(allc)
    return key


# ---------------------------------------------------------------------------
# 字段/表头别名定义
# ---------------------------------------------------------------------------
# 语义等价的字段标签分组（用于 Excel 表头与模板字段的匹配）
ALIAS_GROUPS = [
    {'key': 'name',        'labels': ['测试用例名称', '用例名称', '名称', '用例名']},
    {'key': 'case_id',     'labels': ['测试用例标识', '用例标识', '用例编号', '用例编号', '标识', 'ID']},
    {'key': 'desc',        'labels': ['测试用例描述', '用例描述', '描述', '测试综述']},
    {'key': 'trace',       'labels': ['需求追踪', '需求编号']},
    {'key': 'init',        'labels': ['测试初始化要求', '初始化要求', '初始化条件', '测试初始化条件']},
    {'key': 'prereq',      'labels': ['前提和约束', '前提约束', '前置条件', '前提条件', '约束']},
    {'key': 'termination', 'labels': ['终止条件', '测试终止条件', '测试终止', '测试终止要求']},
    {'key': 'tester',      'labels': ['测试人员', '测试员', '编制人']},
    {'key': 'monitor',     'labels': ['监测人员', '检测人员', '检测员', '监控人员']},
    {'key': 'remark',      'labels': ['备注', '备注说明', '补充说明']},
]

# 测试过程明细表头各列可用标签
STEPS_COL_ALIASES = {
    'seq':      ['序号', '编号', '步骤编号'],
    'step':     ['测试步骤', '测试输入及测试步骤', '测试输入及步骤', '测试输入/步骤',
                 '测试输入', '输入及操作步骤', '步骤', '操作步骤'],
    'expected': ['期望测试结果', '期望结果', '期望'],
    'actual':   ['实际测试结果', '实际结果', '实际'],
    'criterion': ['评价准则', '评分准则', '评价标准', '评分标准', '评价'],
}

# 明细列角色的人名显示（界面下拉、生成说明用）
ROLE_NAMES = {'seq': '序号', 'step': '步骤', 'expected': '期望结果',
              'actual': '实际结果', 'criterion': '评价准则'}

# 明细区结束标记行（作为锚点插入新步骤行之后）
STOP_LABELS = ['测试结果', '测试结论', '测评结论', '执行结果', '结论']

# 不作为可填写字段的标签（小节标题 / 表头词 / 列头词）
_SECTION_NAMES = {'测试过程', '测试过程描述', '测试过程/步骤', '测试相关', '测评相关'}
_NON_FIELD_LABELS = _SECTION_NAMES | set(STOP_LABELS) | \
    {'序号', '编号', '评价准则', '评价', '实际执行结果'}


def _key_for_label(label):
    """返回标签对应的语义 key（无匹配返回 None）。"""
    for g in ALIAS_GROUPS:
        if label in g['labels']:
            return g['key']
    return None


# ---------------------------------------------------------------------------
# 工具函数
# ---------------------------------------------------------------------------
def _tc_text(tc):
    """取单元格内所有文本并规整空白。"""
    return ''.join(t.text or '' for t in tc.iter(qn('w:t'))).strip()


def _tr_tcs(tr):
    return tr.findall(qn('w:tc'))


# 填写文本统一使用 宋体 小五(9pt)，半磅值 sz/szCs 均为 18
_FONT_RPR = (
    '<w:rPr %s>'
    '<w:rFonts w:ascii="宋体" w:hAnsi="宋体" w:eastAsia="宋体"/>'
    '<w:sz w:val="18"/><w:szCs w:val="18"/>'
    '</w:rPr>'
) % nsdecls('w')


def _style_run(r, text):
    """设置 run 内容为 text，并强制为 宋体小五。"""
    for child in list(r):
        if child.tag in (qn('w:rPr'), qn('w:t')):
            r.remove(child)
    r.append(parse_xml(_FONT_RPR))
    t = r.makeelement(qn('w:t'), {})
    t.text = text
    r.append(t)


def _set_tc_text(tc, text):
    """设置某个 <w:tc> 内第一个段落第一个 run 的文本，字体统一为 宋体小五。"""
    paragraphs = tc.findall(qn('w:p'))
    if paragraphs:
        p = paragraphs[0]
    else:
        p = tc.makeelement(qn('w:p'), {})
        tc.append(p)
    runs = p.findall(qn('w:r'))
    if not text:
        for r in runs:
            p.remove(r)
        return
    if runs:
        r = runs[0]
        for extra in runs[1:]:
            p.remove(extra)
    else:
        r = p.makeelement(qn('w:r'), {})
        p.append(r)
    _style_run(r, text)
    p.append(r)


def _set_tc_lines(tc, text):
    """按换行将 text 拆成多段，每行一个段落写入单元格，字体统一为 宋体小五。"""
    paragraphs = tc.findall(qn('w:p'))
    if not text:
        for p in paragraphs:
            tc.remove(p)
        return
    if paragraphs:
        template_p = paragraphs[0]
    else:
        template_p = tc.makeelement(qn('w:p'), {})
    for p in paragraphs[1:]:
        tc.remove(p)

    lines = [ln for ln in str(text).replace('\r\n', '\n').replace('\r', '\n').split('\n')]
    for i, line in enumerate(lines):
        p = template_p if i == 0 else copy.deepcopy(template_p)
        for r in p.findall(qn('w:r')):
            p.remove(r)
        r = p.makeelement(qn('w:r'), {})
        _style_run(r, line)
        p.append(r)
        if i > 0:
            tc.append(p)


def _put_tc(tc, text):
    """按是否含换行决定单行/多行写入。"""
    text = text if isinstance(text, str) else (str(text) if text is not None else '')
    text = text.replace('\r\n', '\n').replace('\r', '\n')
    if '\n' in text:
        _set_tc_lines(tc, text)
    else:
        _set_tc_text(tc, text)


def _is_field_label(text):
    t = text.strip()
    if not t or len(t) > 20:
        return False
    if t in _NON_FIELD_LABELS:
        return False
    if t[0].isdigit():
        return False
    return True


def trs_last(tbl):
    trs = tbl.findall(qn('w:tr'))
    return trs[-1] if trs else None


# ---------------------------------------------------------------------------
# 模板结构分析（字段 + 明细区）
# ---------------------------------------------------------------------------
def analyze_fields(tbl):
    """分析模板第一张表格。返回 dict:
      fields: [ {'label','tc','key'} ]  按出现顺序的填写字段
      steps:  {header_tr, pattern_tr, anchor_tr, seq_col, step_col, exp_col, act_col}
      warnings: list[str]
    """
    trs = tbl.findall(qn('w:tr'))
    warnings = []

    # ---- 识别测试过程明细表头 ----
    header_index = None
    header_cols = {}
    for hi, tr in enumerate(trs):
        tcs = _tr_tcs(tr)
        texts = [_tc_text(tc) for tc in tcs]
        col = {role: next((j for j, t in enumerate(texts) if t in aliases), None)
               for role, aliases in STEPS_COL_ALIASES.items()}
        if col.get('step') is not None and col.get('expected') is not None:
            header_index = hi
            header_cols = col
            break
    if header_index is None:
        warnings.append('未找到"步骤/期望结果"明细表头，不会填充测试过程。')
    step_col = header_cols.get('step')

    # ---- 提取填写字段（跳过明细表头与其数据行） ----
    fields = []
    seen = set()
    for ri, tr in enumerate(trs):
        if header_index is not None and ri == header_index:
            continue  # 明细表头行不是字段
        tcs = _tr_tcs(tr)
        texts = [_tc_text(tc) for tc in tcs]
        if header_index is not None and ri > header_index:
            # 明细数据行：首格为序号数字，或步骤列有值且首格不是字段标签
            is_data = False
            if step_col is not None and step_col < len(texts) and texts[step_col]:
                first = texts[0] if texts else ''
                if first and first[0].isdigit():
                    is_data = True
                elif not _is_field_label(first):
                    is_data = True
            if is_data:
                continue  # 明细数据行
        i = 0
        while i < len(tcs) - 1:
            label = _tc_text(tcs[i]).strip()
            if _is_field_label(label) and label not in seen:
                seen.add(label)
                fields.append({'label': label, 'tc': tcs[i + 1], 'key': _key_for_label(label)})
                i += 2
            else:
                i += 1

    # ---- 明细区锚点/示例行 ----
    if header_index is None:
        steps = {'header_tr': None, 'pattern_tr': None, 'anchor_tr': trs[-1] if trs else None,
                 'seq_col': None, 'step_col': None, 'exp_col': None, 'act_col': None,
                 'crit_col': None, 'section': None, 'detail_headers': [], 'detail_roles': []}
    else:
        # 分区标题（明细表头上方整行为"测试过程/测试过程描述"的标题行）
        section = None
        for tr in trs[:header_index]:
            texts = [_tc_text(tc) for tc in _tr_tcs(tr)]
            nonempty = [t for t in texts if t]
            if nonempty and all(t in _SECTION_NAMES for t in nonempty):
                section = nonempty[0]
                break
        data_trs = trs[header_index + 1:]
        anchor = None
        pattern = None
        for tr in data_trs:
            texts = [_tc_text(tc) for tc in _tr_tcs(tr)]
            if any(t in STOP_LABELS for t in texts):
                anchor = tr
                break
            st = header_cols.get('step')
            if st is not None and st < len(texts) and texts[st]:
                pattern = tr
        if anchor is None:
            anchor = trs[-1]
        if pattern is None:
            pattern = trs[header_index]
        detail_headers = [_tc_text(tc) for tc in _tr_tcs(trs[header_index])]
        # 每个列的自动角色（未识别到任一角色则为 None）
        idx2role = {}
        for role, idx in header_cols.items():
            if idx is not None:
                idx2role[idx] = role
        detail_roles = [idx2role.get(ci) for ci in range(len(detail_headers))]
        steps = {
            'header_tr': trs[header_index], 'pattern_tr': pattern, 'anchor_tr': anchor,
            'seq_col': header_cols.get('seq'), 'step_col': step_col,
            'exp_col': header_cols.get('expected'), 'act_col': header_cols.get('actual'),
            'crit_col': header_cols.get('criterion'),
            'section': section, 'detail_headers': detail_headers, 'detail_roles': detail_roles,
        }

    return {'fields': fields, 'steps': steps, 'warnings': warnings}


def get_template_fields(template_source):
    """识别模板的填写字段。返回 (fields, meta, warnings)。
    fields: [{'label','key'}]；
    meta: {'section': 分区标题或None, 'detail_headers': [明细表头列]}；
    template_source: None=内嵌，否则文件路径。"""
    doc, _ = _open_doc(template_source)
    try:
        if not doc.tables:
            return [], {}, ['文档中没有表格。']
        info = analyze_fields(doc.tables[0]._tbl)
        steps = info['steps']
        hdrs = steps.get('detail_headers') or []
        roles = steps.get('detail_roles') or []
        detail_cols = [{'header': hdrs[i], 'role': (roles[i] if i < len(roles) else None)}
                       for i in range(len(hdrs))]
        meta = {'section': steps.get('section'),
                'detail_headers': hdrs,
                'detail_cols': detail_cols}
        return info['fields'], meta, info['warnings']
    finally:
        close_doc(doc)


def validate_template(template_source):
    """校验模板可用性。返回 (ok, missing, warnings)。"""
    doc, _ = _open_doc(template_source)
    try:
        if not doc.tables:
            return False, [], ['文档中没有表格，无法使用。']
        info = analyze_fields(doc.tables[0]._tbl)
        labels = [f['label'] for f in info['fields']]
        missing = []
        if not any(_key_for_label(l) == 'name' for l in labels):
            missing.append('测试用例名称')
        if info['steps'].get('header_tr') is None:
            missing.append('测试过程(步骤/期望结果)')
        ok = '测试用例名称' not in missing
        return ok, missing, info['warnings']
    finally:
        close_doc(doc)


def _open_doc(template_source):
    if template_source is None or template_source == '':
        return Document(_load_template()), 'embedded'
    return Document(template_source), 'file'


def close_doc(doc):
    if hasattr(doc, 'close'):
        try:
            doc.close()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# 明细区填充
# ---------------------------------------------------------------------------
def _resolve_detail_cols(steps, config=None):
    """返回 {角色: 列索引}。先取自动识别的列，再用配置覆盖手工指定项。
    角色: seq/step/expected/actual/criterion/ignore。ignore 的列不写入。"""
    hdrs = steps.get('detail_headers') or []
    base = {
        'seq': steps.get('seq_col'), 'step': steps.get('step_col'),
        'expected': steps.get('exp_col'), 'actual': steps.get('act_col'),
        'criterion': steps.get('crit_col'),
    }
    colmap = {role: idx for role, idx in base.items() if idx is not None}
    cfg = (config or {}).get('detail_col_roles') or {}
    for header, role in cfg.items():
        if role in (None, '', 'ignore'):
            continue
        if header in hdrs:
            colmap[role] = hdrs.index(header)
    return colmap


def _fill_detail(tbl, info, data, config=None):
    steps = info['steps']
    header_tr = steps.get('header_tr')
    if header_tr is None:
        return
    anchor_tr = steps.get('anchor_tr')
    if anchor_tr is None:
        anchor_tr = trs_last(tbl)
    pattern_tr = steps.get('pattern_tr')
    if pattern_tr is None:
        pattern_tr = header_tr
    colmap = _resolve_detail_cols(steps, config)

    # 删除表头与锚点之间的旧数据行
    tr = header_tr.getnext()
    while tr is not None and tr is not anchor_tr:
        nxt = tr.getnext()
        tbl.remove(tr)
        tr = nxt

    for idx, st in enumerate(data, start=1):
        new_tr = copy.deepcopy(pattern_tr)
        anchor_tr.addprevious(new_tr)
        tcs = _tr_tcs(new_tr)
        for tc in tcs:
            _set_tc_text(tc, '')

        def cell(role):
            ci = colmap.get(role)
            return tcs[ci] if ci is not None and 0 <= ci < len(tcs) else None

        c = cell('seq')
        if c is None and 0 <= 0 < len(tcs):
            c = tcs[0]  # 无序号列时，序号写到第 0 列
        if c is not None:
            _set_tc_text(c, str(idx))
        for role, field in (('step', 'step'), ('expected', 'expected'),
                            ('actual', 'actual'), ('criterion', 'criterion')):
            c = cell(role)
            if c is not None:
                _set_tc_text(c, st.get(field, ''))


# ---------------------------------------------------------------------------
# 填充一个用例表格
# ---------------------------------------------------------------------------
def build_case_table(tbl, cfg, tconfig=None):
    """在模板表格(副本)中填入单个用例。cfg: {'fields':{label:text}, 'steps':[{...}]}。
    tconfig: 模板配置（含 detail_col_roles 覆盖）。info 需按当前表格重新分析。"""
    info = analyze_fields(tbl)
    fvalues = cfg.get('fields') or {}
    for f in info['fields']:
        _put_tc(f['tc'], fvalues.get(f['label'], ''))
    _fill_detail(tbl, info, cfg.get('steps') or [], tconfig)
    return tbl


# ---------------------------------------------------------------------------
# 标题样式与标题段落
# ---------------------------------------------------------------------------
def _find_style_id(doc, name):
    target = name.strip().lower()
    for s in doc.styles:
        try:
            if (s.name or '').strip().lower() == target:
                return s.element.get(qn('w:styleId'))
        except Exception:
            continue
    return None


def _ensure_unique_style_id(doc):
    used = {s.element.get(qn('w:styleId'))
            for s in doc.styles if s.element.get(qn('w:styleId'))}
    used.discard(None)
    n = 100
    while str(n) in used:
        n += 1
    return str(n)


def _ensure_headings(doc):
    id_h3 = _find_style_id(doc, 'heading 3') or '4'
    id_h4 = _find_style_id(doc, 'heading 4')
    if id_h4 is None:
        id_h4 = _ensure_unique_style_id(doc)
        base = _find_style_id(doc, 'heading 3') or '4'
        styles_el = doc.styles.element
        xml = (
            '<w:style %s w:type="paragraph" w:styleId="%s">'
            '<w:name w:val="heading 4"/><w:basedOn w:val="%s"/>'
            '<w:qFormat/><w:uiPriority w:val="0"/>'
            '<w:pPr><w:keepNext/><w:keepLines/><w:outlineLvl w:val="3"/></w:pPr>'
            '<w:rPr><w:szCs w:val="24"/></w:rPr>'
            '</w:style>' % (nsdecls('w'), id_h4, base)
        )
        styles_el.append(parse_xml(xml))
    return id_h3, id_h4


def _make_heading(text, style_id):
    xml = (
        '<w:p %s><w:pPr><w:pStyle w:val="%s"/>'
        '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>'
        '</w:pPr>'
        '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/>'
        '<w:szCs w:val="24"/></w:rPr><w:t>%s</w:t></w:r></w:p>'
    ) % (nsdecls('w'), style_id, text)
    return parse_xml(xml)


def _make_caption(num, name, style_id):
    """生成表标题（表格上方，题注样式）：`表N 测试用例名称`。
    使用 Word 题注样式 Caption，编号为 SEQ 域（表1/表2 自动排序），字体 宋体小五。"""
    rpr = ('<w:rPr><w:rFonts w:ascii="宋体" w:hAnsi="宋体" w:eastAsia="宋体"/>'
           '<w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>')
    xml = (
        '<w:p %s><w:pPr><w:pStyle w:val="%s"/><w:keepNext/><w:keepLines/></w:pPr>'
        '<w:r>%s<w:t xml:space="preserve">表 </w:t></w:r>'
        '<w:r>%s<w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r>%s<w:instrText xml:space="preserve"> SEQ 表 \\* ARABIC </w:instrText></w:r>'
        '<w:r>%s<w:fldChar w:fldCharType="separate"/></w:r>'
        '<w:r>%s<w:t xml:space="preserve">%d</w:t></w:r>'
        '<w:r>%s<w:fldChar w:fldCharType="end"/></w:r>'
        '<w:r>%s<w:t xml:space="preserve"> %s</w:t></w:r>'
        '</w:p>'
    ) % (nsdecls('w'), style_id, rpr, rpr, rpr, rpr, rpr, num, rpr, rpr, name)
    return parse_xml(xml)


def _ensure_caption_style(doc):
    """返回文档中"题注"（caption）样式的 styleId；不存在则自动新建。"""
    sid = _find_style_id(doc, 'caption')
    if sid:
        return sid
    base = _find_style_id(doc, 'normal') or 'a'
    new_id = _ensure_unique_style_id(doc)
    xml = (
        '<w:style %s w:type="paragraph" w:styleId="%s">'
        '<w:name w:val="Caption"/>'
        '<w:basedOn w:val="%s"/>'
        '<w:uiPriority w:val="35"/><w:semiHidden/><w:unhideWhenUsed/>'
        '<w:pPr><w:keepNext/><w:keepLines/></w:pPr>'
        '<w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>'
        '</w:style>' % (nsdecls('w'), new_id, base)
    )
    doc.styles.element.append(parse_xml(xml))
    return new_id


# ---------------------------------------------------------------------------
# 文档装配：多用例 -> 同一 Word
# ---------------------------------------------------------------------------
def generate_document(output_path, cases, template_source=None, h3_title='功能测试',
                      progress_cb=None):
    """
    cases: list[dict]，每项 {name, fields:{label:text,...}, steps:[{step,expected,actual}]}。
    fields 的 key 需与模板识别出的字段 label 一致。
    h3_title: 每组用例表上方的三级分组标题，默认"功能测试"。
    progress_cb: 可选回调 progress_cb(done, total)，每生成一个用例调用一次，用于进度条。"""
    doc, _ = _open_doc(template_source)
    _, tconfig = load_template_config(template_source)
    id_h3, id_h4 = _ensure_headings(doc)
    id_cap = _ensure_caption_style(doc)

    # 多张表时仅保留第一张为基准
    tbl_elements = doc.tables
    if len(tbl_elements) > 1:
        for tbl in list(tbl_elements)[1:]:
            t = tbl._tbl
            t.getparent().remove(t)
    first_tbl = doc.tables[0]._tbl

    body = doc.element.body
    paras = body.findall(qn('w:p'))
    if paras:
        body.remove(paras[0])  # 去掉模板自身的首段标题

    h3 = _make_heading(h3_title, id_h3)
    first_tbl.addprevious(h3)

    for i, case in enumerate(cases):
        fvalues = case.get('fields') or {}
        name = (case.get('name') or fvalues.get('测试用例名称') or '用例%d' % (i + 1)).strip()
        h4 = _make_heading(name, id_h4)       # 四级标题：用例名称
        cap = _make_caption(i + 1, name, id_cap)   # 表标题（题注样式，表1/表2 自动编号）
        if i == 0:
            h3.addnext(h4)
            h4.addnext(cap)
            build_case_table(first_tbl, cases[0], tconfig)
            prev = first_tbl
        else:
            prev.addnext(h4)
            h4.addnext(cap)
            new_tbl = copy.deepcopy(first_tbl)
            cap.addnext(new_tbl)
            build_case_table(new_tbl, cases[i], tconfig)
            prev = new_tbl
        if progress_cb is not None:
            try:
                progress_cb(i + 1, len(cases))
            except Exception:
                pass

    doc.save(output_path)
    close_doc(doc)


# ---------------------------------------------------------------------------
# 独立运行/调试入口
# ---------------------------------------------------------------------------
if __name__ == '__main__':
    op = r'D:\qtsubject\文件软件\_output_demo.docx'
    demo_cases = [
        {
            'name': '初始化功能_初始化Lookup Table_功能-001',
            'fields': {
                '测试用例名称': '初始化功能_初始化Lookup Table_功能-001',
                '测试用例标识': 'DMZC-CSH-LT_GN-001',
                '测试用例描述': '验证 Lookup Table 初始化功能。',
                '需求追踪': '3.2.1.4 初始化Lookup Table',
                '备注': '工具生成',
            },
            'steps': [
                {'step': '所有的子地址控制字为 0', 'expected': '所有子地址控制字保持为 0'},
                {'step': '使能接收子地址的接收双缓存功能', 'expected': '接收双缓存功能已使能'},
            ],
        }
    ]
    ok, missing, warns = validate_template(None)
    print('validate embedded:', ok, 'missing=', missing, 'warns=', warns)
    generate_document(op, demo_cases)
    print('OK ->', op, os.path.getsize(op), 'bytes')