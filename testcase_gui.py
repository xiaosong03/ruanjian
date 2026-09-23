# -*- coding: utf-8 -*-
"""测试用例批量生成工具 - GUI（字段驱动版）
- 选择 Word 模板后识别表格全部字段，动态显示为"公共字段"输入框（每张表填相同内容）
- 未填写的字段从 Excel 读取，按第一行表头匹配模板字段，一行 = 一个用例
- 界面优先：界面有值的字段用界面值，界面未填才用 Excel
"""
import json
import os
import threading
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

import openpyxl

from testcase_generator import (ALIAS_GROUPS, STEPS_COL_ALIASES,
                                _json_stream, _key_for_label, build_request_steps,
                                generate_document, get_template_fields,
                                inject_postman_request, load_template_config,
                                parse_postman_collection, save_template_config)

# 这些语义的字段在界面上用多行输入框，便于填写大段文本
MULTI_KEYS = {'desc', 'init', 'prereq', 'termination'}

# 明细列 / Excel 列的角色下拉选项（显示名 -> 角色key）
ROLE_CHOICES = [('步骤', 'step'), ('期望结果', 'expected'), ('评价准则', 'criterion'),
                ('实际结果', 'actual'), ('忽略', 'ignore')]
_DISPLAY2ROLE = dict(ROLE_CHOICES)
_ROLE2DISPLAY = {r: d for d, r in ROLE_CHOICES}


# ---------------------------------------------------------------------------
# Excel 读取与匹配
# ---------------------------------------------------------------------------
def _match_field(header, field_labels, key_to_label):
    """把 Excel 表头匹配到某个模板字段 label；返回 label 或 None。
    先精确匹配，失败后再做"包含匹配"（表头含别名，如"测试步骤说明"→"测试步骤"）。"""
    h = str(header).strip()
    if not h:
        return None
    for fl in field_labels:
        if h == fl:
            return fl
    for g in ALIAS_GROUPS:
        if h in g['labels']:
            return key_to_label.get(g['key'])
    # 包含匹配：按别名长度降序，取命中某别名的字段
    candidates = []
    for g in ALIAS_GROUPS:
        for alias in g['labels']:
            if len(alias) >= 2 and alias in h:
                candidates.append((len(alias), g['key']))
    if candidates:
        _, key = max(candidates, key=lambda x: x[0])
        return key_to_label.get(key)
    return None


def _match_step(header):
    """把 Excel 表头匹配到明细列角色(step/expected/seq/actual/criterion)，返回角色或 None。
    先精确匹配，失败后再做"包含匹配"。"""
    h = str(header).strip()
    if not h:
        return None
    for role, aliases in STEPS_COL_ALIASES.items():
        if h in aliases:
            return role
    candidates = []
    for role, aliases in STEPS_COL_ALIASES.items():
        for alias in aliases:
            if len(alias) >= 2 and alias in h:
                candidates.append((len(alias), role))
    if candidates:
        _, role = max(candidates, key=lambda x: x[0])
        return role
    return None


def _split_steps(step_text, exp_text):
    """按换行拆分步骤/期望并逐行对齐。"""
    step_lines = [s.strip() for s in step_text.split('\n') if s.strip()] if step_text else []
    exp_lines = [s.strip() for s in exp_text.split('\n') if s.strip()] if exp_text else []
    n = max(len(step_lines), len(exp_lines))
    steps = []
    for i in range(n):
        steps.append({
            'step': step_lines[i] if i < len(step_lines) else '',
            'expected': exp_lines[i] if i < len(exp_lines) else '',
        })
    return steps


def _is_json_text(text):
    """一段文本是否可解析为 JSON 流（含单个或多个拼接的 JSON 值，整段多行）。"""
    if not isinstance(text, str) or not text.strip():
        return False
    return _json_stream(text) is not None


def _align_steps(role_text):
    """把各明细角色文本按换行拆成逐段并对齐。role_text: {角色: 全文}。
    角色可取 step/expected/criterion/actual。返回 [{step,expected,criterion,...}]。

    对齐规则（保证期望/实际与步骤严格一一对应，不推断、不编造）：
      - 步骤(step)、期望(expected)、实际(actual)：严格按行号对齐，某角色行数不足则留空。
      - 评价准则(criterion)：若只写一段（不含换行），视为整条用例统一准则，套用到所有步骤行。
    """
    lines = {}
    for role, text in (role_text or {}).items():
        raw = str(text).replace('\r\n', '\n').replace('\r', '\n')
        # JSON 内容视为一个原子值（不按换行拆散），避免多行 JSON 被拆成多步错位
        if _is_json_text(raw):
            lines[role] = [raw.strip()]
        else:
            lines[role] = [s.strip() for s in raw.split('\n') if s.strip()]
    # 行数以"步骤"为准；若无步骤列则取各角色最大行数
    n = len(lines.get('step') or [])
    if n == 0:
        n = max((len(v) for v in lines.values()), default=0)
    single_criterion = bool(lines.get('criterion')) and len(lines['criterion']) == 1
    steps = []
    for i in range(n):
        d = {}
        for role, arr in lines.items():
            if role == 'criterion' and single_criterion:
                d[role] = arr[0]  # 统一准则：套用到所有步骤行
            else:
                d[role] = arr[i] if i < len(arr) else ''
        steps.append(d)
    return steps


def _case_warnings(cases):
    """生成前用例级一致性检查（不阻断生成）。
    返回 (dup_lines, align_lines)：
      - dup:    重复的测试用例标识；
      - align:  步骤数与期望结果/实际结果/评价准则行数不一致（错位风险）。"""
    dup_lines = []
    seen = {}
    for c in cases:
        cid_label = next((k for k in (c.get('fields') or {})
                          if _key_for_label(k) == 'case_id'), None)
        cid = ((c.get('fields') or {}).get(cid_label) if cid_label else '') \
            or (c.get('name') or '').strip()
        if not cid:
            continue
        if cid in seen:
            dup_lines.append('重复用例标识“{}”：见“{}”与“{}”'.format(cid, seen[cid], c['name']))
        else:
            seen[cid] = c['name']

    align_lines = []
    # 期望/实际/准则 都参与；仅在"真正配对错位"时告警：
    # 某一行一侧有内容、另一侧为空、而更后面那一侧又有内容（内容被空行错位顶跑）。
    # 仅末尾自然没有内容的行（如最后一个步骤无期望结果）不算问题。
    _ROLE_DISP = {'expected': '期望结果', 'actual': '实际结果', 'criterion': '评价准则'}
    for c in cases:
        steps = c.get('steps') or []
        for role, disp in _ROLE_DISP.items():
            has_step = [bool((s.get('step') or '').strip()) for s in steps]
            has_role = [bool((s.get(role) or '').strip()) for s in steps]
            if len(has_step) != len(has_role):
                has_role = has_role + [False] * (len(has_step) - len(has_role))
            misaligned = False
            for i in range(len(steps)):
                if has_step[i] and not has_role[i] and any(v for v in has_role[i + 1:]):
                    misaligned = True
                    break
                if has_role[i] and not has_step[i] and any(v for v in has_step[i + 1:]):
                    misaligned = True
                    break
            if misaligned:
                align_lines.append('用例“{}”：步骤与{}存在漏填/错位（同一行一侧有内容、另一侧为空，但后续行又有该列内容）'
                                   .format(c['name'], disp))
    return dup_lines, align_lines


def _write_warn_txt(out, warns):
    """把警告清单写到 <输出基名>.warn.txt；返回清单文本（无警告则 None）。"""
    if not warns:
        return None
    base, _ = os.path.splitext(out)
    path = base + '.warn.txt'
    try:
        with open(path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(warns))
    except Exception:
        path = None
    return '\n'.join(warns) + ('\n\n警告清单已存：' + path if path else '')


def _apply_placeholders(steps, fields):
    """把步骤文本里的 {字段名} 占位符替换为该用例该字段的实际值。
    - 只对含占位符的步骤生效，其余原样输出（零侵入）。
    - 未匹配到对应字段值的占位符保持原文，避免误伤。
    - 支持 步骤/期望/实际/准则 各角色。
    """
    repl = {f'{{{k}}}': str(v) for k, v in (fields or {}).items() if v}
    if not repl:
        return steps
    out = []
    for s in steps or []:
        d = dict(s)
        for role in ('step', 'expected', 'actual', 'criterion'):
            v = d.get(role)
            if not isinstance(v, str) or not v:
                continue
            for ph, val in repl.items():
                if ph in v:
                    v = v.replace(ph, val)
            d[role] = v
        out.append(d)
    return out


def load_cases_from_excel(path, field_labels, excel_map=None):
    """读取 Excel，返回 (cases, warnings)。
    cases: list[dict]，每项 {fields:{label:value}, steps:[{step,expected,criterion}]}。
    excel_map: 该模板配置的 excel_col_map（{表头: 目标}），用于手工补齐未自动匹配的列。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active

    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise RuntimeError('Excel 为空。')
    headers = [str(h).strip() if h is not None else '' for h in rows[0]]

    key_to_label = {}
    for fl in field_labels:
        k = _key_for_label(fl)
        if k and k not in key_to_label:
            key_to_label[k] = fl

    col_field = {}  # 列索引 -> 模板字段 label
    col_step = {}   # 列索引 -> 明细角色(step/expected/criterion)
    for i, h in enumerate(headers):
        if not h:
            continue
        m = _match_field(h, field_labels, key_to_label)
        if m:
            col_field[i] = m
            continue
        m = _match_step(h)
        if m:
            col_step[i] = m
    # 应用配置手工映射（可覆盖自动结果）
    for h, target in (excel_map or {}).items():
        if target in (None, '', 'ignore'):
            continue
        if h not in headers:
            continue
        i = headers.index(h)
        if target in ('step', 'expected', 'criterion', 'actual'):
            col_step[i] = target
            col_field.pop(i, None)
        elif target in field_labels:
            col_field[i] = target
            col_step.pop(i, None)
    if not col_field and not col_step:
        raise RuntimeError('Excel 表头未匹配到任何模板字段：' + '、'.join(headers))

    cases = []
    warns = []
    for r in rows[1:]:
        if r is None or all(v is None or str(v).strip() == '' for v in r):
            continue
        fields = {}
        for i, lbl in col_field.items():
            v = r[i]
            fields[lbl] = str(v).strip() if v is not None else ''
        role_text = {}
        for i, role in col_step.items():
            v = r[i]
            role_text[role] = str(v).strip() if v is not None else ''
        steps = _align_steps(role_text)
        cases.append({'fields': fields, 'steps': steps})
    if not cases:
        raise RuntimeError('Excel 中未读取到有效用例。')
    return cases, warns, headers


def _read_fill_excel(path):
    """读取回填 Excel：列名含"用例名"取名称，含"预期/期望"取预期结果，含"实际"取实际结果。
    返回 {测试用例名称: {'expected':.., 'actual':..}}（首个同名行生效）。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return {}
    headers = [str(h).strip() if h is not None else '' for h in rows[0]]

    def find_index(cols):
        for i, h in enumerate(headers):
            if h and any(c in h for c in cols):
                return i
        return None

    ki = find_index(['用例名称', '用例名'])
    ei = find_index(['预期', '期望', '希望'])
    ai = find_index(['实际'])
    if ki is None:
        return {}
    out = {}
    for r in rows[1:]:
        if r[ki] is None:
            continue
        name = str(r[ki]).strip()
        if not name:
            continue
        d = {}
        if ei is not None and r[ei] is not None and str(r[ei]).strip():
            d['expected'] = str(r[ei]).strip()
        if ai is not None and r[ai] is not None and str(r[ai]).strip():
            d['actual'] = str(r[ai]).strip()
        if d:
            out.setdefault(name, d)
    return out


# ---------------------------------------------------------------------------
# 主界面
# ---------------------------------------------------------------------------
class App:
    def __init__(self, root):
        self.root = root
        root.title('测试用例批量生成工具')
        root.geometry('820x720')
        root.minsize(760, 600)

        pad = {'padx': 8, 'pady': 4}
        frm = ttk.Frame(root, padding=8)
        frm.pack(fill='both', expand=True)
        frm.columnconfigure(1, weight=1)
        frm.rowconfigure(2, weight=1)  # 字段面板可伸缩

        # ---- 模板选择 ----
        ttk.Label(frm, text='Word 模板 (*.docx):').grid(row=0, column=0, sticky='w', **pad)
        self.tpl_var = tk.StringVar(value='（内置默认模板）')
        ttk.Entry(frm, textvariable=self.tpl_var, width=50).grid(row=0, column=1, sticky='we', **pad)
        ttk.Button(frm, text='浏览…', command=self._pick_tpl).grid(row=0, column=2, **pad)
        ttk.Button(frm, text='识别模板字段', command=lambda: self._identify(show=True)).grid(row=0, column=3, **pad)

        # ---- 分组标题：每张用例表上方的三级标题 ----
        ttk.Label(frm, text='分组标题:').grid(row=1, column=0, sticky='w', **pad)
        self.h3_var = tk.StringVar(value='功能测试')
        ttk.Entry(frm, textvariable=self.h3_var, width=50).grid(row=1, column=1, sticky='we', **pad)

        # ---- 公共字段面板（可滚动） ----
        self.gfrm = ttk.LabelFrame(frm,
                                    text=' 公共字段（界面填写，每个用例表填同样内容；未填写的字段从 Excel 读取） ',
                                    padding=4)
        self.gfrm.grid(row=2, column=0, columnspan=4, sticky='nsew', **pad)
        self._build_field_panel()

        # ---- Excel 数据源 ----
        ttk.Label(frm, text='数据文件 Excel (*.xlsx):').grid(row=3, column=0, sticky='w', **pad)
        self.xl_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self.xl_var, width=50).grid(row=3, column=1, sticky='we', **pad)
        ttk.Button(frm, text='浏览…', command=self._pick_xl).grid(row=3, column=2, **pad)
        ttk.Button(frm, text='灌入 Postman JSON…', command=self._postman_inject).grid(row=3, column=3, **pad)

        # ---- 公共步骤：界面写一次，自动插入到每个用例的步骤最前面（列随模板明细列动态生成） ----
        self.csf = ttk.LabelFrame(
            frm,
            text=' 公共步骤（界面写一次，自动插入到每个用例的步骤最前面；列会随模板明细列自动变化） ',
            padding=4)
        self.csf.grid(row=4, column=0, columnspan=4, sticky='we', **pad)

        # ---- 输出文件 ----
        ttk.Label(frm, text='输出 Word (*.docx):').grid(row=5, column=0, sticky='w', **pad)
        self.out_var = tk.StringVar(value=self._default_output())
        ttk.Entry(frm, textvariable=self.out_var, width=50).grid(row=5, column=1, sticky='we', **pad)
        ttk.Button(frm, text='浏览…', command=self._pick_out).grid(row=5, column=2, **pad)

        # ---- 预览信息 ----
        self.info = tk.StringVar(value='尚未识别模板。')
        ttk.Label(frm, textvariable=self.info, foreground='#2069c5').grid(
            row=6, column=0, columnspan=4, sticky='w', **pad)

        # ---- 执行按钮 + 进度条 ----
        btnfrm = ttk.Frame(frm)
        btnfrm.grid(row=7, column=0, columnspan=4, **pad)
        self.gen_btn = ttk.Button(btnfrm, text='生成 Word 文档', command=self._generate)
        self.gen_btn.pack(side='left', padx=6)
        ttk.Button(btnfrm, text='批量生成…', command=self._open_batch).pack(side='left', padx=6)
        ttk.Button(btnfrm, text='退出', command=root.destroy).pack(side='left')
        self.progress = ttk.Progressbar(frm, mode='determinate', maximum=100)
        self.progress.grid(row=8, column=0, columnspan=4, sticky='we', padx=8, pady=(0, 2))

        self.template_fields = []  # [{'label','key'}]
        self._meta = {}
        self._entries = {}         # label -> StringVar
        self._texts = {}           # label -> tk.Text
        self._excel_rows = None
        self._last_excel_headers = []
        self._detail_role_vars = {}  # 明细列表头 -> StringVar（自动判不准需用户指定）
        self._excel_role_vars = {}   # Excel表头 -> StringVar（未自动匹配需用户指定）
        self._config_key = None
        self._identify(show=False)  # 默认用内置模板识别字段

    # ---------- 可滚动字段面板 ----------
    def _build_field_panel(self):
        self.canvas = tk.Canvas(self.gfrm, highlightthickness=0)
        vsb = ttk.Scrollbar(self.gfrm, orient='vertical', command=self.canvas.yview)
        self.field_frame = ttk.Frame(self.canvas)
        self.fwin = self.canvas.create_window((0, 0), window=self.field_frame, anchor='nw')
        self.canvas.configure(yscrollcommand=vsb.set)
        self.canvas.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')
        self.field_frame.bind('<Configure>',
                              lambda e: self.canvas.configure(scrollregion=self.canvas.bbox('all')))
        self.canvas.bind('<Configure>',
                         lambda e: self.canvas.itemconfigure(self.fwin, width=e.width))
        # 鼠标滚轮控制滚动
        self.canvas.bind_all('<MouseWheel>', self._on_wheel)

    def _on_wheel(self, event):
        if self.canvas.winfo_height() > 0:
            self.canvas.yview_scroll(int(-event.delta / 120), 'units')

    # ---------- 模板识别 ----------
    def _tpl_source(self):
        tpl = self.tpl_var.get().strip()
        return None if tpl in ('', '（内置默认模板）') else tpl

    def _default_output(self):
        """默认输出文件名带时间戳，避免重跑覆盖上一版。"""
        return '_output_{}.docx'.format(datetime.now().strftime('%Y%m%d_%H%M%S'))

    # 公共步骤面板：显示顺序 + 角色显示名
    _COMMON_ROLE_DISPLAY = [('step', '步骤'), ('expected', '期望结果'),
                            ('actual', '实际结果'), ('criterion', '评价准则')]

    def _build_common_steps(self):
        """按当前模板实际检测到的明细列重建公共步骤编辑栏。
        只显示模板里真实存在的明细角色，避免写入无处可去的列。"""
        for w in self.csf.winfo_children():
            w.destroy()
        self._common_txts = {}
        roles = set()
        for dc in self._meta.get('detail_cols', []) or []:
            r = dc.get('role')
            if r:
                roles.add(r)
        cols = [(r, d) for r, d in self._COMMON_ROLE_DISPLAY if r in roles]
        if not cols:  # 连明细角色都没识别到时，退化为固定"步骤/期望"
            cols = [('step', '步骤'), ('expected', '期望结果')]
        for ci, (role, disp) in enumerate(cols):
            sub = ttk.Frame(self.csf)
            sub.grid(row=0, column=ci, sticky='nsew', padx=4, pady=2)
            ttk.Label(sub, text=disp).pack(anchor='w')
            txt = tk.Text(sub, height=3, width=28, wrap='word')
            txt.pack(fill='both', expand=True)
            self._common_txts[role] = txt
            self.csf.columnconfigure(ci, weight=1)

    def _load_common_steps(self):
        """恢复当前模板自己记忆的公共步骤（按角色写回对应编辑框）。"""
        _, cfg = load_template_config(self._tpl_source())
        ps = cfg.get('public_steps') or {}
        for role, txt in self._common_txts.items():
            val = ps.get(role) or ''
            txt.insert('1.0', val)

    def _save_common_steps(self):
        """把当前公共步骤按角色存到当前模板的配置（每个模板各自记忆，不互相串）。"""
        try:
            save_template_config(self._tpl_source(), {
                'public_steps': {role: txt.get('1.0', 'end').strip()
                                 for role, txt in self._common_txts.items()}
            })
        except Exception:
            pass

    def _common_steps_from_gui(self):
        """把界面公共步骤解析为步骤行（各角色逐行与步骤对应；写一段的覆盖全部行）。
        不同模板列不同，这里只取当前界面存在的角色。"""
        return _align_steps({role: txt.get('1.0', 'end')
                             for role, txt in self._common_txts.items()})

    def _pick_tpl(self):
        p = filedialog.askopenfilename(title='选择 Word 模板',
                                       filetypes=[('Word 文档', '*.docx'), ('所有文件', '*.*')])
        if p:
            self.tpl_var.set(p)
            self._identify(show=True)

    def _identify(self, show):
        src = self._tpl_source()
        try:
            fields, meta, warns = get_template_fields(src)
        except Exception as e:
            self.info.set(f'模板读取失败：{e}')
            return
        self.template_fields = fields
        self._meta = meta
        self._config_key, _ = load_template_config(src)
        self._detail_role_vars.clear()
        self._excel_role_vars.clear()
        self._refresh_fields()
        self._build_common_steps()   # 明细列随模板变化
        self._load_common_steps()    # 恢复该模板自己的公共步骤

        label_list = '、'.join(f['label'] for f in fields) or '（无字段）'
        extra = f'（提示：{"；".join(warns[:3])}）' if warns else ''
        if show:
            self.info.set(f'已识别字段：{label_list}{extra}')
        else:
            self.info.set('已识别模板字段。' + extra)

    def _refresh_fields(self):
        # 重建前保存当前界面已输入的值，重建后回填，避免重建清空用户数据
        saved = {}
        for label, v in self._entries.items():
            saved[label] = v.get()
        for label, t in self._texts.items():
            saved[label] = t.get('1.0', 'end').strip()
        for w in self.field_frame.winfo_children():
            w.destroy()
        self._entries.clear()
        self._texts.clear()
        # 表头
        ttk.Label(self.field_frame, text='字段', font=('', 9, 'bold')).grid(
            row=0, column=0, sticky='w', padx=6, pady=(2, 2))
        ttk.Label(self.field_frame, text='填写内容', font=('', 9, 'bold')).grid(
            row=0, column=1, sticky='w', padx=6, pady=(2, 2))
        self.field_frame.columnconfigure(1, weight=1)

        # 当前已导入 Excel 中已提供数据的字段（界面可不填，生成时自动使用 Excel 值）
        excel_have = set()
        for er in (self._excel_rows or []):
            for lbl, val in (er.get('fields') or {}).items():
                if val:
                    excel_have.add(lbl)

        for ri, f in enumerate(self.template_fields, start=1):
            label = f['label']
            key = f.get('key')
            ttk.Label(self.field_frame, text=f'{label}:').grid(
                row=ri, column=0, sticky='nw', padx=6, pady=2)
            if label in excel_have:
                ttk.Label(self.field_frame, text='←Excel有，可不填',
                          foreground='#2e7d32').grid(
                    row=ri, column=2, sticky='nw', padx=(2, 6), pady=2)
            prev = saved.get(label, '')
            if key in MULTI_KEYS:
                txt = tk.Text(self.field_frame, width=66, height=2, wrap='word')
                txt.grid(row=ri, column=1, sticky='we', padx=6, pady=2)
                if prev:
                    txt.insert('1.0', prev)
                self._texts[label] = txt
            else:
                v = tk.StringVar(value=prev)
                ent = ttk.Entry(self.field_frame, textvariable=v, width=70)
                ent.grid(row=ri, column=1, sticky='we', padx=6, pady=2)
                self._entries[label] = v

        # ---- 明细区字段（只读展示，由 Excel 自动填写） ----
        section = (self._meta or {}).get('section')
        detail_headers = (self._meta or {}).get('detail_headers') or []
        if section or detail_headers:
            r0 = len(self.template_fields) + 1
            ttk.Separator(self.field_frame, orient='horizontal').grid(
                row=r0, column=0, columnspan=2, sticky='we', padx=6, pady=4)
            ro = r0 + 1
            if section:
                ttk.Label(self.field_frame, text='■ ' + section,
                          foreground='#888888').grid(row=ro, column=0, columnspan=2,
                                                     sticky='w', padx=6, pady=2)
                ro += 1
            if detail_headers:
                text = ('明细列（序号自动编号；由 Excel 对应列按行自动填写）：' +
                        (' | '.join(detail_headers) if detail_headers else ''))
                ttk.Label(self.field_frame, text=text, foreground='#888888', wraplength=600,
                          justify='left').grid(row=ro, column=0, columnspan=2,
                                               sticky='w', padx=6, pady=2)
            ro += 1
        else:
            ro = len(self.template_fields) + 1
        # 每个明细列都可下拉纠正角色（自动记忆到该模板配置）；自动识别结果作为默认值
        dc = (self._meta or {}).get('detail_cols') or []
        if dc:
            ro += 1
            ttk.Label(self.field_frame, text='明细列角色（自动识别，可下拉纠正后自动记忆）：',
                      foreground='#aa4a00', wraplength=600, justify='left').grid(
                row=ro, column=0, columnspan=2, sticky='w', padx=6, pady=2)
            ro += 1
            _, tcfg = load_template_config(self._tpl_source())
            for c in dc:
                header = c['header']
                auto_role = c.get('role')      # 自动识别角色
                cur = tcfg['detail_col_roles'].get(header) or auto_role or 'ignore'
                var = tk.StringVar(value=_ROLE2DISPLAY.get(cur, '忽略'))
                self._detail_role_vars[header] = var
                auto_txt = ('（自动:' + _ROLE2DISPLAY.get(auto_role, '未识别') + '）'
                            if auto_role else '（未识别）')
                ttk.Label(self.field_frame, text=f'明细列「{header}」{auto_txt}:').grid(
                    row=ro, column=0, sticky='w', padx=(12, 6), pady=1)
                cb = ttk.Combobox(self.field_frame, textvariable=var,
                                  values=[d for d, _ in ROLE_CHOICES], state='readonly', width=14)
                cb.grid(row=ro, column=1, sticky='w', padx=6, pady=1)
                cb.bind('<<ComboboxSelected>>',
                        lambda _e, h=header: self._on_detail_role(h))
                ro += 1
        # Excel 未自动匹配的列：用户下拉指定（自动记忆）
        ro += self._render_excel_map(ro)
        self.canvas.configure(scrollregion=self.canvas.bbox('all'))

    def _render_excel_map(self, ro0):
        """渲染 Excel 未自动匹配列的映射下拉框。返回占用的行数。"""
        headers = self._last_excel_headers
        if not headers or not self.template_fields:
            return 0
        labels = [f['label'] for f in self.template_fields]
        key_to_label = {}
        for fl in labels:
            k = _key_for_label(fl)
            if k and k not in key_to_label:
                key_to_label[k] = fl
        undet = [h for h in headers
                 if h and not _match_field(h, labels, key_to_label) and not _match_step(h)]
        if not undet:
            return 0
        _, tcfg = load_template_config(self._tpl_source())
        ro = ro0 + 1
        ttk.Label(self.field_frame, text='以下 Excel 列未自动匹配，请指定其归属（自动记忆）：',
                  foreground='#aa4a00', wraplength=600, justify='left').grid(
            row=ro, column=0, columnspan=2, sticky='w', padx=6, pady=2)
        ro += 1
        choices = [d for d, _ in ROLE_CHOICES] + labels
        for h in undet:
            cur = tcfg['excel_col_map'].get(h)
            disp = _ROLE2DISPLAY.get(cur) if cur in _ROLE2DISPLAY else cur
            var = tk.StringVar(value=disp or '忽略')
            self._excel_role_vars[h] = var
            ttk.Label(self.field_frame, text=f'Excel列「{h}」:').grid(
                row=ro, column=0, sticky='w', padx=(12, 6), pady=1)
            cb = ttk.Combobox(self.field_frame, textvariable=var, values=choices,
                              state='readonly', width=20)
            cb.grid(row=ro, column=1, sticky='w', padx=6, pady=1)
            cb.bind('<<ComboboxSelected>>', lambda _e, h=h: self._on_excel_role(h))
            ro += 1
        return ro - ro0

    def _on_detail_role(self, header):
        role = _DISPLAY2ROLE.get(self._detail_role_vars[header].get(), 'ignore')
        save_template_config(self._tpl_source(), {'detail_col_roles': {header: role}})

    def _on_excel_role(self, header):
        val = self._excel_role_vars[header].get()
        target = _DISPLAY2ROLE.get(val, val)  # 角色名 -> 角色key；否则为模板字段label
        if target == '忽略':
            target = 'ignore'
        save_template_config(self._tpl_source(), {'excel_col_map': {header: target}})
        self._load_preview(rebuild=False)

    # ---------- 文件选择 ----------
    def _pick_xl(self):
        p = filedialog.askopenfilename(title='选择 Excel 数据文件',
                                       filetypes=[('Excel 工作簿', '*.xlsx'), ('所有文件', '*.*')])
        if p:
            self.xl_var.set(p)
            self._load_preview()

    def _postman_inject(self):
        """灌入 Postman JSON：把单个接口信息写入已生成中间 Word 的对应用例步骤格。"""
        PostmanInjectDialog(self.root, self)

    def _pick_out(self):
        p = filedialog.asksaveasfilename(title='保存输出文档', defaultextension='.docx',
                                         filetypes=[('Word 文档', '*.docx')],
                                         initialfile=self._default_output())
        if p:
            self.out_var.set(p)

    # ---------- 逻辑 ----------
    def _collect_interface(self):
        """界面公共字段：{label: 值}，仅多有值的非空字段。"""
        cfg = {}
        for f in self.template_fields:
            label = f['label']
            if label in self._texts:
                v = self._texts[label].get('1.0', 'end').strip()
            else:
                v = self._entries[label].get().strip()
            cfg[label] = v
        return cfg

    def _load_preview(self, rebuild=True):
        xl = self.xl_var.get()
        if not xl:
            return
        try:
            _, tcfg = load_template_config(self._tpl_source())
            labels = [f['label'] for f in self.template_fields]
            rows, warns, headers = load_cases_from_excel(xl, labels, tcfg.get('excel_col_map'))
            self._excel_rows = rows
            self._last_excel_headers = headers
            total_steps = sum(len(r['steps']) for r in rows)
            extra = f'（提示：{"; ".join(warns[:3])}）' if warns else ''
            self.info.set(f'已导入 {len(rows)} 条用例，共 {total_steps} 个测试步骤。{extra}')
            if rebuild:
                self._refresh_fields()
        except Exception as e:
            self._excel_rows = None
            self.info.set(f'Excel 导入失败：{e}')

    # ---------- 空字段预检 ----------
    # 提示语无需改名；关键字段（名称/标识）为空时输出易混淆，其它为空仅提示。
    def _precheck_cases(self, cases):
        """统计各用例中"将为空"的字段（界面未填、Excel 也没有）。
        返回 (critical, others)：{字段label: 出现个数}。"""
        critical = {}
        others = {}
        for c in (cases or []):
            for lbl, v in (c.get('fields') or {}).items():
                if v:
                    continue
                if _key_for_label(lbl) in ('name', 'case_id'):
                    critical[lbl] = critical.get(lbl, 0) + 1
                else:
                    others[lbl] = others.get(lbl, 0) + 1
        return critical, others

    def _precheck_confirm(self, cases, title_suffix=''):
        """生成前空字段预检：有字段将为空时列清单询问是否继续。返回 True=继续。"""
        critical, others = self._precheck_cases(cases)
        if not critical and not others:
            return True
        parts = []
        if critical:
            detail = '；'.join('{}：{}个'.format(l, n) for l, n in critical.items())
            parts.append('【关键必填缺失，输出后易混淆】' + detail)
        if others:
            detail = '；'.join('{}：{}个'.format(l, n) for l, n in others.items())
            parts.append('【其它字段留空】' + detail)
        msg = '\n\n'.join(parts) + '\n\n仍继续生成吗？（选“否”可先回填再生成）'
        title = '空字段预检' + title_suffix
        return bool(messagebox.askyesno(title, msg))

    def _generate(self):
        if not self.template_fields:
            messagebox.showwarning('提示', '请先识别模板字段。')
            return
        if self.xl_var.get():
            self._load_preview()

        iface = self._collect_interface()
        common_steps = self._common_steps_from_gui()   # 界面公共步骤，插入每个用例最前
        self._save_common_steps()                      # 记住本次填写的公共步骤
        if common_steps:
            self.info.set('提示：公共步骤有 {} 条，将自动插入到每个用例的最前面。'.format(len(common_steps)))
        name_label = next((f['label'] for f in self.template_fields
                           if _key_for_label(f['label']) == 'name'), None)

        rows = self._excel_rows or [None]
        cases = []
        for idx, er in enumerate(rows, start=1):
            fields = {}
            for f in self.template_fields:
                lbl = f['label']
                ival = iface.get(lbl, '')
                if ival:
                    fields[lbl] = ival  # 界面优先
                else:
                    ev = ((er or {}).get('fields') or {}).get(lbl, '') if er else ''
                    fields[lbl] = ev
            # 私有步骤来自 Excel，公共步骤插到最前，两者按行一一对应
            steps = common_steps + (er['steps'] if er else [])
            # 占位符替换：步骤/期望文本里的 {字段名} 换成该用例该字段的实际值
            steps = _apply_placeholders(steps, fields)
            name = (fields.get(name_label) if name_label else '') or ('用例%d' % idx)
            cases.append({'name': name, 'fields': fields, 'steps': steps})

        out = self.out_var.get()
        if not out:
            messagebox.showwarning('提示', '请指定输出文件路径。')
            return

        # ---- 用例级一致性警告（重复标识/步骤期望错位），只提示不阻断 ----
        dup_w, al_w = _case_warnings(cases)
        warn_text = _write_warn_txt(out, dup_w + al_w)
        if warn_text:
            messagebox.showwarning('生成警告', warn_text + '\n\n仍将继续生成。')

        # ---- 空字段预检：界面未填、Excel 也没有的字段提前暴露，避免生成后返工 ----
        if not self._precheck_confirm(cases):
            return

        src = self._tpl_source()
        h3_title = self.h3_var.get().strip() or '功能测试'
        self.gen_btn.config(state='disabled')
        self.progress.config(value=0)
        self.info.set('正在生成…')
        total = len(cases)

        def on_progress(done, t):
            pct = (done / t * 100) if t else 0
            self.root.after(0, lambda: self.progress.config(value=pct))

        def work():
            log_path = None
            try:
                generate_document(out, cases, template_source=src,
                                  h3_title=h3_title, progress_cb=on_progress)
                log_path = self._write_log(out, cases)
                self.root.after(0, self._done, True,
                                '生成成功：{}（{} 个用例）'.format(out, len(cases)), log_path)
            except Exception as e:
                self.root.after(0, self._done, False, '生成失败：{}'.format(e), log_path)

        threading.Thread(target=work, daemon=True).start()

    def _write_log(self, out, cases):
        """写一份生成日志（txt），记录时间、数据源、匹配情况、用例数，便于排查。返回日志路径。"""
        try:
            base, _ = os.path.splitext(out)
            log_path = base + '.log.txt'
            lines = []
            lines.append('测试用例批量生成工具 - 生成日志')
            lines.append('生成时间：' + datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            lines.append('Word 模板：' + (self.tpl_var.get() or '（内置模板）'))
            lines.append('Excel 数据：' + (self.xl_var.get() or '（未使用，仅界面字段）'))
            lines.append('分组标题：' + (self.h3_var.get().strip() or '功能测试'))
            lines.append('输出文件：' + out)
            lines.append('用例数量：{}'.format(len(cases)))
            total_steps = sum(len(c['steps']) for c in cases)
            lines.append('测试步骤总数：{}'.format(total_steps))
            try:
                cs = self._common_steps_from_gui()
                if cs:
                    lines.append('公共步骤：{} 条（已插入每个用例最前）'.format(len(cs)))
                else:
                    lines.append('公共步骤：无')
            except Exception:
                pass
            lines.append('')
            lines.append('---- Excel 表头匹配情况 ----')
            headers = self._last_excel_headers
            if not headers:
                lines.append('（无 Excel 数据）')
            else:
                labels = [f['label'] for f in self.template_fields]
                key_to_label = {}
                for fl in labels:
                    k = _key_for_label(fl)
                    if k and k not in key_to_label:
                        key_to_label[k] = fl
                for h in headers:
                    if not h:
                        continue
                    if _match_field(h, labels, key_to_label):
                        lines.append('[字段] {} -> {}'.format(h, _match_field(h, labels, key_to_label)))
                    elif _match_step(h):
                        lines.append('[明细] {} -> {}'.format(h, _match_step(h)))
                    else:
                        lines.append('[未匹配] {}！！'.format(h))
            lines.append('')
            lines.append('---- 各用例简要 ----')
            for i, c in enumerate(cases, start=1):
                lines.append('用例{0}: {1}（步骤{2}条）'.format(
                    i, c['name'], len(c['steps'])))
            with open(log_path, 'w', encoding='utf-8') as f:
                f.write('\n'.join(lines))
            return log_path
        except Exception:
            return None

    def _done(self, ok, msg, log_path=None):
        self.gen_btn.config(state='normal')
        self.progress.config(value=100 if ok else 0)
        self.info.set(msg)
        if ok:
            text = msg
            if log_path:
                text += '\n\n日志已保存：{}'.format(log_path)
            messagebox.showinfo('完成', text)
        else:
            messagebox.showerror('错误', msg)

    # ---------- 批量多模板（最简版） ----------
    def _open_batch(self):
        BatchWindow(self)

    def _gen_task_cases(self, src, xl, iface, common_steps):
        """针对单个批量任务构造用例。模板字段来自该任务自身模板；
        公共字段/公共步骤沿用主窗口（第一次任务）填的作为默认。"""
        fields_list, _meta, _warns = get_template_fields(src)
        if not fields_list:
            raise RuntimeError('模板未识别到字段。')
        _, tcfg = load_template_config(src)
        labels = [f['label'] for f in fields_list]
        rows, _w, _h = load_cases_from_excel(xl, labels, tcfg.get('excel_col_map'))
        name_label = next((f['label'] for f in fields_list
                           if _key_for_label(f['label']) == 'name'), None)
        cases = []
        for idx, er in enumerate(rows, start=1):
            fields = {}
            for f in fields_list:
                lbl = f['label']
                ival = iface.get(lbl, '')
                fields[lbl] = ival if ival else ((er or {}).get('fields') or {}).get(lbl, '')
            steps = common_steps + (er['steps'] if er else [])
            steps = _apply_placeholders(steps, fields)
            name = (fields.get(name_label) if name_label else '') or ('用例%d' % idx)
            cases.append({'name': name, 'fields': fields, 'steps': steps})
        return cases


class BatchWindow:
    """批量生成（最简版）：可添加去重多个"模板+Excel+输出名"任务，
    公共字段/公共步骤沿用主窗口（第一次任务）填的作为默认。"""

    def __init__(self, app):
        self.app = app
        self.root = tk.Toplevel(app.root)
        self.root.title('批量生成（多模板）')
        self.root.geometry('780x420')
        self.root.transient(app.root)

        pad = {'padx': 6, 'pady': 4}
        frm = ttk.Frame(self.root, padding=8)
        frm.pack(fill='both', expand=True)

        cols = ('tpl', 'xl', 'out')
        self.tree = ttk.Treeview(frm, columns=cols, show='headings', height=10)
        for c, t, w in (('tpl', 'Word 模板', 250),
                        ('xl', 'Excel', 230),
                        ('out', '输出 .docx', 240)):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor='w')
        self.tree.pack(fill='both', expand=True, **pad)
        self.tasks = []  # [{'src','xl','out',
                         #    'fields':[{'label','key'}],  # 该任务模板自己的字段
                         #    'meta': detail cols ...
                         #    'override': {'fields':{label:val}, 'public_steps':{role:raw}, 'configured':bool}}]

        btns = ttk.Frame(frm)
        btns.pack(fill='x', **pad)
        ttk.Button(btns, text='添加任务', command=self._add).pack(side='left', padx=4)
        ttk.Button(btns, text='补填…', command=self._edit).pack(side='left', padx=4)
        ttk.Button(btns, text='删除所选', command=self._remove).pack(side='left', padx=4)
        ttk.Button(btns, text='清除', command=self._clear).pack(side='left', padx=4)
        self.run_btn = ttk.Button(btns, text='生成全部', command=self._generate_all)
        self.run_btn.pack(side='right', padx=4)

        self.status = tk.StringVar(value='已添加 0 个任务')
        ttk.Label(frm, textvariable=self.status, foreground='#2069c5').pack(anchor='w', **pad)
        self.progress = ttk.Progressbar(frm, mode='determinate', maximum=100)
        self.progress.pack(fill='x', padx=6, pady=(0, 2))

    def _add(self):
        src = filedialog.askopenfilename(title='选择 Word 模板',
                                         filetypes=[('Word 文档', '*.docx'), ('所有文件', '*.*')])
        if not src:
            return
        xl = filedialog.askopenfilename(title='选择 Excel 数据文件',
                                        filetypes=[('Excel 工作簿', '*.xlsx'), ('所有文件', '*.*')])
        if not xl:
            return
        default_out = os.path.join(os.path.dirname(xl) or '.', self.app._default_output())
        out = filedialog.asksaveasfilename(title='保存输出文档', defaultextension='.docx',
                                           filetypes=[('Word 文档', '*.docx')],
                                           initialdir=os.path.dirname(default_out),
                                           initialfile=os.path.basename(default_out))
        if not out:
            return
        # 去重：相同 模板+Excel+输出名 不重复添加
        for t in self.tasks:
            if (os.path.normpath(t['src']), os.path.normpath(t['xl']), os.path.normpath(t['out'])) == \
               (os.path.normpath(src), os.path.normpath(xl), os.path.normpath(out)):
                messagebox.showwarning('提示', '该任务已存在（模板/Excel/输出名均相同）。')
                return
        try:
            fields_list, meta, _w = get_template_fields(src)
        except Exception as e:
            messagebox.showerror('模板读取失败', str(e))
            return
        self.tasks.append({'src': src, 'xl': xl, 'out': out,
                           'fields': fields_list, 'meta': meta,
                           'override': {'fields': {}, 'public_steps': {},
                                        'configured': False}})
        self.tree.insert('', 'end',
                         values=(os.path.basename(src), os.path.basename(xl), out))
        self.status.set('已添加 {} 个任务'.format(len(self.tasks)))

    def _edit(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning('提示', '请先选择要补填的任务。')
            return
        idx = self.tree.index(sel[0])
        TaskEditDialog(self, self.tasks[idx])

    def _remove(self):
        for iid in self.tree.selection():
            idx = self.tree.index(iid)
            del self.tasks[idx]
            self.tree.delete(iid)
        self.status.set('已添加 {} 个任务'.format(len(self.tasks)))

    def _clear(self):
        self.tasks.clear()
        for i in self.tree.get_children():
            self.tree.delete(i)
        self.status.set('已添加 0 个任务')

    def _generate_all(self):
        if not self.tasks:
            messagebox.showwarning('提示', '请先添加任务。')
            return
        main_iface = self.app._collect_interface()            # 主窗口公共字段
        main_steps = self.app._common_steps_from_gui()         # 公共步骤默认沿用
        # 主线程先构建各任务用例，做空字段预检（弹窗确认后再启动生成线程）
        plan = []        # [(task, cases)]
        prep_fail = []   # (task, errmsg)
        warn_n = 0       # 有警告的任务数
        for t in self.tasks:
            try:
                task_iface = dict(main_iface)
                task_iface.update(t['override']['fields'])
                cs = _align_steps(t['override']['public_steps']) if t['override'].get('configured') else main_steps
                cs_this = self.app._gen_task_cases(t['src'], t['xl'], task_iface, cs)
                plan.append((t, cs_this))
                # 用例级警告：该任务重复标识/步骤期望错位，写 warn.txt，不阻断
                dup_w, al_w = _case_warnings(cs_this)
                if _write_warn_txt(t['out'], dup_w + al_w):
                    warn_n += 1
            except Exception as e:
                prep_fail.append((t, str(e)))
        if not plan:
            if prep_fail:
                self._finish([(os.path.basename(t['src']), False, m) for t, m in prep_fail])
            else:
                messagebox.showwarning('提示', '没有可生成的任务。')
            return
        all_cases = [c for _t, cs in plan for c in cs]
        if not self.app._precheck_confirm(all_cases, '（批量）'):
            return

        self.run_btn.config(state='disabled')
        self.progress.config(value=0)
        total = len(plan)
        results = prep_fail and [(os.path.basename(t['src']), False, m) for t, m in prep_fail] or []
        results = list(results)

        def work():
            for i, (t, cases) in enumerate(plan, start=1):
                try:
                    generate_document(t['out'], cases, template_source=t['src'],
                                      h3_title=self.app.h3_var.get().strip() or '功能测试')
                    results.append((os.path.basename(t['src']), True, '成功，{} 个用例'.format(len(cases))))
                except Exception as e:
                    results.append((os.path.basename(t['src']), False, str(e)))
                self.root.after(0, lambda d=i: self.progress.config(value=d / total * 100))
            self.root.after(0, lambda: self._finish(results, warn_n))

        threading.Thread(target=work, daemon=True).start()

    def _finish(self, results, warn_n=0):
        self.run_btn.config(state='normal')
        self.progress.config(value=100)
        ok_n = sum(1 for _, ok, _ in results if ok)
        fail = [(n, m) for n, _ok, m in results if not _ok]
        self.status.set('成功 {} / {}，失败 {}'.format(ok_n, len(results), len(fail)))
        warn_part = '；{} 个任务有警告（已存 .warn.txt）'.format(warn_n) if warn_n else ''
        err_path = None
        if fail:
            # 失败明细同时写入一份错误 txt，便于排查
            try:
                base_dir = (os.path.dirname(os.path.abspath(self.tasks[0]['out']))
                            if getattr(self, 'tasks', None) else os.getcwd())
                err_path = os.path.join(
                    base_dir, '批量生成错误_{}.txt'.format(datetime.now().strftime('%Y%m%d_%H%M%S')))
                with open(err_path, 'w', encoding='utf-8') as f:
                    f.write('批量生成 - 失败明细\n生成时间：{}\n\n'.format(
                        datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
                    for n, m in fail:
                        f.write('【{}】\n{}\n\n'.format(n, m))
            except Exception:
                err_path = None
            msg = '\n'.join('{}：{}'.format(n, m) for n, m in fail)
            if err_path:
                msg += '\n\n失败清单已存：{}'.format(err_path)
            messagebox.showwarning('批量生成完成',
                                   '成功 {} 份，失败 {} 份{}。\n\n失败明细：\n{}'.format(
                                       ok_n, len(fail), warn_part, msg))
        else:
            messagebox.showinfo('批量生成完成',
                                '全部成功：{} 份文档已生成{}。'.format(ok_n, warn_part))


class PostmanInjectDialog:
    """灌入 Postman JSON（单个接口）：选中间 Word + Collection json，按接口名写入对应用例步骤格。"""

    def __init__(self, root, app):
        self.app = app
        self.doc_path = None
        self.reqs = []
        self.fill_map = {}
        self._build(root)

    def _build(self, root):
        win = tk.Toplevel(root)
        win.title('灌入 Postman JSON · 单接口')
        win.geometry('560x300')
        win.resizable(False, False)
        win.transient(root)
        win.grab_set()
        pad = {'padx': 8, 'pady': 4}
        frm = ttk.Frame(win, padding=8)
        frm.pack(fill='both', expand=True)

        ttk.Label(frm, text='中间 Word（已生成、步骤列留空）:').grid(row=0, column=0, sticky='w', **pad)
        self.dv = tk.StringVar()
        ttk.Entry(frm, textvariable=self.dv, width=46).grid(row=0, column=1, sticky='we', **pad)
        ttk.Button(frm, text='浏览…', command=self._pick_doc).grid(row=0, column=2, **pad)

        ttk.Label(frm, text='Postman Collection (*.json):').grid(row=1, column=0, sticky='w', **pad)
        self.jv = tk.StringVar()
        ttk.Entry(frm, textvariable=self.jv, width=46).grid(row=1, column=1, sticky='we', **pad)
        ttk.Button(frm, text='浏览…', command=self._pick_json).grid(row=1, column=2, **pad)

        ttk.Label(frm, text='回填 Excel（可选）:').grid(row=2, column=0, sticky='w', **pad)
        self.fv = tk.StringVar()
        ttk.Entry(frm, textvariable=self.fv, width=46).grid(row=2, column=1, sticky='we', **pad)
        ttk.Button(frm, text='浏览…', command=self._pick_fill).grid(row=2, column=2, **pad)
        ttk.Label(frm, text='列：测试用例名称 / 预期结果 / 实际结果').grid(row=3, column=1, sticky='w', **pad)

        ttk.Label(frm, text='接口（按名称匹配用例）:').grid(row=4, column=0, sticky='w', **pad)
        self.names = ttk.Combobox(frm, state='readonly', width=46)
        self.names.grid(row=4, column=1, sticky='we', **pad)
        ttk.Label(frm, text='（下拉选择单个接口）').grid(row=4, column=2, sticky='w', **pad)

        btn = ttk.Button(frm, text='灌入所选接口', command=lambda: self._run(win))
        btn.grid(row=5, column=1, sticky='we', **pad)

        frm.columnconfigure(1, weight=1)

    def _pick_doc(self):
        p = filedialog.askopenfilename(title='选择中间 Word',
                                       filetypes=[('Word 文档', '*.docx')])
        if p:
            self.doc_path = p
            self.dv.set(p)

    def _pick_json(self):
        p = filedialog.askopenfilename(title='选择 Postman Collection',
                                       filetypes=[('Postman Collection', '*.json'), ('所有文件', '*.*')])
        if not p:
            return
        try:
            self.reqs = parse_postman_collection(p)
        except Exception as e:
            messagebox.showerror('解析失败', '无法解析 Postman Collection：\n{}'.format(e))
            return
        if not self.reqs:
            messagebox.showwarning('未找到接口', '该 Collection 中没有可用接口。')
            return
        self.jv.set(p)
        self.names['values'] = [r['name'] for r in self.reqs]
        if self.reqs:
            self.names.current(0)

    def _pick_fill(self):
        p = filedialog.askopenfilename(title='选择回填 Excel（可选）',
                                       filetypes=[('Excel 工作簿', '*.xlsx'), ('所有文件', '*.*')])
        if not p:
            return
        try:
            self.fill_map = _read_fill_excel(p)
        except Exception as e:
            messagebox.showerror('读取失败', '无法读取回填 Excel：\n{}'.format(e))
            return
        self.fv.set(p)

    def _run(self, win):
        if not self.doc_path:
            messagebox.showwarning('缺少文件', '请先选择中间 Word。')
            return
        if not self.reqs:
            messagebox.showwarning('缺少接口', '请先选择 Postman Collection。')
            return
        sel = self.names.get()
        if not sel:
            messagebox.showwarning('未选择接口', '请从下拉框选择一个接口。')
            return
        req = next((r for r in self.reqs if r['name'] == sel), None)
        if req is None:
            return
        out = filedialog.asksaveasfilename(title='保存灌入后的文档',
                                           defaultextension='.docx',
                                           filetypes=[('Word 文档', '*.docx')],
                                           initialfile=os.path.basename(self.doc_path))
        if not out:
            return
        steps = build_request_steps(req)
        fill = (self.fill_map or {}).get(sel)
        try:
            matched, msg = inject_postman_request(self.doc_path, out, sel, steps, fill=fill)
        except Exception as e:
            messagebox.showerror('灌入失败', str(e))
            return
        if matched:
            win.destroy()
            messagebox.showinfo('灌入成功', msg)
        else:
            messagebox.showwarning('未匹配', msg)


class TaskEditDialog:
    """单个批量任务的公共参数补填：默认沿用主窗口（第一次任务）的值，
    可补填该任务模板特有的字段与公共步骤；保存后该任务单独使用。"""

    _ROLE_DISP = [('step', '步骤'), ('expected', '期望结果'),
                  ('actual', '实际结果'), ('criterion', '评价准则')]

    def __init__(self, batch, task):
        self.batch = batch
        self.task = task
        app = batch.app
        ov = task['override']

        root = tk.Toplevel(batch.root)
        self.root = root
        root.title('补填公共参数 - ' + os.path.basename(task['src']))
        root.geometry('660x540')
        root.transient(batch.root)

        outer = ttk.Frame(root, padding=8)
        outer.pack(fill='both', expand=True)

        canvas = tk.Canvas(outer, highlightthickness=0)
        vsb = ttk.Scrollbar(outer, orient='vertical', command=canvas.yview)
        inner = ttk.Frame(canvas)
        win = canvas.create_window((0, 0), window=inner, anchor='nw')
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')
        inner.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.bind('<Configure>', lambda e: canvas.itemconfigure(win, width=e.width))
        canvas.bind_all('<MouseWheel>', lambda e: canvas.yview_scroll(int(-e.delta / 120), 'units'))

        frm = ttk.LabelFrame(inner, text=' 公共字段（默认沿用主窗口，可在此补填/修改） ', padding=4)
        frm.grid(sticky='we', padx=4, pady=4)
        main_iface = app._collect_interface()
        self.field_vars = {}
        for ri, f in enumerate(task['fields']):
            label = f['label']
            init = ov['fields'].get(label)
            if init is None:
                init = main_iface.get(label, '')
            ttk.Label(frm, text='{}:'.format(label)).grid(
                row=ri, column=0, sticky='nw', padx=4, pady=2)
            if f.get('key') in MULTI_KEYS:
                t = tk.Text(frm, width=60, height=2, wrap='word')
                t.grid(row=ri, column=1, sticky='we', padx=4, pady=2)
                if init:
                    t.insert('1.0', init)
                self.field_vars[label] = t
            else:
                v = tk.StringVar(value=init)
                ttk.Entry(frm, textvariable=v, width=70).grid(
                    row=ri, column=1, sticky='we', padx=4, pady=2)
                self.field_vars[label] = v
            frm.columnconfigure(1, weight=1)

        # ---- 公共步骤区（按该任务模板实际存在的明细列显示） ----
        self.step_txts = {}
        roles = set((dc.get('role') for dc in task['meta'].get('detail_cols', []) if dc.get('role')))
        cols = [(r, d) for r, d in self._ROLE_DISP if r in roles]
        if not cols:
            cols = [('step', '步骤'), ('expected', '期望结果')]
        sfrm = ttk.LabelFrame(
            inner, text=' 公共步骤（默认沿用主窗口；写一次插入每个用例步骤最前；可留空表示该任务无公共步骤） ',
            padding=4)
        sfrm.grid(sticky='we', padx=4, pady=4)
        main_txts = app._common_txts
        used_steps = ov['public_steps'] if ov.get('configured') else {}
        for ci, (role, disp) in enumerate(cols):
            sub = ttk.Frame(sfrm)
            sub.grid(row=0, column=ci, sticky='nsew', padx=4, pady=2)
            ttk.Label(sub, text=disp).pack(anchor='w')
            t = tk.Text(sub, height=4, width=30, wrap='word')
            init = used_steps.get(role)
            if init is None and role in main_txts:
                init = main_txts[role].get('1.0', 'end').strip()
            if init:
                t.insert('1.0', init)
            t.pack(fill='both', expand=True)
            self.step_txts[role] = t
            sfrm.columnconfigure(ci, weight=1)

        btns = ttk.Frame(inner)
        btns.grid(sticky='we', padx=4, pady=6)
        ttk.Button(btns, text='保存并用于该任务', command=self._save).pack(side='left', padx=4)
        ttk.Button(btns, text='取消', command=root.destroy).pack(side='left', padx=4)

    def _save(self):
        ov = self.task['override']
        fields = {}
        for label, w in self.field_vars.items():
            if isinstance(w, tk.Text):
                fields[label] = w.get('1.0', 'end').strip()
            else:
                fields[label] = w.get().strip()
        ov['fields'] = fields
        ov['public_steps'] = {role: t.get('1.0', 'end').strip()
                              for role, t in self.step_txts.items()}
        ov['configured'] = True
        self.batch.status.set('已添加 {} 个任务（含补填）'.format(len(self.batch.tasks)))
        self.root.destroy()


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == '__main__':
    main()