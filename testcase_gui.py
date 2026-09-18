# -*- coding: utf-8 -*-
"""测试用例批量生成工具 - GUI（字段驱动版）
- 选择 Word 模板后识别表格全部字段，动态显示为"公共字段"输入框（每张表填相同内容）
- 未填写的字段从 Excel 读取，按第一行表头匹配模板字段，一行 = 一个用例
- 界面优先：界面有值的字段用界面值，界面未填才用 Excel
"""
import os
import threading
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

import openpyxl

from testcase_generator import (ALIAS_GROUPS, STEPS_COL_ALIASES,
                                _key_for_label, generate_document,
                                get_template_fields, load_template_config,
                                save_template_config)

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


def _align_steps(role_text):
    """把各明细角色文本按换行拆成逐段并对齐。role_text: {角色: 全文}。
    角色可取 step/expected/criterion/actual。返回 [{step,expected,criterion,...}]。"""
    lines = {}
    for role, text in (role_text or {}).items():
        lines[role] = [s.strip() for s in str(text).split('\n') if s.strip()]
    n = max((len(v) for v in lines.values()), default=0)
    steps = []
    for i in range(n):
        d = {}
        for role, arr in lines.items():
            d[role] = arr[i] if i < len(arr) else ''
        steps.append(d)
    return steps


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

        # ---- 输出文件 ----
        ttk.Label(frm, text='输出 Word (*.docx):').grid(row=4, column=0, sticky='w', **pad)
        self.out_var = tk.StringVar(value=self._default_output())
        ttk.Entry(frm, textvariable=self.out_var, width=50).grid(row=4, column=1, sticky='we', **pad)
        ttk.Button(frm, text='浏览…', command=self._pick_out).grid(row=4, column=2, **pad)

        # ---- 预览信息 ----
        self.info = tk.StringVar(value='尚未识别模板。')
        ttk.Label(frm, textvariable=self.info, foreground='#2069c5').grid(
            row=5, column=0, columnspan=4, sticky='w', **pad)

        # ---- 执行按钮 + 进度条 ----
        btnfrm = ttk.Frame(frm)
        btnfrm.grid(row=6, column=0, columnspan=4, **pad)
        self.gen_btn = ttk.Button(btnfrm, text='生成 Word 文档', command=self._generate)
        self.gen_btn.pack(side='left', padx=6)
        ttk.Button(btnfrm, text='退出', command=root.destroy).pack(side='left')
        self.progress = ttk.Progressbar(frm, mode='determinate', maximum=100)
        self.progress.grid(row=7, column=0, columnspan=4, sticky='we', padx=8, pady=(0, 2))

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

        for ri, f in enumerate(self.template_fields, start=1):
            label = f['label']
            key = f.get('key')
            ttk.Label(self.field_frame, text=f'{label}:').grid(
                row=ri, column=0, sticky='nw', padx=6, pady=2)
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

    def _generate(self):
        if not self.template_fields:
            messagebox.showwarning('提示', '请先识别模板字段。')
            return
        if self.xl_var.get():
            self._load_preview()

        iface = self._collect_interface()
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
            steps = er['steps'] if er else []
            name = (fields.get(name_label) if name_label else '') or ('用例%d' % idx)
            cases.append({'name': name, 'fields': fields, 'steps': steps})

        out = self.out_var.get()
        if not out:
            messagebox.showwarning('提示', '请指定输出文件路径。')
            return

        # ---- 必填校验（Feature2）：用例名称字段为空的用例需用户确认 ----
        if name_label:
            missing = [c for c in cases if not (c['fields'].get(name_label) or '').strip()]
        else:
            missing = [c for c in cases if not (c.get('name') or '').strip()]
        if missing:
            if not messagebox.askyesno(
                    '完整性校验',
                    '有 {} 个用例的"用例名称"为空，输出后难以区分。\n\n'
                    '仍继续生成吗？（选"否"可先回填名称）'.format(len(missing))):
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


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == '__main__':
    main()