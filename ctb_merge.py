"""Merge a fresh CTB Clear Track export into Roman's working file.

usage: python -I merge.py NEW.xlsx OUT.xlsx OLD_WORKING.xlsx [EXTRA_OLD.xlsx ...]

- Grid data, HB tabs: always from NEW.
- LOB Health Status (col N): recomputed from NEW grid colors.
- Notes: harvested from every OLD file (PM Notes sheet, column N text, column N comments,
  Shortage Priority input columns) into one PM Notes log, carried forward.
- Master List_raw: copied from the first OLD file that has it.
"""
import sys, re, copy, datetime as dt
from collections import defaultdict, Counter
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter as L

NEW, OUT, OLDS = sys.argv[1], sys.argv[2], sys.argv[3:]
STOCK = 'FF5CBFAE'; ONTIME = 'FFB9E8DF'; Y = 'FFFBC72F'; O = 'FFEE7130'; R = 'FFC62828'
NOPEG = 'FFF8FAFC'; NOBOM = 'FFF1F3F5'; CURRENT = 'FFE2E8F0'; NOREC = 'FFFEF3C7'
BAD = {Y: '1-6d late', O: '7-13d late', R: '14d+ late', NOREC: 'no recovery in MD04 horizon'}
HDR = 15
SEV = {'Healthy': 0, 'Risk': 1, 'Unhealthy': 2}
FONT = 'Arial'
thin = Side(style='thin', color='D0D5DD'); BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
CF = {'Healthy': ('C6EFCE', '006100'), 'Risk': ('FFEB9C', '9C5700'), 'Unhealthy': ('FFC7CE', '9C0006')}
FAR = dt.datetime(2100, 1, 1)
NEXT_STEPS = ['Follow up with buyer', 'Follow up with planner', 'Escalate', 'Waiting on response',
              'Monitor delivery', 'Check SAP', 'Cleared', 'No action']
SUMMARY_TABS = {'How to Use', 'Outline CTB', 'Order CTB', 'Shortage Priority', 'PM Notes', 'What Changed',
                'Risk Watchlist', 'Buyer Rollup', 'Shared Components', 'Status Review Log', 'Master List_raw'}


# ------------------------------------------------------------------ parsing helpers
def rgb(c): return c.fill.fgColor.rgb if c.fill.fill_type else None


def parse_hdr(s):
    if not s: return None
    m = re.match(r'(\S+)\s+·\s+Qty\s+([\d.]+)', s)
    f = re.search(r'Finish (\d\d/\d\d/\d\d)', s); cl = re.search(r'Clear (\d\d/\d\d/\d\d)', s)
    sl = re.search(r'slip \+(\d+)d', s); p = re.search(r'Pacer (.*)$', s, re.M)
    return dict(order=m.group(1) if m else None, qty=float(m.group(2)) if m else None,
                finish=dt.datetime.strptime(f.group(1), '%m/%d/%y') if f else None,
                clear=dt.datetime.strptime(cl.group(1), '%m/%d/%y') if cl else None,
                slip=int(sl.group(1)) if sl else 0,
                pacers=[x.strip() for x in p.group(1).split(',')] if p else [])


def parse_comment(c):
    t = c.comment.text if c.comment else ''
    q = re.search(r'Qty for this order: ([\d.]+)', t)
    late = re.search(r'\+(\d+) days? late', t)
    docs = re.search(r'Supply documents?: (.*?)\.?$', t, re.M)
    return dict(qty=float(q.group(1)) if q else None, late=int(late.group(1)) if late else None,
                docs=docs.group(1).strip().rstrip('.') if docs else '', pastdue='PAST-DUE SUPPLY' in t,
                qm=bool(re.search(r'Supply documents?: QM', t)))


def is_ctb(ws): return str(ws['A1'].value or '').startswith('Clear Track')


def outline_of(ws):
    return str(ws['A1'].value).split('—')[-1].strip()


def asof(ws):
    m = re.search(r'as of (\d{4}-\d\d-\d\d)', str(ws['A3'].value))
    return dt.datetime.strptime(m.group(1), '%Y-%m-%d') if m else None


def lastcol(ws):
    lc = 15
    for col in range(15, ws.max_column + 1):
        if ws.cell(HDR, col).value: lc = col
    return lc


def read_ctb(ws):
    """Return dict with orders and component rows (status + gap detail)."""
    lc = lastcol(ws); run = asof(ws)
    hdrs = {col: parse_hdr(ws.cell(HDR, col).value) for col in range(15, lc + 1)}
    comps = {}
    for r in range(HDR + 1, ws.max_row + 1):
        comp = ws.cell(r, 1).value
        if comp in (None, ''): continue
        cells = [(col, ws.cell(r, col)) for col in range(15, lc + 1)]
        parsed = {col: parse_comment(c) for col, c in cells}
        unhealthy = False; reasons = []
        late_cols, dep_cols, qm_cols = [], [], []
        for col, c in cells:
            k = rgb(c); v = c.value; p = parsed[col]
            if k in BAD: unhealthy = True; late_cols.append(col)
            elif k == ONTIME or k == CURRENT or p['pastdue']: dep_cols.append(col)
            elif k == STOCK and (p['qm'] or (isinstance(v, dt.datetime) and run and v < run - dt.timedelta(days=4))):
                qm_cols.append(col)
        if late_cols: reasons.append('Late supply')
        if any(parsed[c]['pastdue'] for c, _ in cells): reasons.append('Past-due PO (ack date passed, not received)')
        if qm_cols: reasons.append('On hand but in Quality (QM lot)')
        if any(rgb(ws.cell(r, c)) == CURRENT for c in dep_cols): reasons.append('Current order, no assessment')
        if any(rgb(ws.cell(r, c)) == ONTIME for c in dep_cols): reasons.append('Depends on inbound delivery')
        status = 'Unhealthy' if unhealthy else ('Risk' if (dep_cols or qm_cols) else 'Healthy')
        d = dict(row=r, comp=comp, desc=ws.cell(r, 2).value,
                 supplier=(ws.cell(r, 3).value or '').split('·')[-1].strip(), qtyea=ws.cell(r, 4).value,
                 nett=ws.cell(r, 6).value, buyer=ws.cell(r, 10).value, planner=ws.cell(r, 11).value,
                 cc=ws.cell(r, 12).value, ss=ws.cell(r, 13).value,
                 cccat=(ws.cell(r, 12).comment.text.split('Command Center:')[-1].strip()
                        if ws.cell(r, 12).comment and 'Command Center' in ws.cell(r, 12).comment.text else ''),
                 status=status, reasons=reasons,
                 pacer=sum(1 for col in hdrs if hdrs[col] and str(comp) in hdrs[col]['pacers']))
        if not d['buyer']: d['buyer'] = f'No buyer (interplant) - planner {d["planner"]}'
        if late_cols:
            fc = late_cols[0]; h = hdrs[fc]; p = parsed[fc]; sup = ws.cell(r, fc).value
            d.update(gap_col=fc, need=h['finish'], order=h['order'], supply=sup, late=p['late'],
                     worst=max((parsed[c]['late'] or 0) for c in late_cols), nlate=len(late_cols),
                     qtylate=round(sum(parsed[c]['qty'] or 0 for c in late_cols), 2), docs=p['docs'],
                     last_supply=max([ws.cell(r, c).value for c in late_cols if isinstance(ws.cell(r, c).value, dt.datetime)] or [None]))
        dep = dep_cols or qm_cols
        if dep:
            col = dep[0]
            d.update(dep_date=ws.cell(r, col).value, dep_need=hdrs[col]['finish'], dep_order=hdrs[col]['order'],
                     dep_docs=parsed[col]['docs'], pastdue=any(parsed[c]['pastdue'] for c, _ in cells))
        comps[str(comp)] = d
    return dict(hdrs=hdrs, comps=comps, run=run, lc=lc)


# ------------------------------------------------------------------ notes harvesting
ENTRY = re.compile(r'(\d{1,2}/\d{1,2}/\d{4}),\s*([A-Za-z][^:\n]{1,40}?):\s*')


def split_entries(text):
    text = text.split('Comment:')[-1] if '[Threaded comment]' in text else text
    text = text.replace('--- Prior ---', '')
    out = []
    ms = list(ENTRY.finditer(text))
    for i, m in enumerate(ms):
        body = text[m.end(): ms[i + 1].start() if i + 1 < len(ms) else len(text)]
        body = ' '.join(body.split())
        author = re.sub(r'\s*\(PM\)$', '', m.group(2).strip())
        out.append((dt.datetime.strptime(m.group(1), '%m/%d/%Y'), author, body))
    return out


ledger = []          # dicts
seen = set()


def add_note(outline, comp, date, author, text, source='', esd='', step='', fu=None, origin=''):
    text = (text or '').strip()
    if not text: return
    if isinstance(date, dt.datetime): date = date.replace(hour=0, minute=0, second=0, microsecond=0)
    k = (outline, str(comp), date.date() if date else None, (author or '').lower()[:12], text[:50].lower())
    if k in seen: return
    seen.add(k)
    ledger.append(dict(outline=outline, comp=comp, date=date, source=source, author=author, text=text,
                       esd=esd, step=step, fu=fu, origin=origin))


master_src = None
old_grid = {}   # outline -> read_ctb of first old file with CTB tabs
for path in OLDS:
    wb = openpyxl.load_workbook(path)
    saved = wb.properties.modified or dt.datetime.now()
    tag = path.split('/')[-1]
    if master_src is None and 'Master List_raw' in wb.sheetnames: master_src = (path, wb)
    if 'PM Notes' in wb.sheetnames:            # future runs: the log is the source of truth
        ws = wb['PM Notes']
        hdr = [ws.cell(4, c).value for c in range(1, ws.max_column + 1)]
        ix = {h: i + 1 for i, h in enumerate(hdr) if h}
        for r in range(5, ws.max_row + 1):
            g = lambda h: ws.cell(r, ix[h]).value if h in ix else None
            if not g('Outline') or not g('Note'): continue
            add_note(str(g('Outline')), g('Component') or 'OUTLINE', g('Note Date'), g('Author'), g('Note'),
                     g('Source') or '', g('ESD / CTB Date') or '', g('Next Step') or '', g('Next Follow-Up'),
                     g('Carried From') or tag)
    for ws in wb.worksheets:
        if not is_ctb(ws): continue
        ol = outline_of(ws)
        if ol not in old_grid: old_grid[ol] = read_ctb(ws)
        for r in range(HDR + 1, ws.max_row + 1):
            comp = ws.cell(r, 1).value
            if comp in (None, ''): continue
            c = ws.cell(r, 14)
            if c.comment:
                for d, a, t in split_entries(c.comment.text):
                    pm = 'corona' in a.lower()
                    add_note(ol, comp, d, a, t, 'PM' if pm else 'Buyer / CC remark',
                             step='Follow up with buyer' if pm else '', origin=tag)
            v = c.value
            if isinstance(v, str) and v.strip() and v.strip() not in ('PEGGED', 'Healthy', 'Risk', 'Unhealthy'):
                add_note(ol, comp, saved, 'Corona, Roman', v.strip(), 'PM', origin=tag + ' (column N text, dated by file save)')
    if 'Shortage Priority' in wb.sheetnames:
        ws = wb['Shortage Priority']
        hdr = {ws.cell(4, c).value: c for c in range(1, ws.max_column + 1)}
        for r in range(5, ws.max_row + 1):
            esd = ws.cell(r, hdr.get('ESD / Delivery Note', 1)).value if 'ESD / Delivery Note' in hdr else None
            fu = ws.cell(r, hdr['Next Follow-Up']).value if 'Next Follow-Up' in hdr else None
            if not (esd or fu): continue
            ol, comp = str(ws.cell(r, hdr['Outline']).value), ws.cell(r, hdr['Component']).value
            req = ws.cell(r, hdr['Date Requested']).value if 'Date Requested' in hdr else None
            for e in ledger:   # attach to that day's PM note if there is one
                if e['outline'] == ol and str(e['comp']) == str(comp) and e['source'] == 'PM' and req and e['date'] and e['date'].date() == req.date():
                    e['esd'] = e['esd'] or esd; e['fu'] = e['fu'] or fu
                    e['source'] = (ws.cell(r, hdr['Status Source (CC / SAP / Buyer)']).value or 'PM') + ' + SAP'
                    break
            else:
                add_note(ol, comp, req or saved, 'Corona, Roman', 'Status request logged', 'PM', esd or '', 'Follow up with buyer', fu, tag)
ledger.sort(key=lambda e: (e['date'] or dt.datetime(1900, 1, 1)))

# ------------------------------------------------------------------ build on the NEW workbook
wb = openpyxl.load_workbook(NEW)
for ws in wb.worksheets:   # fix tab names that don't match the outline in A1
    if is_ctb(ws) and ws.title != outline_of(ws) and outline_of(ws) not in wb.sheetnames:
        print('renamed tab', ws.title, '->', outline_of(ws)); ws.title = outline_of(ws)
ctb_tabs = [ws for ws in wb.worksheets if is_ctb(ws)]
hb_tabs = [ws for ws in wb.worksheets if ws.title.startswith('HB ')]
new_grid = {ws.title: read_ctb(ws) for ws in ctb_tabs}
REPORT = max(g['run'] for g in new_grid.values())
TODAY = max(dt.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0), REPORT)

# HB: customer per order position, unit cost + make/buy per component
hb = {}
for ws in hb_tabs:
    ol = ws.title[3:].strip()
    comps = {str(ws.cell(1, c).value): dict(cost=ws.cell(4, c).value, mb=ws.cell(5, c).value, lvl=ws.cell(6, c).value)
             for c in range(4, ws.max_column + 1) if ws.cell(1, c).value}
    rows = [(ws.cell(r, 1).value, ws.cell(r, 2).value) for r in range(9, ws.max_row + 1) if ws.cell(r, 2).value]
    hb[ol] = dict(comps=comps, rows=rows)


def cust_for(ol, col):
    g = new_grid[ol]; cols = sorted(g['hdrs']); i = cols.index(col)
    rows = hb.get(ol, {}).get('rows', [])
    return rows[i][0] if i < len(rows) else ''


def cost_for(ol, comp):
    try: return float(hb[ol]['comps'][str(comp)]['cost'])
    except Exception: return None


# copy Master List_raw from old working file
if master_src:
    src = master_src[1]['Master List_raw']
    ml = wb.create_sheet('Master List_raw')
    for row in src.iter_rows():
        for c in row:
            n = ml.cell(c.row, c.column, c.value)
            if c.has_style:
                n.font = copy.copy(c.font); n.fill = copy.copy(c.fill); n.border = copy.copy(c.border)
                n.alignment = copy.copy(c.alignment); n.number_format = c.number_format
    for k, d in src.column_dimensions.items(): ml.column_dimensions[k].width = d.width
    for k, d in src.row_dimensions.items(): ml.row_dimensions[k].height = d.height
    for c in range(45, ml.max_column + 1):     # drop rollup columns from a previous merge; rebuilt below
        if c >= 45:
            for r in range(1, ml.max_row + 1): ml.cell(r, c).value = None
else:
    ml = None
mlrows = {}
if ml:
    for r in range(2, ml.max_row + 1):
        if ml.cell(r, 2).value is None: continue
        mlrows[str(ml.cell(r, 2).value).strip()] = dict(row=r, desc=ml.cell(r, 3).value, cust=ml.cell(r, 6).value,
                                                        opp=ml.cell(r, 22).value, notes=ml.cell(r, 27).value)


# ------------------------------------------------------------------ styling helpers
def add_status_cf(ws, rng):
    first = rng.split(':')[0].replace('$', '')
    for k, (bg, fg) in CF.items():
        ws.conditional_formatting.add(rng, CellIsRule(operator='equal', formula=[f'"{k}"'],
                                      fill=PatternFill('solid', fgColor=bg, bgColor=bg), font=Font(color=fg, bold=True)))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({first},2)="No"'],
                                  fill=PatternFill('solid', fgColor='E4E7EC', bgColor='E4E7EC'), font=Font(color='344054', bold=True)))


def header(ws, row, cols, widths, inputs=()):
    for i, (h, w) in enumerate(zip(cols, widths), 1):
        c = ws.cell(row, i, h)
        c.font = Font(name=FONT, bold=True, color='FFFFFF', size=10)
        c.fill = PatternFill('solid', fgColor='B54708' if i in inputs else '1E293B')
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True); c.border = BORDER
        ws.column_dimensions[L(i)].width = w
    ws.row_dimensions[row].height = 44


def body(ws, r, ncol, inputs=(), wrap=(), datefmt=()):
    for i in range(1, ncol + 1):
        c = ws.cell(r, i)
        c.font = Font(name=FONT, size=10); c.border = BORDER
        c.alignment = Alignment(vertical='top', wrap_text=i in wrap)
        if i in inputs: c.fill = PatternFill('solid', fgColor='FFF9DB')
        if isinstance(c.value, dt.datetime) or i in datefmt: c.number_format = 'mm/dd/yy'


def title(ws, text, sub, merge_to=12):
    ws['A1'] = text; ws['A1'].font = Font(name=FONT, bold=True, size=14, color='0F172A')
    ws['A2'] = sub; ws['A2'].font = Font(name=FONT, size=10, italic=True, color='475569')
    ws['A2'].alignment = Alignment(wrap_text=True, vertical='top')
    if merge_to > 1: ws.merge_cells(f'A2:{L(merge_to)}2')
    ws.row_dimensions[2].height = 30
    ws.sheet_view.showGridLines = False


def status_lookup(ol, ref):
    return f"=IFERROR(INDEX('{ol}'!$N:$N,MATCH({ref},'{ol}'!$A:$A,0)),\"Off report\")"


HR = 4
NOTE_LAST = 2000   # PM Notes rows covered by lookups


def note_lookup(key_ref, col_letter):
    rng = f"'PM Notes'!$A$5:$A${NOTE_LAST}"
    return f"=IFERROR(LOOKUP(2,1/({rng}={key_ref}),'PM Notes'!${col_letter}$5:${col_letter}${NOTE_LAST}),\"\")"


# ------------------------------------------------------------------ column N on each CTB tab
latest_note = defaultdict(list)
for e in ledger: latest_note[(e['outline'], str(e['comp']))].append(e)
for ws in ctb_tabs:
    g = new_grid[ws.title]
    ws.cell(HDR, 14).value = 'LOB Health Status'
    ws.cell(HDR, 14).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.column_dimensions['N'].width = max(ws.column_dimensions['N'].width or 0, 18)
    for comp, d in g['comps'].items():
        c = ws.cell(d['row'], 14); c.value = d['status']
        notes = latest_note.get((ws.title, comp), [])
        if notes:
            txt = '\n\n'.join(f"{e['date']:%m/%d/%Y}, {e['author']}: {e['text']}" + (f" | ESD: {e['esd']}" if e['esd'] else '')
                              for e in reversed(notes[-4:]))
            cm = Comment('PM Notes (carried forward; full history on PM Notes tab)\n\n' + txt, 'Corona, Roman')
            cm.width = 420; cm.height = 300; c.comment = cm
    add_status_cf(ws, f'N{HDR + 1}:N{ws.max_row}')
    dv = DataValidation(type='list', formula1='"Healthy,Risk,Unhealthy"', allow_blank=True)
    ws.add_data_validation(dv); dv.add(f'N{HDR + 1}:N{ws.max_row}')
    ws.sheet_properties.tabColor = '1E293B'
for ws in hb_tabs: ws.sheet_properties.tabColor = '98A2B3'

# ------------------------------------------------------------------ change detection
def change(old, new):
    if old is None: return 'New to report'
    if new is None: return 'Off report'
    if old == new: return f'Still {new}'
    if new == 'Healthy': return 'Cleared'
    return 'Worse' if SEV[new] > SEV[old] else 'Improved'


unhealthy, risks, changes = [], [], []
for ol, g in new_grid.items():
    og = old_grid.get(ol, {'comps': {}})
    for comp, d in g['comps'].items():
        o = og['comps'].get(comp)
        d['old_status'] = o['status'] if o else None
        d['old_supply'] = o.get('supply') if o else None
        d['change'] = change(d['old_status'], d['status'])
        if d['status'] == 'Unhealthy': unhealthy.append(dict(d, outline=ol))
        elif d['status'] == 'Risk': risks.append(dict(d, outline=ol))
        if d['change'] not in ('Still Healthy', 'Still Risk') and not (d['change'] == 'New to report' and d['status'] == 'Healthy'):
            changes.append(dict(d, outline=ol))
    for comp, o in og['comps'].items():
        if comp not in g['comps']:
            changes.append(dict(o, outline=ol, status=None, old_status=o['status'], change='Off report', old_supply=o.get('supply')))
for ol in old_grid:
    if ol not in new_grid: print('WARNING outline missing from new file:', ol)

by_comp = defaultdict(list)
for u in unhealthy: by_comp[str(u['comp'])].append(u['outline'])
for u in unhealthy: u['also'] = ', '.join(o for o in by_comp[str(u['comp'])] if o != u['outline'])
unhealthy.sort(key=lambda u: (u['need'] or FAR, -(u['late'] or 0), -u['pacer']))
risks.sort(key=lambda k: (k.get('dep_date') or FAR))

# ------------------------------------------------------------------ PM Notes (the log)
pn = wb.create_sheet('PM Notes')
title(pn, 'PM Notes - Running Log (carried forward every refresh)',
      'One row per note. Add new notes at the bottom; the newest row per Outline + Component feeds Shortage Priority and Outline CTB. '
      'Use Component = OUTLINE for an outline-level note or CTB date. Key column fills itself.', 11)
pcols = ['Key', 'Outline', 'Component', 'Note Date', 'Source', 'Author', 'Note', 'ESD / CTB Date', 'Next Step', 'Next Follow-Up', 'Carried From']
header(pn, HR, pcols, [22, 14, 16, 11, 16, 18, 70, 24, 20, 12, 26], inputs=range(2, 12))
for i, e in enumerate(ledger, 1):
    r = HR + i
    vals = [None, e['outline'], e['comp'], e['date'], e['source'], e['author'], e['text'], e['esd'], e['step'], e['fu'], e['origin']]
    for j, v in enumerate(vals, 1): pn.cell(r, j, v)
for r in range(HR + 1, HR + 301):
    pn.cell(r, 1, f'=IF(B{r}="","",B{r}&"|"&C{r})')
    body(pn, r, len(pcols), inputs=range(2, 12), wrap=(7, 8), datefmt=(4, 10))
    pn.cell(r, 1).font = Font(name=FONT, size=9, color='98A2B3')
dv = DataValidation(type='list', formula1='"' + ','.join(NEXT_STEPS) + '"', allow_blank=True)
pn.add_data_validation(dv); dv.add(f'I{HR + 1}:I{NOTE_LAST}')
dv2 = DataValidation(type='list', formula1='"PM,Buyer email,Command Center,SAP,Planner,Buyer / CC remark"', allow_blank=True)
pn.add_data_validation(dv2); dv2.add(f'E{HR + 1}:E{NOTE_LAST}')
pn.freeze_panes = f'C{HR + 1}'; pn.auto_filter.ref = f'A{HR}:K{HR + 300}'

# ------------------------------------------------------------------ Shortage Priority
sp = wb.create_sheet('Shortage Priority')
title(sp, 'Unhealthy Shortages - Prioritized by Earliest Gap',
      f'Data: CTB Clear Track run {REPORT:%m/%d/%Y}. Sorted by need-by (finish) of the first late order, then days late. '
      'Note columns (orange) read live from the PM Notes tab; log new statuses there.', 14)
cols = ['Priority #', 'Outline', 'Customer (1st Gap Order)', 'Opportunity', 'Component', 'Description', 'Make / Buy', 'Supplier',
        'Buyer', 'Planner', 'In CC?', 'CC Category', 'Pacer on # Orders', 'First Gap Need-By', 'First Gap Order',
        'Supply Date Shown', 'Days Late (1st Gap)', 'Worst Days Late', '# Late Orders', 'Qty on Late Orders',
        '$ on Late Orders', 'PO / Supply Doc', 'Days Until Need-By', 'Urgency', 'vs 10/01 Run', 'Supply Date Last Run',
        'Supply Moved (days)', 'Also Unhealthy On', 'Current LOB (live)', 'Latest Note', 'Note Date', 'Next Step',
        'Next Follow-Up', 'Key']
widths = [8, 14, 22, 14, 16, 24, 8, 24, 22, 20, 7, 16, 9, 11, 14, 11, 9, 9, 8, 9, 11, 26, 9, 18, 14, 11, 9, 14, 12, 55, 10, 18, 11, 6]
NI = (30, 31, 32, 33)
header(sp, HR, cols, widths, inputs=NI)
for i, u in enumerate(unhealthy, 1):
    r = HR + i; ol = u['outline']; m = mlrows.get(ol, {})
    cost = cost_for(ol, u['comp'])
    mb = (hb.get(ol, {}).get('comps', {}).get(str(u['comp'])) or {}).get('mb')
    moved = (u['supply'] - u['old_supply']).days if isinstance(u.get('old_supply'), dt.datetime) and isinstance(u['supply'], dt.datetime) else None
    vals = [i, ol, cust_for(ol, u['gap_col']), m.get('opp'), u['comp'], u['desc'], mb, u['supplier'], u['buyer'], u['planner'],
            'Y' if u['cc'] == 'TRUE' else '', u['cccat'], u['pacer'] or None, u['need'], u['order'], u['supply'], u['late'],
            u['worst'], u['nlate'], u['qtylate'], round(u['qtylate'] * cost, 2) if cost else None, u['docs'], None, None,
            u['change'], u.get('old_supply'), moved, u['also']]
    for j, v in enumerate(vals, 1): sp.cell(r, j, v)
    sp.cell(r, 23, f'=IF(N{r}="","",N{r}-TODAY())')
    sp.cell(r, 24, f'=IF(N{r}="","Check manually",IF(N{r}<TODAY(),"1 - Need date passed",IF(N{r}-TODAY()<=30,"2 - Next 30 days",IF(N{r}-TODAY()<=60,"3 - 31-60 days","4 - 60+ days"))))')
    sp.cell(r, 29, status_lookup(ol, f'E{r}'))
    sp.cell(r, 34, f'=B{r}&"|"&E{r}')
    for col, letter in zip((30, 31, 32, 33), ('G', 'D', 'I', 'J')): sp.cell(r, col, note_lookup(f'$AH{r}', letter))
    body(sp, r, len(cols), inputs=NI, wrap=(30,), datefmt=(14, 16, 26, 31, 33))
    sp.cell(r, 21).number_format = '$#,##0'; sp.cell(r, 34).font = Font(name=FONT, size=8, color='98A2B3')
last = HR + len(unhealthy)
sp.freeze_panes = f'F{HR + 1}'; sp.auto_filter.ref = f'A{HR}:{L(len(cols))}{last}'
add_status_cf(sp, f'AC{HR + 1}:AC{last}')
for txt, bg, fg in [('1 -', 'FFC7CE', '9C0006'), ('2 -', 'FFEB9C', '9C5700')]:
    sp.conditional_formatting.add(f'X{HR + 1}:X{last}', FormulaRule(formula=[f'LEFT(X{HR + 1},3)="{txt}"'],
                                  fill=PatternFill('solid', fgColor=bg, bgColor=bg), font=Font(color=fg, bold=True)))
sp.conditional_formatting.add(f'Y{HR + 1}:Y{last}', CellIsRule(operator='equal', formula=['"Worse"'], font=Font(color='B42318', bold=True)))
sp.conditional_formatting.add(f'Y{HR + 1}:Y{last}', CellIsRule(operator='equal', formula=['"New to report"'], font=Font(color='B54708', bold=True)))
sp.conditional_formatting.add(f'AA{HR + 1}:AA{last}', CellIsRule(operator='greaterThan', formula=['0'], font=Font(color='B42318', bold=True)))
sp.conditional_formatting.add(f'AA{HR + 1}:AA{last}', CellIsRule(operator='lessThan', formula=['0'], font=Font(color='067647', bold=True)))
sp.conditional_formatting.add(f'AD{HR + 1}:AD{last}', FormulaRule(formula=[f'AD{HR + 1}=""'], fill=PatternFill('solid', fgColor='FEE4E2', bgColor='FEE4E2')))

# ------------------------------------------------------------------ Outline CTB + Order CTB
oc = wb.create_sheet('Outline CTB')
title(oc, 'Outline Clear-to-Build Dates',
      f'One row per assigned outline. Clear dates come from the CTB order headers (run {REPORT:%m/%d/%Y}); "Clear" = date every component is available. '
      'PM CTB Date / Note read live from PM Notes (Component = OUTLINE).', 14)
ocols = ['Outline', 'Description', 'Customer(s)', 'Opportunity', 'LOB Status (live)', '# Unhealthy', '# Risk', '# Orders in Window',
         '# Orders CTB Now', '# Orders Late', 'Next Order Finish', 'Next Order Clear Date', 'All Orders Clear By', 'Max Slip (days)',
         'Main Pacers (orders paced)', 'All Clear By - 10/01 Run', 'Moved (days)', 'Orders Dropped Since 10/01',
         'PM CTB Date', 'PM Outline Note', 'Note Date', 'Key']
oinputs = (19, 20, 21)
header(oc, HR, ocols, [14, 24, 26, 14, 13, 9, 8, 9, 9, 8, 11, 11, 11, 9, 34, 11, 9, 20, 13, 50, 10, 6], inputs=oinputs)
order_rows = []
outlines = list(mlrows) if mlrows else list(new_grid)
for i, ol in enumerate(outlines, 1):
    r = HR + i; m = mlrows.get(ol, {})
    oc.cell(r, 1, ol); oc.cell(r, 2, m.get('desc')); oc.cell(r, 4, m.get('opp'))
    oc.cell(r, 22, f'=A{r}&"|OUTLINE"')
    for col, letter in zip((19, 20, 21), ('H', 'G', 'D')): oc.cell(r, col, note_lookup(f'$V{r}', letter))
    if ol in new_grid:
        g = new_grid[ol]; hd = [g['hdrs'][c] for c in sorted(g['hdrs'])]
        s = f"'{ol}'!$N:$N"
        oc.cell(r, 5, f'=IF(COUNTIF({s},"Unhealthy")>0,"Unhealthy",IF(COUNTIF({s},"Risk")>0,"Risk","Healthy"))')
        oc.cell(r, 6, f'=COUNTIF({s},"Unhealthy")'); oc.cell(r, 7, f'=COUNTIF({s},"Risk")')
        custs = sorted({c for c, _ in hb.get(ol, {}).get('rows', []) if c})
        oc.cell(r, 3, ', '.join(custs) or m.get('cust'))
        oc.cell(r, 8, len(hd))
        oc.cell(r, 9, sum(1 for h in hd if h['clear'] and h['clear'] <= REPORT))
        oc.cell(r, 10, sum(1 for h in hd if h['slip'] > 0))
        nxt = min(hd, key=lambda h: h['finish'])
        oc.cell(r, 11, nxt['finish']); oc.cell(r, 12, nxt['clear'])
        oc.cell(r, 13, max(h['clear'] for h in hd)); oc.cell(r, 14, max(h['slip'] for h in hd) or 0)
        pc = Counter(p for h in hd if h['slip'] > 0 for p in h['pacers'])
        oc.cell(r, 15, ', '.join(f'{p} ({n})' for p, n in pc.most_common(3)) or 'None late')
        if ol in old_grid:
            ohd = [old_grid[ol]['hdrs'][c] for c in sorted(old_grid[ol]['hdrs'])]
            oldmax = max(h['clear'] for h in ohd); oc.cell(r, 16, oldmax)
            oc.cell(r, 17, (max(h['clear'] for h in hd) - oldmax).days)
            dropped = [h['order'] for h in ohd if h['order'] not in {x['order'] for x in hd}]
            oc.cell(r, 18, f'{len(dropped)} (shipped, completed or moved out)' if dropped else '0')
        for c in sorted(g['hdrs']):
            h = g['hdrs'][c]
            order_rows.append((ol, h, cust_for(ol, c)))
    else:
        oc.cell(r, 5, 'No CTB Data'); oc.cell(r, 15, 'Did not populate in CTB - pull status from SAP (MD04) / Command Center')
    body(oc, r, len(ocols), inputs=oinputs, wrap=(3, 15, 20), datefmt=(11, 12, 13, 16, 19, 21))
    oc.cell(r, 22).font = Font(name=FONT, size=8, color='98A2B3')
olast = HR + len(outlines)
add_status_cf(oc, f'E{HR + 1}:E{olast}')
oc.conditional_formatting.add(f'Q{HR + 1}:Q{olast}', CellIsRule(operator='greaterThan', formula=['0'], font=Font(color='B42318', bold=True)))
oc.conditional_formatting.add(f'Q{HR + 1}:Q{olast}', CellIsRule(operator='lessThan', formula=['0'], font=Font(color='067647', bold=True)))
oc.freeze_panes = f'B{HR + 1}'; oc.auto_filter.ref = f'A{HR}:{L(len(ocols))}{olast}'

od = wb.create_sheet('Order CTB')
title(od, 'Order-Level Clear-to-Build (from CTB headers + HB customer rows)',
      'Each production order in the window with its customer (from the HB tab), finish date, clear date and pacing part.', 10)
dcols = ['Outline', 'Order', 'Customer', 'Qty', 'Finish (Need)', 'Clear Date', 'Slip (days)', 'CTB Status', 'Pacer(s)']
header(od, HR, dcols, [14, 16, 34, 6, 11, 11, 9, 16, 50])
for i, (ol, h, cust) in enumerate(order_rows, 1):
    r = HR + i
    st = 'CTB now' if h['clear'] and h['clear'] <= REPORT else ('On track' if h['slip'] == 0 else 'Late')
    for j, v in enumerate([ol, h['order'], cust, h['qty'], h['finish'], h['clear'], h['slip'], st,
                           ', '.join(h['pacers'][:6]) + (f' (+{len(h["pacers"]) - 6} more)' if len(h['pacers']) > 6 else '')], 1):
        od.cell(r, j, v)
    body(od, r, len(dcols), wrap=(3, 9))
dl = HR + len(order_rows)
od.conditional_formatting.add(f'H{HR + 1}:H{dl}', CellIsRule(operator='equal', formula=['"Late"'], fill=PatternFill('solid', fgColor='FFC7CE', bgColor='FFC7CE'), font=Font(color='9C0006', bold=True)))
od.conditional_formatting.add(f'H{HR + 1}:H{dl}', CellIsRule(operator='equal', formula=['"CTB now"'], fill=PatternFill('solid', fgColor='C6EFCE', bgColor='C6EFCE'), font=Font(color='006100', bold=True)))
od.freeze_panes = f'C{HR + 1}'; od.auto_filter.ref = f'A{HR}:I{dl}'

# ------------------------------------------------------------------ What Changed
wc = wb.create_sheet('What Changed')
title(wc, 'What Changed Since the 10/01 Run',
      'Status per component computed with the same color rule on both runs. Rows that were and still are Healthy, or still Risk, are left out.', 10)
order_key = {'Worse': 0, 'New to report': 1, 'Still Unhealthy': 2, 'Improved': 3, 'Cleared': 4, 'Off report': 5}
changes.sort(key=lambda d: (order_key.get(d['change'], 9), d['outline'], str(d['comp'])))
ccols = ['Outline', 'Component', 'Description', 'Buyer', 'Status 10/01', 'Status Now', 'Change', 'First Gap Supply 10/01', 'First Gap Supply Now', 'Moved (days)', 'Why (now)']
header(wc, HR, ccols, [14, 16, 26, 22, 11, 11, 14, 12, 12, 9, 40])
for i, d in enumerate(changes, 1):
    r = HR + i
    ns = d.get('supply') if d['status'] == 'Unhealthy' else None
    mv = (ns - d['old_supply']).days if isinstance(ns, dt.datetime) and isinstance(d.get('old_supply'), dt.datetime) else None
    for j, v in enumerate([d['outline'], d['comp'], d['desc'], d.get('buyer'), d['old_status'], d['status'], d['change'],
                           d.get('old_supply'), ns, mv, '; '.join(d['reasons']) if d['status'] else ''], 1):
        wc.cell(r, j, v)
    body(wc, r, len(ccols), wrap=(11,), datefmt=(8, 9))
wl = HR + len(changes)
add_status_cf(wc, f'E{HR + 1}:F{wl}')
for txt, col in [('Worse', 'B42318'), ('New to report', 'B54708'), ('Cleared', '067647'), ('Improved', '067647')]:
    wc.conditional_formatting.add(f'G{HR + 1}:G{wl}', CellIsRule(operator='equal', formula=[f'"{txt}"'], font=Font(color=col, bold=True)))
wc.freeze_panes = f'C{HR + 1}'; wc.auto_filter.ref = f'A{HR}:K{wl}'

# ------------------------------------------------------------------ Risk Watchlist
rw = wb.create_sheet('Risk Watchlist')
title(rw, 'Risk Watchlist - Parts Depending on Inbound Supply',
      'Not short yet, but only healthy if the delivery lands. Sorted by earliest inbound date.', 12)
rcols = ['#', 'Outline', 'Component', 'Description', 'Supplier', 'Buyer', 'In CC?', 'Why Risk', 'Earliest Inbound / QM Date',
         'Need-By (that order)', 'Order', 'Past-Due PO?', 'PO / Supply Doc', 'Status 10/01', 'Current LOB (live)', 'Latest Note', 'Key']
header(rw, HR, rcols, [6, 14, 16, 24, 24, 20, 7, 38, 12, 11, 14, 9, 28, 11, 12, 50, 6], inputs=(16,))
for i, k in enumerate(risks, 1):
    r = HR + i
    for j, v in enumerate([i, k['outline'], k['comp'], k['desc'], k['supplier'], k['buyer'], 'Y' if k['cc'] == 'TRUE' else '',
                           '; '.join(k['reasons']), k.get('dep_date'), k.get('dep_need'), k.get('dep_order'),
                           'Y' if k.get('pastdue') else '', k.get('dep_docs'), k['old_status']], 1):
        rw.cell(r, j, v)
    rw.cell(r, 15, status_lookup(k['outline'], f'C{r}'))
    rw.cell(r, 17, f'=B{r}&"|"&C{r}')
    rw.cell(r, 16, note_lookup(f'$Q{r}', 'G'))
    body(rw, r, len(rcols), inputs=(16,), wrap=(8, 16), datefmt=(9, 10))
    rw.cell(r, 17).font = Font(name=FONT, size=8, color='98A2B3')
rl = HR + len(risks)
add_status_cf(rw, f'N{HR + 1}:O{rl}')
rw.conditional_formatting.add(f'L{HR + 1}:L{rl}', CellIsRule(operator='equal', formula=['"Y"'], fill=PatternFill('solid', fgColor='FFC7CE', bgColor='FFC7CE'), font=Font(color='9C0006', bold=True)))
rw.freeze_panes = f'D{HR + 1}'; rw.auto_filter.ref = f'A{HR}:{L(len(rcols))}{rl}'

# ------------------------------------------------------------------ Buyer Rollup
br = wb.create_sheet('Buyer Rollup')
title(br, 'Buyer Rollup - One Email per Buyer', 'Live counts off Shortage Priority. Earliest need-by tells you whose inbox to hit first.', 6)
header(br, HR, ['Buyer', 'Unhealthy Lines', 'Earliest Need-By', 'Lines With No Note Yet', 'Outlines'], [34, 11, 13, 13, 60])
buyers = defaultdict(list)
for u in unhealthy: buyers[u['buyer']].append(u)
rngB = f"'Shortage Priority'!$I${HR + 1}:$I${last}"
for i, b in enumerate(sorted(buyers, key=lambda b: min(x['need'] or FAR for x in buyers[b])), 1):
    r = HR + i
    br.cell(r, 1, b); br.cell(r, 2, f'=COUNTIF({rngB},A{r})')
    br.cell(r, 3, f"=IF(B{r}=0,\"\",_xlfn.MINIFS('Shortage Priority'!$N${HR + 1}:$N${last},{rngB},A{r}))")
    br.cell(r, 4, f"=COUNTIFS({rngB},A{r},'Shortage Priority'!$AD${HR + 1}:$AD${last},\"\")")
    br.cell(r, 5, ', '.join(sorted({x['outline'] for x in buyers[b]})))
    body(br, r, 5, wrap=(5,), datefmt=(3,))
br.freeze_panes = f'A{HR + 1}'

# ------------------------------------------------------------------ Shared Components
sc = wb.create_sheet('Shared Components')
title(sc, 'Shared Components - Parts on More Than One Outline',
      'Same part, same stock: CTB pegs each outline on its own, so a shared part can look covered on one outline while the other consumes the same inventory. Status is live.', 10)
present = defaultdict(dict)
for ol, g in new_grid.items():
    for comp, d in g['comps'].items(): present[comp][ol] = d
multi = {k: v for k, v in present.items() if len(v) > 1}
souts = sorted({o for v in multi.values() for o in v})
shdr = ['Component', 'Description', 'Buyer', 'Nettable OH', '# Outlines'] + [f'Status: {o}' for o in souts] + ['Short on 2+?']
header(sc, HR, shdr, [16, 26, 22, 10, 9] + [15] * len(souts) + [10])
sev_sort = lambda k: (-max(SEV[d['status']] for d in multi[k].values()), -sum(SEV[d['status']] for d in multi[k].values()), k)
for i, k in enumerate(sorted(multi, key=sev_sort), 1):
    r = HR + i; d0 = next(iter(multi[k].values()))
    for j, v in enumerate([d0['comp'], d0['desc'], d0['buyer'], d0['nett'], len(multi[k])], 1): sc.cell(r, j, v)
    for j, o in enumerate(souts):
        sc.cell(r, 6 + j, f"=IFERROR(INDEX('{o}'!$N:$N,MATCH($A{r},'{o}'!$A:$A,0)),\"Not on BOM\")")
    rng = f'F{r}:{L(5 + len(souts))}{r}'
    sc.cell(r, 6 + len(souts), f'=IF(COUNTIF({rng},"Unhealthy")>1,"Y","")')
    body(sc, r, len(shdr))
sl = HR + len(multi)
if multi:
    add_status_cf(sc, f'F{HR + 1}:{L(5 + len(souts))}{sl}')
    sc.conditional_formatting.add(f'{L(6 + len(souts))}{HR + 1}:{L(6 + len(souts))}{sl}', CellIsRule(operator='equal', formula=['"Y"'], fill=PatternFill('solid', fgColor='FFC7CE', bgColor='FFC7CE'), font=Font(color='9C0006', bold=True)))
sc.freeze_panes = f'B{HR + 1}'; sc.auto_filter.ref = f'A{HR}:{L(len(shdr))}{sl}'

# ------------------------------------------------------------------ Master List rollup
if ml:
    mc0 = 45
    for i, h in enumerate(['LOB Health Status', '# Unhealthy Parts', '# Risk Parts', 'Next Order Clear Date', 'All Orders Clear By', 'PM CTB Date']):
        c = ml.cell(1, mc0 + i, h); c.font = Font(name=FONT, bold=True, color='FFFFFF', size=11)
        c.fill = PatternFill('solid', fgColor='1E293B'); c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ml.column_dimensions[L(mc0 + i)].width = [16, 11, 10, 12, 12, 12][i]
    for ol, m in mlrows.items():
        r = m['row']
        src = f"MATCH($B{r},'Outline CTB'!$A${HR + 1}:$A${olast},0)"
        for i, col in enumerate(['E', 'F', 'G', 'L', 'M', 'S']):
            ml.cell(r, mc0 + i, f"=IFERROR(INDEX('Outline CTB'!${col}${HR + 1}:${col}${olast},{src}),\"\")")
            ml.cell(r, mc0 + i).font = Font(name=FONT, size=11, bold=(i == 0))
            ml.cell(r, mc0 + i).alignment = Alignment(horizontal='center', vertical='top')
            if i >= 3: ml.cell(r, mc0 + i).number_format = 'mm/dd/yy'
    add_status_cf(ml, f'AS2:AS{ml.max_row}')

# ------------------------------------------------------------------ How to Use
hu = wb.create_sheet('How to Use')
title(hu, 'CTB Line of Balance - Working File', f'Merged {TODAY:%m/%d/%Y}: grid + HB data from the {REPORT:%m/%d} CTB run; notes carried from: ' + ', '.join(p.split('/')[-1] for p in OLDS), 1)
lines = [
    ('Where your work lives', True),
    ('PM Notes is the one place notes live. Each refresh carries it forward untouched, so log statuses there (one row per update). You can still drop comments into column N on an outline tab; the next merge picks those up too.', False),
    ('Column N (LOB Health Status) is recalculated from the new grid every refresh, so a part flips to Healthy on its own once the parts land. Your history stays in PM Notes and in the N cell comment.', False),
    ('', False),
    ('LOB Health Status rule', True),
    ('Healthy = every pegged cell dark green. Risk = light green (inbound PO), past-due PO, or stock sitting in Quality (QM lot). Unhealthy = yellow, orange, red, or "no recovery in MD04 horizon". Grey = ignored.', False),
    ('', False),
    ('Tabs, in working order', True),
    ('Outline CTB - the goal view: LOB status, next order clear date, all-orders-clear-by date and pacing parts per outline, vs the 10/01 run. Put your committed CTB date in PM Notes with Component = OUTLINE.', False),
    ('Shortage Priority - every Unhealthy part ranked by earliest gap, with customer and $ on late orders (from HB), and whether it got better or worse since 10/01.', False),
    ('What Changed - parts that cleared, improved, got worse or are new since 10/01.', False),
    ('Risk Watchlist, Buyer Rollup, Shared Components - same as before, rebuilt on the new data.', False),
    ('Order CTB - every production order with its customer, finish date, clear date and pacer.', False),
    ('Outline tabs (dark) are the CTB Clear Track grid. HB tabs (grey) are the Horse Blanket: the same grid turned sideways, one row per customer order with unit cost and make/buy per part.', False),
    ('', False),
    ('Next refresh', True),
    ('Send me the new CTB export plus this file. Grid and HB tabs get replaced, PM Notes and Master List carry over, and What Changed compares against this run.', False),
]
for i, (t, b) in enumerate(lines, 4):
    c = hu.cell(i, 1, t); c.font = Font(name=FONT, size=11 if b else 10, bold=b, color='0F172A' if b else '344054')
    c.alignment = Alignment(wrap_text=True, vertical='top')
hu.column_dimensions['A'].width = 140

# ------------------------------------------------------------------ tab order
front = ['How to Use', 'Outline CTB', 'Shortage Priority', 'PM Notes', 'What Changed', 'Risk Watchlist', 'Buyer Rollup',
         'Shared Components', 'Order CTB', 'Master List_raw']
for t in front: wb[t].sheet_properties.tabColor = wb[t].sheet_properties.tabColor or 'B54708'
order = [wb[t] for t in front if t in wb.sheetnames] + ctb_tabs + hb_tabs
wb._sheets = order
wb.active = 0
wb.save(OUT)
print('report', REPORT.date(), 'unhealthy', len(unhealthy), 'risk', len(risks), 'changes', Counter(d['change'] for d in changes))
print('ledger', len(ledger)); [print('  ', e['outline'], e['comp'], e['date'], e['author'], e['text'][:70], '| esd', str(e['esd'])[:40]) for e in ledger]
