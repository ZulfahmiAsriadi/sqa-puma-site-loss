import sys
import os

# Prevent crash when running headless/detached where sys.stdout is None
if sys.stdout is None:
    try:
        sys.stdout = open(os.devnull, "w")
    except Exception:
        pass
if sys.stderr is None:
    try:
        sys.stderr = open(os.devnull, "w")
    except Exception:
        pass

import datetime
current_time_str = datetime.datetime.now().strftime("%d %b %Y, %H:%M WIT")
build_timestamp_hash = str(int(datetime.datetime.now().timestamp()))
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
import json
import os

excel_path = r"C:\Users\Zulfahmi Asriadi\.gemini\antigravity-ide\scratch\tracking_temp.xlsx"
z = zipfile.ZipFile(excel_path, 'r')

shared_strings = []
if 'xl/sharedStrings.xml' in z.namelist():
    with z.open('xl/sharedStrings.xml') as f:
        for event, elem in ET.iterparse(f):
            if elem.tag.endswith('si'):
                shared_strings.append("".join(elem.itertext()))
                elem.clear()

def parse_cell(c_elem):
    t = c_elem.attrib.get('t')
    v_elem = c_elem.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
    if v_elem is None:
        v_elem = c_elem.find('v')
    if v_elem is not None and v_elem.text:
        val = v_elem.text
        if t == 's':
            idx = int(val)
            return shared_strings[idx] if idx < len(shared_strings) else f"[s:{idx}]"
        return val
    is_elem = c_elem.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}is')
    if is_elem is None:
        is_elem = c_elem.find('is')
    if is_elem is not None:
        return "".join(is_elem.itertext())
    return None

def col_letter(ref):
    return ''.join([c for c in ref if c.isalpha()])

# Dynamic sheet resolution & column mapping for sheet Aug 26 & SEP '26
from collections import Counter

wb_tree = ET.fromstring(z.read('xl/workbook.xml'))
rels_tree = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
rels = {rel.attrib['Id']: rel.attrib['Target'] for rel in rels_tree}

sheet_aug_path = None
sheet_sep_path = None
sheet_ssgj_path = None
ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
for sheet in wb_tree.findall('.//main:sheet', ns):
    s_name = sheet.attrib.get('name', '').strip()
    s_low = s_name.lower()
    r_id = sheet.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
    target = rels.get(r_id)
    if target:
        full_target = 'xl/' + target if not target.startswith('xl/') else target
        if 'aug' in s_low and '26' in s_low:
            sheet_aug_path = full_target
        elif 'sep' in s_low and '26' in s_low:
            sheet_sep_path = full_target
        elif 'ssgj' in s_low:
            sheet_ssgj_path = full_target

if not sheet_aug_path:
    sheet_aug_path = 'xl/worksheets/sheet7.xml'
if not sheet_sep_path:
    sheet_sep_path = 'xl/worksheets/sheet8.xml'
if not sheet_ssgj_path and 'xl/worksheets/sheet10.xml' in z.namelist():
    sheet_ssgj_path = 'xl/worksheets/sheet10.xml'

# Load SSGJ lookup for any new sites in SEP '26
ssgj_lookup = {}
if sheet_ssgj_path and sheet_ssgj_path in z.namelist():
    with z.open(sheet_ssgj_path) as f_ssgj:
        for event, elem in ET.iterparse(f_ssgj, events=('end',)):
            if elem.tag.endswith('row'):
                if int(elem.attrib.get('r', 0)) > 1:
                    d_s = {}
                    for c in elem:
                        val = parse_cell(c)
                        if val is not None:
                            d_s[col_letter(c.attrib.get('r', ''))] = val
                    if d_s.get('A'):
                        ssgj_lookup[d_s['A'].strip()] = d_s
                elem.clear()

SPECIAL_KAB_MAP = {
    'EPW018': 'BIAK NUMFOR',
    'EPW296': 'RAJA AMPAT',
    'COK137': 'TELUK WONDAMA',
    'COK336': 'NABIRE',
    'COK033': 'MERAUKE',
}

def eval_priority_sales(gap, rhi, trans, avail, pl):
    g = str(gap or '').strip().upper()
    r = str(rhi or '').strip().upper()
    t = str(trans or '').strip().upper()
    a = str(avail or '').strip().upper()
    p = str(pl or '').strip().upper()
    t_main = t in ('RADIO IP TELKOM', 'RADIO IP TELKOMSEL', 'FO TELKOM', 'FO BTS HOTEL')
    t_ext = t_main or t in ('OFDM', 'OFDM CLUSTER', 'PRT', 'STARLINK')
    if g == '<3 MIO' and r == 'PASS' and t_main and a == 'ACHIEVED' and p == 'CLEAR':
        return 'P1'
    if g == '3-5 MIO' and r == 'PASS' and t_main and a == 'ACHIEVED' and p == 'CLEAR':
        return 'P2'
    if g == '5-10 MIO' and r == 'FAIL' and t_ext and a == 'NOT ACHIEVED' and p in ('SPIKE', 'CONSECUTIVE'):
        return 'P3'
    return 'Not Priority'

def eval_priority_net(gap, rhi, trans, avail, pl):
    g = str(gap or '').strip().upper()
    r = str(rhi or '').strip().upper()
    t = str(trans or '').strip().upper()
    a = str(avail or '').strip().upper()
    p = str(pl or '').strip().upper()
    t_main = t in ('RADIO IP TELKOM', 'RADIO IP TELKOMSEL', 'FO TELKOM', 'FO BTS HOTEL')
    t_ext = t_main or t in ('OFDM', 'OFDM CLUSTER', 'PRT', 'STARLINK')
    if g == '<3 MIO' and r == 'FAIL' and t_main and a == 'NOT ACHIEVED' and p in ('SPIKE', 'CONSECUTIVE'):
        return 'P1'
    if g == '3-5 MIO' and r == 'FAIL' and t_main and a == 'NOT ACHIEVED' and p in ('SPIKE', 'CONSECUTIVE'):
        return 'P2'
    if g == '5-10 MIO' and r == 'FAIL' and t_ext and a == 'NOT ACHIEVED' and p in ('SPIKE', 'CONSECUTIVE'):
        return 'P3'
    return 'Not Priority'

def parse_sheet_dataset(sheet_path, fallback_aug_raw=None, prefix_kab_map=None):
    headers_r3 = {}
    with z.open(sheet_path) as f_hdr:
        for event, elem in ET.iterparse(f_hdr, events=('end',)):
            if elem.tag.endswith('row'):
                if int(elem.attrib.get('r', 0)) == 3:
                    for c in elem:
                        ref = c.attrib.get('r', '')
                        val = parse_cell(c)
                        if val is not None:
                            headers_r3[col_letter(ref)] = str(val).strip().lower()
                    break
                elem.clear()

    is_compact_sep = (len(headers_r3) <= 30 and 'kabupaten' not in headers_r3.values())
    is_w_appended_sep = (headers_r3.get('w') == 'branch' and 'revenue' in headers_r3.get('e', ''))

    col_breakdown = 'BB' if is_w_appended_sep else 'AZ'
    col_nilai_cost = 'BC' if is_w_appended_sep else None
    col_plan_improv = 'BD' if is_w_appended_sep else 'BA'
    col_cek = 'BE' if is_w_appended_sep else 'BB'

    for let, text in headers_r3.items():
        if 'breakdown cost' in text:
            col_breakdown = let
        elif 'nilai cost' in text:
            col_nilai_cost = let
        elif 'plan improvement' in text:
            col_plan_improv = let
        elif text == 'cek' or 'kategori' in text:
            col_cek = let

    col_letters = {
        'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M',
        'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
        'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK',
        'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS',
        'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 'BA',
        'BB', 'BC', 'BD', 'BE', 'BF', 'BG'
    }

    parsed_rows = []
    raw_by_sid = {}
    branch_kab = defaultdict(lambda: defaultdict(lambda: {
        'total': 0, 'profit': 0, 'loss': 0,
        'rev': 0, 'cost': 0, 'margin': 0,
        'gap_tiers': {'<3 Mio': 0, '3-5 Mio': 0, '5-10 Mio': 0, '>10 Mio': 0}
    }))

    def to_int(v):
        try:
            return round(float(v))
        except Exception:
            return 0

    def to_pct(v):
        try:
            return round(float(v) * 100, 2)
        except Exception:
            return 0.0

    with z.open(sheet_path) as f:
        for event, elem in ET.iterparse(f, events=('end',)):
            if elem.tag.endswith('row'):
                r_idx = int(elem.attrib.get('r', 0))
                if r_idx > 3:
                    c_dict = {}
                    for c in elem:
                        ref = c.attrib.get('r', '')
                        val = parse_cell(c)
                        let = col_letter(ref)
                        if let in col_letters and val is not None:
                            c_dict[let] = val

                    if c_dict.get('A'):
                        sid = str(c_dict.get('A', '')).strip()
                        raw_by_sid[sid] = c_dict
                        fb = (fallback_aug_raw or {}).get(sid, {})
                        s_ssgj = ssgj_lookup.get(sid, {})

                        if is_compact_sep:
                            rev = to_int(c_dict.get('E'))
                            cost = to_int(c_dict.get('F'))
                            margin = to_int(c_dict.get('G'))
                            margin_pct = to_pct(c_dict.get('H'))
                            pnl = c_dict.get('I', 'Profit')
                            nop = c_dict.get('J') or fb.get('P', '-')
                            branch = fb.get('Q') or s_ssgj.get('AE') or nop.replace('NOP ', '').strip() or 'Unknown'
                            kab = fb.get('R') or s_ssgj.get('AG') or SPECIAL_KAB_MAP.get(sid)
                            if not kab and prefix_kab_map:
                                mc = prefix_kab_map.get((sid[:3], nop), Counter()).most_common(1)
                                kab = mc[0][0] if mc else 'KOTA AMBON'
                            kab = kab or 'Unknown'
                            type_site = c_dict.get('K') or fb.get('S', '')
                            bts_equip = fb.get('T') or ('Combat' if type_site == 'Combat' else 'Macro')
                            site_class = fb.get('U', 'BRONZE')
                            status_oa = fb.get('V', 'OA')
                            inline_ssgj = fb.get('K') or ('Inline' if sid in ssgj_lookup else 'Not Inline')
                            prog_group = fb.get('L') or s_ssgj.get('C', '')
                            if pnl == 'Profit':
                                gap = 'Profit'
                            else:
                                am = abs(margin)
                                if am < 3000000: gap = '<3 Mio'
                                elif am < 5000000: gap = '3-5 Mio'
                                elif am < 10000000: gap = '5-10 Mio'
                                else: gap = '>10 Mio'
                            if pnl == 'Profit':
                                p_sales = 'Not Priority'
                                p_net = 'Not Priority'
                            else:
                                old_ps = fb.get('N', 'Not Priority')
                                old_pn = fb.get('O', 'Not Priority')
                                p_sales = old_ps if old_ps in ('P1', 'P2', 'P3') else eval_priority_sales(gap, fb.get('AN'), fb.get('AT'), fb.get('AV'), fb.get('AX'))
                                p_net = old_pn if old_pn in ('P1', 'P2', 'P3') else eval_priority_net(gap, fb.get('AN'), fb.get('AT'), fb.get('AV'), fb.get('AX'))
                            status_period = c_dict.get('T') or fb.get('AD', '')
                            src_dict = fb
                        elif is_w_appended_sep:
                            rev = to_int(c_dict.get('E'))
                            cost = to_int(c_dict.get('F'))
                            margin = to_int(c_dict.get('G'))
                            margin_pct = to_pct(c_dict.get('H'))
                            pnl = c_dict.get('I', 'Profit')
                            nop = c_dict.get('J') or fb.get('P', '-')
                            branch = c_dict.get('W') or fb.get('Q') or nop.replace('NOP ', '').strip() or 'Unknown'
                            kab = c_dict.get('X') or fb.get('R') or SPECIAL_KAB_MAP.get(sid) or 'Unknown'
                            gap = c_dict.get('Y', '')
                            p_sales = c_dict.get('Z', 'Not Priority')
                            p_net = c_dict.get('AA', 'Not Priority')
                            inline_ssgj = c_dict.get('AB', '')
                            prog_group = c_dict.get('AC', '')
                            type_site = c_dict.get('K', '')
                            bts_equip = c_dict.get('AD', '')
                            site_class = c_dict.get('AE', '-')
                            status_oa = c_dict.get('AF', '-')
                            status_period = c_dict.get('T', '')
                            # Remap W..BF to standard AH..BD keys so the rest of row extraction works identically
                            src_dict = {
                                'AE': c_dict.get('AG', ''),
                                'AF': c_dict.get('AH', ''),
                                'AG': c_dict.get('AI', ''),
                                'AH': c_dict.get('AJ', ''),
                                'AI': c_dict.get('AK', ''),
                                'AJ': c_dict.get('AL', ''),
                                'AK': c_dict.get('AM', ''),
                                'AL': c_dict.get('AN', ''),
                                'AM': c_dict.get('AO', ''),
                                'AN': c_dict.get('AP', ''),
                                'AO': c_dict.get('AQ', ''),
                                'AP': c_dict.get('AR', ''),
                                'AQ': c_dict.get('AS', ''),
                                'AR': c_dict.get('AT', ''),
                                'AS': c_dict.get('AU', ''),
                                'AT': c_dict.get('AV', ''),
                                'AU': c_dict.get('AW', ''),
                                'AV': c_dict.get('AX', ''),
                                'AW': c_dict.get('AY', ''),
                                'AX': c_dict.get('AZ', ''),
                                'AY': c_dict.get('BA', ''),
                                col_breakdown: c_dict.get('BB', ''),
                                (col_nilai_cost or 'BA'): c_dict.get('BC', ''),
                                col_plan_improv: c_dict.get('BD', ''),
                                col_cek: c_dict.get('BE', ''),
                            }
                        else:
                            rev = to_int(c_dict.get('F'))
                            cost = to_int(c_dict.get('G'))
                            margin = to_int(c_dict.get('H'))
                            margin_pct = to_pct(c_dict.get('I'))
                            pnl = c_dict.get('J', '')
                            gap = c_dict.get('M', '')
                            p_sales = c_dict.get('N', 'Not Priority')
                            p_net = c_dict.get('O', 'Not Priority')
                            nop = c_dict.get('P', '-')
                            branch = c_dict.get('Q') or fb.get('Q') or nop.replace('NOP ', '').strip() or 'Unknown'
                            kab = c_dict.get('R') or fb.get('R') or SPECIAL_KAB_MAP.get(sid) or 'Unknown'
                            type_site = c_dict.get('S', '')
                            bts_equip = c_dict.get('T', '')
                            site_class = c_dict.get('U', '-')
                            status_oa = c_dict.get('V', '-')
                            inline_ssgj = c_dict.get('K', '')
                            prog_group = c_dict.get('L', '')
                            status_period = c_dict.get('AD', '')
                            src_dict = c_dict

                        payload_val = src_dict.get('AF')
                        payload_formatted = '-'
                        try:
                            p_float = float(payload_val)
                            payload_formatted = str(round(p_float / (1024*1024))) + " GB"
                        except Exception:
                            if payload_val:
                                payload_formatted = str(payload_val)

                        action_sqa = src_dict.get('AH', '')
                        update_action_sqa = src_dict.get('AI', '')
                        action_sales = src_dict.get('AJ', '')
                        update_action_sales = src_dict.get('AK', '')
                        action_nop = src_dict.get('AL', '')
                        update_action_nop = src_dict.get('AM', '')

                        breakdown_val = str(src_dict.get(col_breakdown, '') or '').strip()
                        if breakdown_val in ['#N/A', '-', 'None', 'nan']:
                            breakdown_val = ''

                        nilai_cost_val = ''
                        nc_col = col_nilai_cost or 'BA'
                        raw_nc = src_dict.get(nc_col, '')
                        if raw_nc and str(raw_nc).strip() not in ['#N/A', '-', 'None', 'nan']:
                            try:
                                nilai_cost_val = str(round(float(raw_nc)))
                            except Exception:
                                nilai_cost_val = str(raw_nc).strip()

                        pi_col = col_plan_improv if not is_compact_sep else 'BB'
                        plan_improv_val = str(src_dict.get(pi_col, '') or '').strip()
                        if plan_improv_val in ['#N/A', '-', 'None', 'nan']:
                            plan_improv_val = ''

                        ck_col = col_cek if not is_compact_sep else 'BC'
                        cek_val = str(src_dict.get(ck_col, '') or '').strip()
                        if cek_val in ['#N/A', '-', 'None', 'nan']:
                            cek_val = ''

                        row = [
                            sid,                        # 0: siteid
                            c_dict.get('B') or fb.get('B', ''), # 1: sitename
                            nop,                        # 2: nop (Column P)
                            branch,                     # 3: branch (Column Q)
                            kab,                        # 4: kabupaten (Column R)
                            type_site,                  # 5: type_site (Column S)
                            bts_equip,                  # 6: bts_equip (Column T)
                            site_class,                 # 7: site_class (Column U)
                            status_oa,                  # 8: status_oa (Column V)
                            rev,                        # 9: revenue (Column F)
                            cost,                       # 10: cost (Column G)
                            margin,                     # 11: margin (Column H)
                            margin_pct,                 # 12: margin_pct (Column I)
                            pnl,                        # 13: pnl (Column J)
                            gap,                        # 14: gap_dist (Column M)
                            inline_ssgj,                # 15: inline_ssgj (Column K)
                            prog_group,                 # 16: prog_group (Column L)
                            p_sales,                    # 17: priority_sales (Column N)
                            p_net,                      # 18: priority_network (Column O)
                            payload_formatted,          # 19: payload (Column AF)
                            action_sqa,                 # 20: action_sqa (Column AH)
                            update_action_sqa,          # 21: update_action_sqa (Column AI)
                            action_sales,               # 22: action_sales (Column AJ)
                            update_action_sales,        # 23: update_action_sales (Column AK)
                            action_nop,                 # 24: action_nop (Column AL)
                            update_action_nop,          # 25: update_action_nop (Column AM)
                            src_dict.get('AN', ''),     # 26: rhi_last (AN)
                            src_dict.get('AO', ''),     # 27: rhi_root_cause (AO)
                            src_dict.get('AP', ''),     # 28: rhi_propose (AP)
                            src_dict.get('AQ', ''),     # 29: thi_last (AQ)
                            src_dict.get('AR', ''),     # 30: thi_root_cause (AR)
                            src_dict.get('AS', ''),     # 31: thi_propose (AS)
                            src_dict.get('AT', ''),     # 32: transport_type (AT)
                            src_dict.get('AU', ''),     # 33: bdbh_util (AU)
                            src_dict.get('AV', ''),     # 34: avail_status (AV)
                            src_dict.get('AW', ''),     # 35: plan_avail (AW)
                            src_dict.get('AX', ''),     # 36: status_pl_last (AX)
                            src_dict.get('AY', ''),     # 37: propose_pl (AY)
                            breakdown_val,              # 38: breakdown_cost
                            plan_improv_val,            # 39: plan_improv_cost
                            status_period,              # 40: status_period (AD)
                            src_dict.get('AE', ''),     # 41: intervention (AE)
                            src_dict.get('AG', ''),     # 42: site_loss_tracker (AG)
                            cek_val,                    # 43: col_cek (BC)
                            nilai_cost_val              # 44: nilai_cost (BA)
                        ]
                        parsed_rows.append(row)

                        entry = branch_kab[branch][kab]
                        entry['total'] += 1
                        entry['rev'] += rev
                        entry['cost'] += cost
                        entry['margin'] += margin
                        if pnl == 'Profit':
                            entry['profit'] += 1
                        elif pnl == 'Loss':
                            entry['loss'] += 1
                            if gap in entry['gap_tiers']:
                                entry['gap_tiers'][gap] += 1

                elem.clear()

    k_list = []
    for br, kabs in branch_kab.items():
        for kab, data in kabs.items():
            pct_loss = round((data['loss'] / data['total'] * 100), 1) if data['total'] > 0 else 0.0
            pct_profit = round((data['profit'] / data['total'] * 100), 1) if data['total'] > 0 else 0.0
            k_list.append({
                'branch': br,
                'kabupaten': kab,
                'total': data['total'],
                'profit': data['profit'],
                'loss': data['loss'],
                'pct_loss': pct_loss,
                'pct_profit': pct_profit,
                'rev': data['rev'],
                'cost': data['cost'],
                'margin': data['margin'],
                'tier_lt3': data['gap_tiers']['<3 Mio'],
                'tier_3_5': data['gap_tiers']['3-5 Mio'],
                'tier_5_10': data['gap_tiers']['5-10 Mio'],
                'tier_gt10': data['gap_tiers']['>10 Mio']
            })
    k_list.sort(key=lambda x: (x['loss'], x['pct_loss']), reverse=True)

    b_summary = {}
    for br in ['AMBON', 'JAYAPURA', 'SORONG', 'TIMIKA']:
        kabs_for_br = [k for k in k_list if k['branch'] == br and k['loss'] > 0]
        kabs_for_br.sort(key=lambda x: x['loss'], reverse=True)
        tot_loss_br = sum(k['loss'] for k in kabs_for_br)
        b_summary[br] = {
            'total_loss': tot_loss_br,
            'kabs': [
                {
                    'name': k['kabupaten'],
                    'loss': k['loss'],
                    'total': k['total'],
                    'pct_of_branch_loss': round(k['loss'] / tot_loss_br * 100, 1) if tot_loss_br else 0,
                    'tier_gt10': k['tier_gt10']
                } for k in kabs_for_br
            ]
        }

    all_kabs_with_loss = [k for k in k_list if k['loss'] > 0]
    all_kabs_with_loss.sort(key=lambda x: x['loss'], reverse=True)
    tot_loss_all = sum(k['loss'] for k in all_kabs_with_loss)
    b_summary['ALL'] = {
        'total_loss': tot_loss_all,
        'kabs': [
            {
                'name': f"{k['kabupaten']} ({k['branch']})",
                'loss': k['loss'],
                'total': k['total'],
                'pct_of_branch_loss': round(k['loss'] / tot_loss_all * 100, 1) if tot_loss_all else 0,
                'tier_gt10': k['tier_gt10']
            } for k in all_kabs_with_loss[:8]
        ]
    }
    return parsed_rows, k_list, b_summary, raw_by_sid

def build_period_meta(p_rows, p_kabs, period_key):
    br_stats = {}
    gap_mat = {}
    tot_all = len(p_rows)
    prof_all = sum(1 for r in p_rows if r[13] == 'Profit')
    loss_all = sum(1 for r in p_rows if r[13] == 'Loss')

    for br in ['AMBON', 'JAYAPURA', 'SORONG', 'TIMIKA']:
        b_rows = [r for r in p_rows if r[3] == br]
        btot = len(b_rows)
        bprof = sum(1 for r in b_rows if r[13] == 'Profit')
        bloss = sum(1 for r in b_rows if r[13] == 'Loss')
        btarget = round(btot * 0.93)
        bgap = bprof - btarget
        ppct = round(bprof / btot * 100, 2) if btot else 0.0
        lpct = round(bloss / btot * 100, 2) if btot else 0.0
        br_stats[br] = {
            'total_num': btot,
            'profit_num': bprof,
            'loss_num': bloss,
            'target_num': btarget,
            'gap_num': bgap,
            'total': f"{btot:,}",
            'profit': f"{bprof:,}",
            'loss': f"{bloss:,}",
            'target': f"{btarget:,}",
            'gap': f"{bgap:+d}" if bgap != 0 else "0",
            'pPct': f"{ppct:.2f}%",
            'lPct': f"{lpct:.2f}%",
            'pPctNum': ppct
        }
        lt3 = sum(1 for r in b_rows if r[13] == 'Loss' and r[14] == '<3 Mio')
        m35 = sum(1 for r in b_rows if r[13] == 'Loss' and r[14] == '3-5 Mio')
        m510 = sum(1 for r in b_rows if r[13] == 'Loss' and r[14] == '5-10 Mio')
        gt10 = sum(1 for r in b_rows if r[13] == 'Loss' and r[14] == '>10 Mio')
        gap_mat[br] = {'lt3': lt3, 'm3_5': m35, 'm5_10': m510, 'gt10': gt10, 'total': bloss}

    target_all = sum(br_stats[b]['target_num'] for b in ['AMBON', 'JAYAPURA', 'SORONG', 'TIMIKA'])
    gap_all = prof_all - target_all
    ppct_all = round(prof_all / tot_all * 100, 2) if tot_all else 0.0
    lpct_all = round(loss_all / tot_all * 100, 2) if tot_all else 0.0
    br_stats['ALL'] = {
        'total_num': tot_all,
        'profit_num': prof_all,
        'loss_num': loss_all,
        'target_num': target_all,
        'gap_num': gap_all,
        'total': f"{tot_all:,}",
        'profit': f"{prof_all:,}",
        'loss': f"{loss_all:,}",
        'target': f"{target_all:,}",
        'gap': f"{gap_all:+d}" if gap_all != 0 else "0",
        'pPct': f"{ppct_all:.2f}%",
        'lPct': f"{lpct_all:.2f}%",
        'pPctNum': ppct_all
    }
    lt3_t = sum(gap_mat[b]['lt3'] for b in ['AMBON', 'JAYAPURA', 'SORONG', 'TIMIKA'])
    m35_t = sum(gap_mat[b]['m3_5'] for b in ['AMBON', 'JAYAPURA', 'SORONG', 'TIMIKA'])
    m510_t = sum(gap_mat[b]['m5_10'] for b in ['AMBON', 'JAYAPURA', 'SORONG', 'TIMIKA'])
    gt10_t = sum(gap_mat[b]['gt10'] for b in ['AMBON', 'JAYAPURA', 'SORONG', 'TIMIKA'])
    gap_mat['TOTAL'] = {
        'lt3': lt3_t, 'm3_5': m35_t, 'm5_10': m510_t, 'gt10': gt10_t, 'total': loss_all,
        'lt3_pct': round(lt3_t / loss_all * 100, 1) if loss_all else 0.0,
        'm3_5_pct': round(m35_t / loss_all * 100, 1) if loss_all else 0.0,
        'm5_10_pct': round(m510_t / loss_all * 100, 1) if loss_all else 0.0,
        'gt10_pct': round(gt10_t / loss_all * 100, 1) if loss_all else 0.0
    }
    nop_counts = {
        nop: sum(1 for r in p_rows if r[2] == nop)
        for nop in ['NOP AMBON', 'NOP JAYAPURA', 'NOP SORONG', 'NOP TIMIKA']
    }
    oa_count = sum(1 for r in p_rows if r[8] == 'OA')
    not_oa_count = tot_all - oa_count
    h2_all = sum(1 for r in p_rows if r[42] == 'Program H2')
    h2_loss = sum(1 for r in p_rows if r[13] == 'Loss' and r[42] == 'Program H2')
    return {
        'periodKey': period_key,
        'rows': p_rows,
        'kabData': p_kabs,
        'branchStats': br_stats,
        'gapMatrix': gap_mat,
        'nopCounts': nop_counts,
        'oaCount': oa_count,
        'notOaCount': not_oa_count,
        'h2AllCount': h2_all,
        'h2LossCount': h2_loss
    }

aug_rows, aug_kab_list, aug_branch_kab_summary, aug_raw_by_sid = parse_sheet_dataset(sheet_aug_path)

prefix_kab_map = {}
for sid_k, d_k in aug_raw_by_sid.items():
    kab_k = d_k.get('R', '')
    nop_k = d_k.get('P', '')
    if kab_k:
        prefix_kab_map.setdefault((sid_k[:3], nop_k), Counter())[kab_k] += 1

sep_rows, sep_kab_list, sep_branch_kab_summary, _ = parse_sheet_dataset(
    sheet_sep_path, fallback_aug_raw=aug_raw_by_sid, prefix_kab_map=prefix_kab_map
)
z.close()

sep_meta = build_period_meta(sep_rows, sep_kab_list, 'SEP26')
sep_meta['branchKabSummary'] = sep_branch_kab_summary
aug_meta = build_period_meta(aug_rows, aug_kab_list, 'AUG26')
aug_meta['branchKabSummary'] = aug_branch_kab_summary

# Default active dataset on load is SEP '26
rows = sep_rows
kab_list = sep_kab_list
branch_kab_summary = sep_branch_kab_summary

json_aug_rows = json.dumps(rows, separators=(',', ':'))
json_kab_list = json.dumps(kab_list, separators=(',', ':'))
json_branch_kab_summary = json.dumps(branch_kab_summary, separators=(',', ':'))
json_period_datasets = json.dumps({'SEP26': sep_meta, 'AUG26': aug_meta}, separators=(',', ':'))

import base64
_logo_file = r"C:\Users\Zulfahmi Asriadi\sqa-puma-site-loss\sqa_puma_logo.png"
if not os.path.exists(_logo_file):
    _logo_file = r"c:\Users\Zulfahmi Asriadi\OneDrive - PT Telekomunikasi Selular\SQA PUMA\Site loss\sqa_puma_logo.png"
with open(_logo_file, "rb") as _lf:
    _logo_b64 = base64.b64encode(_lf.read()).decode("utf-8")
logo_data_uri = f"data:image/png;base64,{_logo_b64}"

html_template = f"""<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=5.0, user-scalable=yes">
  <title>SQA Puma Update Site LOSS</title>
  <script>
    (function() {{
      try {{
        var params = new URLSearchParams(window.location.search);
        var q = (params.get('theme') || params.get('mode') || '').toLowerCase();
        var theme = (q === 'light' || q === 'white' || q === 'putih') ? 'light'
                  : (q === 'dark' || q === 'gelap') ? 'dark'
                  : (localStorage.getItem('sqa_puma_theme') === 'light' ? 'light' : 'dark');
        document.documentElement.setAttribute('data-theme', theme);
      }} catch (e) {{
        document.documentElement.setAttribute('data-theme', 'dark');
      }}
    }})();
  </script>
  <link rel="icon" type="image/png" href="sqa_puma_logo.png">
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;600&display=swap" rel="stylesheet">
  <style>
    :root {{
      --bg-primary: #090d16;
      --bg-secondary: #0f172a;
      --bg-card: rgba(30, 41, 59, 0.65);
      --bg-card-hover: rgba(51, 65, 85, 0.75);
      --border-color: rgba(255, 255, 255, 0.08);
      --border-accent: rgba(255, 42, 68, 0.3);
      --text-main: #f8fafc;
      --text-muted: #94a3b8;
      --text-dim: #64748b;
      --tsel-red: #ff2a44;
      --tsel-red-glow: rgba(255, 42, 68, 0.25);
      --accent-green: #10b981;
      --accent-amber: #f59e0b;
      --accent-cyan: #06b6d4;
      --accent-blue: #3b82f6;
      --accent-purple: #8b5cf6;
      --radius-sm: 8px;
      --radius-md: 14px;
      --radius-lg: 20px;
      --shadow-lg: 0 20px 25px -5px rgba(0, 0, 0, 0.4), 0 8px 10px -6px rgba(0, 0, 0, 0.4);
    }}

    * {{ box-sizing: border-box; margin: 0; padding: 0; }}

    body {{
      font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      background: radial-gradient(circle at 15% 15%, #131c31 0%, #080d18 55%, #050810 100%);
      color: var(--text-main);
      min-height: 100vh;
      padding: 24px;
      line-height: 1.5;
      overflow-x: hidden;
      font-feature-settings: "zero" 0;
    }}

    .dashboard-container {{ max-width: 1720px; margin: 0 auto; }}

    .glass {{
      background: var(--bg-card);
      backdrop-filter: blur(16px);
      -webkit-backdrop-filter: blur(16px);
      border: 1px solid var(--border-color);
      border-radius: var(--radius-md);
      box-shadow: var(--shadow-lg);
    }}

    header {{
      padding: 22px 30px;
      margin-bottom: 24px;
      display: flex;
      flex-wrap: wrap;
      justify-content: space-between;
      align-items: center;
      gap: 20px;
      position: relative;
      overflow: hidden;
    }}

    header::after {{
      content: '';
      position: absolute;
      top: 0;
      left: 0;
      width: 100%;
      height: 3px;
      background: linear-gradient(90deg, var(--tsel-red), #ff8a00, var(--accent-cyan), var(--tsel-red));
      background-size: 300% 100%;
      animation: gradientShift 8s ease infinite;
    }}

    @keyframes gradientShift {{
      0% {{ background-position: 0% 50%; }}
      50% {{ background-position: 100% 50%; }}
      100% {{ background-position: 0% 50%; }}
    }}

    .brand-section {{ display: flex; align-items: center; gap: 18px; }}

    .brand-logo {{
      width: 48px;
      height: 48px;
      background: linear-gradient(135deg, #ffffff, #f8fafc);
      border-radius: 12px;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 4px;
      box-shadow: 0 4px 16px rgba(255, 42, 68, 0.3), 0 0 0 1px rgba(255, 255, 255, 0.15);
      flex-shrink: 0;
      transition: transform 0.25s cubic-bezier(0.4, 0, 0.2, 1), box-shadow 0.25s cubic-bezier(0.4, 0, 0.2, 1);
      overflow: hidden;
    }}

    .brand-logo:hover {{
      transform: translateY(-2px) scale(1.05);
      box-shadow: 0 8px 25px rgba(255, 42, 68, 0.45), 0 0 0 1px rgba(255, 42, 68, 0.4);
    }}

    .brand-logo img {{
      width: 100%;
      height: 100%;
      object-fit: contain;
      display: block;
    }}

    .brand-text h1 {{
      font-size: 25px;
      font-weight: 800;
      letter-spacing: -0.5px;
      background: linear-gradient(180deg, #ffffff, #cbd5e1);
      -webkit-background-clip: text;
      -webkit-text-fill-color: transparent;
    }}

    .brand-text p {{
      font-size: 13px;
      color: var(--text-muted);
      display: flex;
      align-items: center;
      gap: 8px;
      margin-top: 2px;
    }}

    .badge-live {{
      display: inline-flex;
      align-items: center;
      gap: 6px;
      padding: 3px 10px;
      background: rgba(16, 185, 129, 0.15);
      border: 1px solid rgba(16, 185, 129, 0.3);
      color: #34d399;
      font-size: 11px;
      font-weight: 600;
      border-radius: 20px;
    }}

    .pulse-dot {{
      width: 7px;
      height: 7px;
      background: #34d399;
      border-radius: 50%;
      box-shadow: 0 0 0 0 rgba(52, 211, 153, 0.7);
      animation: pulse 2s infinite;
    }}

    @keyframes pulse {{
      0% {{ box-shadow: 0 0 0 0 rgba(52, 211, 153, 0.7); }}
      70% {{ box-shadow: 0 0 0 8px rgba(52, 211, 153, 0); }}
      100% {{ box-shadow: 0 0 0 0 rgba(52, 211, 153, 0); }}
    }}

    .header-actions {{ display: flex; align-items: center; gap: 12px; }}

    .meta-tag {{
      background: rgba(255, 255, 255, 0.05);
      border: 1px solid var(--border-color);
      padding: 6px 14px;
      border-radius: var(--radius-sm);
      font-size: 12px;
      color: var(--text-muted);
      font-family: 'JetBrains Mono', monospace;
    }}

    .btn-action {{
      background: rgba(255, 42, 68, 0.15);
      border: 1px solid rgba(255, 42, 68, 0.4);
      color: #ff6b7d;
      padding: 8px 16px;
      border-radius: var(--radius-sm);
      font-size: 13px;
      font-weight: 600;
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      gap: 8px;
      transition: all 0.2s ease;
    }}

    .btn-action:hover {{
      background: var(--tsel-red);
      color: #fff;
      transform: translateY(-1px);
      box-shadow: 0 6px 18px var(--tsel-red-glow);
    }}

    /* Filter Bar */
    .filter-bar {{
      display: flex;
      align-items: center;
      justify-content: space-between;
      flex-wrap: wrap;
      gap: 14px;
      padding: 14px 24px;
      margin-bottom: 24px;
    }}

    .filter-group {{
      display: flex;
      align-items: center;
      gap: 8px;
      flex-wrap: wrap;
    }}

    .filter-label {{
      font-size: 12px;
      color: var(--text-dim);
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 0.8px;
      margin-right: 6px;
    }}

    .filter-pill {{
      background: rgba(255, 255, 255, 0.04);
      border: 1px solid var(--border-color);
      color: var(--text-muted);
      padding: 7px 16px;
      border-radius: 30px;
      font-size: 13px;
      font-weight: 600;
      cursor: pointer;
      transition: all 0.25s ease;
    }}

    .filter-pill:hover {{ background: rgba(255, 255, 255, 0.08); color: #fff; }}

    .filter-pill.active {{
      background: linear-gradient(135deg, var(--tsel-red), #c9182b);
      border-color: transparent;
      color: #fff;
      box-shadow: 0 4px 12px var(--tsel-red-glow);
    }}

    .target-badge {{
      background: rgba(245, 158, 11, 0.15);
      border: 1px solid rgba(245, 158, 11, 0.35);
      color: #fbbf24;
      padding: 4px 10px;
      border-radius: 6px;
      font-weight: 700;
      font-family: 'JetBrains Mono', monospace;
      font-size: 13px;
    }}

    /* KPI Grid (All rounded except percentages) */
    .kpi-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
      gap: 20px;
      margin-bottom: 24px;
    }}

    .kpi-card {{
      padding: 22px 24px;
      position: relative;
      overflow: hidden;
      transition: all 0.25s ease;
    }}

    .kpi-card:hover {{
      transform: translateY(-3px);
      border-color: rgba(255, 255, 255, 0.15);
    }}

    .kpi-header {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 12px;
    }}

    .kpi-title {{
      font-size: 13px;
      font-weight: 600;
      color: var(--text-muted);
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }}

    .kpi-icon {{
      width: 36px;
      height: 36px;
      border-radius: 10px;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 16px;
    }}

    .kpi-val {{
      font-size: 32px;
      font-weight: 800;
      letter-spacing: -1px;
      color: #fff;
      margin-bottom: 6px;
      display: flex;
      align-items: baseline;
      gap: 6px;
    }}

    .kpi-val .unit {{
      font-size: 14px;
      font-weight: 600;
      color: var(--text-dim);
    }}

    .kpi-sub {{
      font-size: 12px;
      color: var(--text-muted);
      display: flex;
      align-items: center;
      gap: 6px;
    }}

    .kpi-sub .trend-up {{ color: #34d399; font-weight: 700; }}
    .kpi-sub .trend-down {{ color: #f87171; font-weight: 700; }}

    .progress-bar-container {{
      width: 100%;
      height: 6px;
      background: rgba(255, 255, 255, 0.08);
      border-radius: 10px;
      margin-top: 14px;
      overflow: hidden;
    }}

    .progress-bar-fill {{
      height: 100%;
      border-radius: 10px;
      transition: width 0.6s cubic-bezier(0.4, 0, 0.2, 1);
    }}

    /* Section Grid */
    .section-grid-2 {{
      display: grid;
      grid-template-columns: 2fr 1.15fr;
      gap: 24px;
      margin-bottom: 24px;
    }}

    @media (max-width: 1024px) {{
      .section-grid-2 {{ grid-template-columns: 1fr; }}
    }}

    .chart-card {{ padding: 24px 28px; }}

    .card-title-bar {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 20px;
      flex-wrap: wrap;
      gap: 12px;
    }}

    .card-title {{
      font-size: 16px;
      font-weight: 700;
      color: #fff;
      display: flex;
      align-items: center;
      gap: 10px;
    }}

    .card-title .dot {{
      width: 8px;
      height: 8px;
      border-radius: 50%;
      background: var(--tsel-red);
    }}

    .chart-legend {{ display: flex; gap: 14px; flex-wrap: wrap; }}

    .legend-item {{
      display: flex;
      align-items: center;
      gap: 6px;
      font-size: 12px;
      color: var(--text-muted);
    }}

    .legend-color {{ width: 10px; height: 10px; border-radius: 3px; }}

    /* SVG Chart */
    .chart-svg-container {{
      width: 100%;
      height: 310px;
      position: relative;
    }}

    svg.chart-svg {{ width: 100%; height: 100%; overflow: visible; }}

    .grid-line {{
      stroke: rgba(255, 255, 255, 0.05);
      stroke-dasharray: 4, 4;
    }}

    .axis-label {{
      font-family: 'JetBrains Mono', monospace;
      font-size: 11px;
      fill: var(--text-dim);
    }}

    .chart-line {{
      fill: none;
      stroke-width: 3.5;
      stroke-linecap: round;
      stroke-linejoin: round;
      transition: all 0.3s ease;
    }}

    .data-point {{
      cursor: pointer;
      transition: all 0.2s ease;
    }}

    .data-point:hover {{ r: 6.5; stroke-width: 3; }}

    /* Main Tabs Bar */
    .tabs-wrapper {{ margin-bottom: 24px; }}

    .tabs-nav {{
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      border-bottom: 1px solid var(--border-color);
      padding-bottom: 16px;
      margin-bottom: 20px;
    }}

    .tab-btn {{
      background: rgba(255, 255, 255, 0.04);
      border: 1px solid var(--border-color);
      border-radius: 10px;
      color: var(--text-muted);
      padding: 10px 18px;
      font-size: 13.5px;
      font-weight: 600;
      font-family: inherit;
      cursor: pointer;
      position: relative;
      transition: all 0.2s ease;
      white-space: nowrap;
      display: inline-flex;
      align-items: center;
      gap: 8px;
    }}

    .tab-btn:hover {{
      background: rgba(255, 255, 255, 0.08);
      color: #fff;
    }}

    .tab-btn.active {{
      background: linear-gradient(135deg, var(--tsel-red), #c9182b);
      border-color: transparent;
      color: #fff;
      box-shadow: 0 4px 14px var(--tsel-red-glow);
    }}

    /* Top Tab Switcher directly in Branch Focus Bar */
    .branch-tab-switcher {{
      display: flex;
      align-items: center;
      gap: 6px;
      padding-left: 14px;
      border-left: 1px solid rgba(255, 255, 255, 0.12);
    }}
    @media (max-width: 900px) {{
      .branch-tab-switcher {{
        border-left: none;
        padding-left: 0;
        width: 100%;
        margin-top: 8px;
      }}
    }}
    .top-subtab-btn {{
      background: rgba(255, 255, 255, 0.05);
      border: 1px solid rgba(255, 255, 255, 0.12);
      color: var(--text-muted);
      padding: 6px 14px;
      border-radius: 20px;
      font-size: 12px;
      font-weight: 600;
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      gap: 6px;
      transition: all 0.2s ease;
      white-space: nowrap;
    }}
    .top-subtab-btn:hover {{
      color: #fff;
      border-color: rgba(255, 255, 255, 0.25);
      background: rgba(255, 255, 255, 0.08);
    }}
    .top-subtab-btn.active {{
      background: linear-gradient(135deg, var(--tsel-red) 0%, #b91c1c 100%);
      border-color: var(--tsel-red);
      color: #fff;
      box-shadow: 0 2px 10px var(--tsel-red-glow);
    }}

    /* Sub-Tabs Bar inside Branch Focus */
    .subtabs-bar {{
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      margin-bottom: 20px;
      background: rgba(15, 23, 42, 0.6);
      padding: 6px;
      border-radius: 12px;
      border: 1px solid rgba(255, 255, 255, 0.08);
      width: fit-content;
    }}

    .subtab-btn {{
      background: transparent;
      border: 1px solid transparent;
      color: var(--text-muted);
      padding: 9px 20px;
      border-radius: 8px;
      font-size: 13px;
      font-weight: 700;
      font-family: inherit;
      cursor: pointer;
      transition: all 0.2s ease;
      display: inline-flex;
      align-items: center;
      gap: 8px;
    }}

    .subtab-btn:hover {{
      color: #fff;
      background: rgba(255, 255, 255, 0.06);
    }}

    .subtab-btn.active {{
      background: linear-gradient(135deg, #f59e0b, #d97706);
      color: #090d16;
      font-weight: 800;
      box-shadow: 0 4px 12px rgba(245, 158, 11, 0.3);
    }}

    .branch-subtab-pane {{ display: none; }}
    .branch-subtab-pane.active {{
      display: block;
      animation: fadeIn 0.3s ease;
    }}

    .tab-pane {{ display: none; }}
    .tab-pane.active {{
      display: block;
      animation: fadeIn 0.3s ease;
    }}

    @keyframes fadeIn {{
      from {{ opacity: 0; transform: translateY(6px); }}
      to {{ opacity: 1; transform: translateY(0); }}
    }}

    /* Table Styles */
    .table-container {{
      overflow-x: auto;
      border-radius: var(--radius-sm);
      max-height: 650px;
      position: relative;
    }}

    table.data-table {{
      width: 100%;
      border-collapse: collapse;
      text-align: left;
      font-size: 12.5px;
    }}

    table.data-table th {{
      background: rgba(15, 23, 42, 0.95);
      color: var(--text-muted);
      font-weight: 700;
      text-transform: uppercase;
      font-size: 11px;
      letter-spacing: 0.6px;
      padding: 13px 16px;
      border-bottom: 1px solid var(--border-color);
      white-space: nowrap;
      position: sticky;
      top: 0;
      z-index: 10;
      backdrop-filter: blur(8px);
    }}

    table.data-table td {{
      padding: 12px 16px;
      border-bottom: 1px solid rgba(255, 255, 255, 0.04);
      color: var(--text-main);
      white-space: nowrap;
      font-variant-numeric: normal;
    }}

    table.data-table tbody tr:hover {{ background: rgba(255, 255, 255, 0.035); }}

    table.data-table tr.total-row {{
      background: rgba(255, 42, 68, 0.06);
      font-weight: 700;
      border-top: 2px solid rgba(255, 42, 68, 0.3);
    }}

    th.new-col-th {{
      background: rgba(6, 182, 212, 0.18);
      border-bottom: 2px solid #06b6d4;
      color: #67e8f9;
    }}

    td.new-col-td {{ background: rgba(6, 182, 212, 0.03); }}

    /* Badges */
    .badge {{
      display: inline-flex;
      align-items: center;
      gap: 5px;
      padding: 3px 9px;
      border-radius: 20px;
      font-size: 11px;
      font-weight: 700;
      font-family: 'JetBrains Mono', monospace;
    }}

    .badge-success {{
      background: rgba(16, 185, 129, 0.15);
      color: #34d399;
      border: 1px solid rgba(16, 185, 129, 0.3);
    }}

    .badge-danger {{
      background: rgba(239, 68, 68, 0.15);
      color: #f87171;
      border: 1px solid rgba(239, 68, 68, 0.3);
    }}

    .badge-warning {{
      background: rgba(245, 158, 11, 0.15);
      color: #fbbf24;
      border: 1px solid rgba(245, 158, 11, 0.3);
    }}

    .badge-info {{
      background: rgba(6, 182, 212, 0.15);
      color: #38bdf8;
      border: 1px solid rgba(6, 182, 212, 0.3);
    }}

    .badge-purple {{
      background: rgba(139, 92, 246, 0.15);
      color: #c084fc;
      border: 1px solid rgba(139, 92, 246, 0.3);
    }}

    /* Specific Badges for Site Class */
    .badge-diamond {{
      background: rgba(56, 189, 248, 0.2);
      color: #38bdf8;
      border: 1px solid #38bdf8;
    }}
    .badge-platinum {{
      background: rgba(192, 132, 252, 0.2);
      color: #c084fc;
      border: 1px solid #c084fc;
    }}
    .badge-gold {{
      background: rgba(251, 191, 36, 0.2);
      color: #fbbf24;
      border: 1px solid #fbbf24;
    }}
    .badge-silver {{
      background: rgba(148, 163, 184, 0.2);
      color: #cbd5e1;
      border: 1px solid #94a3b8;
    }}
    .badge-bronze {{
      background: rgba(217, 119, 6, 0.2);
      color: #f59e0b;
      border: 1px solid #d97706;
    }}

    .table-controls-bar {{
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      margin-bottom: 16px;
      background: rgba(15, 23, 42, 0.4);
      padding: 14px 18px;
      border-radius: var(--radius-sm);
      border: 1px solid var(--border-color);
    }}

    .search-input-wrapper {{
      display: flex;
      align-items: center;
      gap: 10px;
      flex: 1;
      min-width: 260px;
      max-width: 400px;
      position: relative;
    }}

    .search-input {{
      width: 100%;
      background: rgba(255, 255, 255, 0.05);
      border: 1px solid var(--border-color);
      border-radius: var(--radius-sm);
      padding: 8px 14px 8px 36px;
      font-size: 13px;
      color: #fff;
      font-family: inherit;
      outline: none;
      transition: all 0.2s ease;
    }}

    .search-input:focus {{
      border-color: var(--accent-cyan);
      box-shadow: 0 0 0 3px rgba(6, 182, 212, 0.2);
    }}

    .search-icon {{
      position: absolute;
      left: 12px;
      color: var(--text-dim);
      pointer-events: none;
    }}

    .filter-select {{
      background: rgba(255, 255, 255, 0.05);
      border: 1px solid var(--border-color);
      color: var(--text-main);
      padding: 8px 14px;
      border-radius: var(--radius-sm);
      font-size: 12.5px;
      font-family: inherit;
      outline: none;
      cursor: pointer;
    }}

    .filter-select option {{ background: #0f172a; color: #fff; }}

    .pagination-bar {{
      display: flex;
      align-items: center;
      justify-content: space-between;
      margin-top: 16px;
      font-size: 13px;
      color: var(--text-muted);
      flex-wrap: wrap;
      gap: 12px;
    }}

    .pagination-controls {{ display: flex; gap: 6px; align-items: center; }}

    .page-btn {{
      background: rgba(255, 255, 255, 0.06);
      border: 1px solid var(--border-color);
      color: #fff;
      padding: 6px 12px;
      border-radius: 6px;
      font-size: 12px;
      cursor: pointer;
      font-family: inherit;
      transition: all 0.2s;
    }}

    .page-btn:hover:not(:disabled) {{ background: rgba(255, 255, 255, 0.12); }}
    .page-btn:disabled {{ opacity: 0.35; cursor: not-allowed; }}

    .gap-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
      gap: 18px;
      margin-top: 18px;
    }}

    .gap-card {{
      padding: 20px;
      border-radius: var(--radius-sm);
      background: rgba(15, 23, 42, 0.5);
      border: 1px solid var(--border-color);
      display: flex;
      flex-direction: column;
      justify-content: space-between;
    }}

    .gap-card-header {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 12px;
    }}

    .gap-tier {{ font-weight: 700; font-size: 15px; color: #fff; }}
    .gap-count {{ font-size: 28px; font-weight: 800; color: #fff; }}
    .gap-pct {{ font-size: 13px; color: var(--text-muted); }}

    .insights-card {{
      padding: 24px 28px;
      border-left: 4px solid var(--tsel-red);
      margin-top: 24px;
    }}

    .insights-title {{
      font-size: 15px;
      font-weight: 700;
      color: #fff;
      margin-bottom: 12px;
      display: flex;
      align-items: center;
      gap: 8px;
    }}

    .insights-list {{
      list-style: none;
      display: flex;
      flex-direction: column;
      gap: 10px;
    }}

    .insights-list li {{
      position: relative;
      padding-left: 20px;
      font-size: 13px;
      color: var(--text-muted);
      line-height: 1.6;
    }}

    .insights-list li::before {{
      content: '▸';
      position: absolute;
      left: 0;
      color: var(--tsel-red);
      font-weight: bold;
    }}

    .insights-list li strong {{ color: var(--text-main); }}

    .custom-tooltip {{
      position: absolute;
      background: rgba(15, 23, 42, 0.95);
      border: 1px solid rgba(255, 255, 255, 0.15);
      padding: 10px 14px;
      border-radius: 8px;
      font-size: 12px;
      color: #fff;
      pointer-events: none;
      box-shadow: 0 10px 25px rgba(0, 0, 0, 0.5);
      z-index: 100;
      opacity: 0;
      transition: opacity 0.15s ease;
      backdrop-filter: blur(8px);
    }}

    .mini-bar-bg {{
      width: 90px;
      height: 7px;
      background: rgba(255, 255, 255, 0.08);
      border-radius: 4px;
      overflow: hidden;
      display: inline-block;
      vertical-align: middle;
      margin-left: 8px;
    }}
    .mini-bar-fill {{ height: 100%; border-radius: 4px; }}

    @media print {{
      body {{ background: #fff; color: #000; padding: 0; }}
      .glass {{ box-shadow: none; border: 1px solid #ccc; background: #fff; }}
      .btn-action, .filter-bar, .table-controls-bar, .pagination-bar {{ display: none; }}
    }}
    /* ==========================================================================
       ULTRA-SMOOTH TRANSITIONS & MICRO-ANIMATIONS
       ========================================================================== */
    html {{
      scroll-behavior: smooth;
    }}
    * {{
      -webkit-tap-highlight-color: transparent;
    }}
    
    .tab-pane, .branch-subtab-pane {{
      display: none;
      opacity: 0;
      transform: translateY(12px) scale(0.995);
      transition: opacity 0.35s cubic-bezier(0.16, 1, 0.3, 1), transform 0.35s cubic-bezier(0.16, 1, 0.3, 1);
    }}
    .tab-pane.active, .branch-subtab-pane.active {{
      display: block;
      animation: smoothTabEntrance 0.35s cubic-bezier(0.16, 1, 0.3, 1) forwards;
    }}
    @keyframes smoothTabEntrance {{
      0% {{
        opacity: 0;
        transform: translateY(14px) scale(0.995);
      }}
      100% {{
        opacity: 1;
        transform: translateY(0) scale(1);
      }}
    }}

    .filter-pill, .top-subtab-btn, .subtab-btn, .tab-btn, .btn-action, .page-btn {{
      transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1);
      user-select: none;
      -webkit-user-select: none;
    }}
    .filter-pill:active, .top-subtab-btn:active, .subtab-btn:active, .tab-btn:active, .page-btn:active {{
      transform: scale(0.95);
    }}
    .filter-pill:hover, .top-subtab-btn:hover, .subtab-btn:hover, .tab-btn:hover {{
      transform: translateY(-1px);
    }}

    .progress-bar-fill {{
      transition: width 0.7s cubic-bezier(0.16, 1, 0.3, 1);
    }}

    .data-table tbody tr {{
      transition: background 0.18s ease;
    }}

    /* Mobile Swipe Hint Banner */
    .mobile-swipe-hint {{
      display: none;
      align-items: center;
      gap: 8px;
      font-size: 11px;
      color: #38bdf8;
      background: rgba(6, 182, 212, 0.12);
      border: 1px dashed rgba(6, 182, 212, 0.35);
      padding: 8px 12px;
      border-radius: 10px;
      margin-bottom: 12px;
      font-weight: 500;
    }}

    /* Floating Action Button (FAB) for Mobile */
    .mobile-fab {{
      position: fixed;
      bottom: 24px;
      right: 20px;
      z-index: 9999;
      background: linear-gradient(135deg, var(--tsel-red) 0%, #b91c1c 100%);
      color: #fff;
      border: 1px solid rgba(255, 255, 255, 0.2);
      border-radius: 30px;
      padding: 10px 16px;
      font-size: 12px;
      font-weight: 700;
      box-shadow: 0 4px 18px rgba(255, 42, 68, 0.5);
      cursor: pointer;
      display: flex;
      align-items: center;
      gap: 6px;
      opacity: 0;
      transform: translateY(20px);
      pointer-events: none;
      transition: all 0.3s cubic-bezier(0.16, 1, 0.3, 1);
    }}
    .mobile-fab.visible {{
      opacity: 1;
      transform: translateY(0);
      pointer-events: auto;
    }}
    .mobile-fab:hover {{
      box-shadow: 0 6px 22px rgba(255, 42, 68, 0.7);
    }}

    /* ==========================================================================
       RESPONSIVE DESIGN: TABLET & SMARTPHONE (HP)
       ========================================================================== */
    @media (max-width: 1024px) {{
      .dashboard-container {{ padding: 18px 14px; }}
      .section-grid-2 {{ grid-template-columns: 1fr; gap: 16px; }}
      .kpi-grid {{ grid-template-columns: repeat(2, 1fr); gap: 14px; }}
      .chart-card {{ padding: 18px 20px; }}
      .chart-card canvas {{ height: 280px !important; }}
      .mobile-swipe-hint {{ display: flex; }}
      
      .data-table th:first-child,
      .data-table td:first-child {{
        position: sticky;
        left: 0;
        z-index: 4;
        background: #0d1627;
        box-shadow: 3px 0 8px rgba(0, 0, 0, 0.4);
      }}
    }}

    @media (max-width: 768px) {{
      .dashboard-container {{ padding: 10px 8px; }}
      
      /* Mobile Header */
      .header {{
        flex-direction: column;
        align-items: flex-start;
        padding: 16px 14px;
        gap: 12px;
        border-radius: 16px;
      }}
      .logo-icon {{ width: 38px; height: 38px; font-size: 18px; }}
      .header-title h1 {{ font-size: 17px; line-height: 1.3; }}
      .header-title p {{ font-size: 11px; flex-wrap: wrap; gap: 6px; }}
      .header-actions {{
        width: 100%;
        display: flex;
        justify-content: space-between;
        align-items: center;
      }}
      .meta-tag {{ font-size: 10.5px; padding: 4px 10px; }}
      .btn-action {{ font-size: 11px; padding: 6px 12px; }}

      /* Mobile Branch Focus Filter Bar */
      .filter-bar {{
        padding: 12px 14px;
        flex-direction: column;
        align-items: stretch;
        gap: 10px;
        border-radius: 16px;
      }}
      .filter-group {{
        display: flex;
        align-items: center;
        gap: 6px;
        overflow-x: auto;
        white-space: nowrap;
        -webkit-overflow-scrolling: touch;
        scrollbar-width: none;
        padding-bottom: 4px;
      }}
      .filter-group::-webkit-scrollbar {{ display: none; }}
      .filter-label {{ font-size: 11px; margin-right: 2px; }}
      .filter-pill {{
        padding: 7px 12px;
        font-size: 11px;
        white-space: nowrap;
        min-height: 38px;
        display: inline-flex;
        align-items: center;
      }}

      /* Mobile Tab Switcher directly in Branch Focus */
      .branch-tab-switcher {{
        width: 100%;
        border-left: none;
        padding-left: 0;
        border-top: 1px solid rgba(255, 255, 255, 0.08);
        padding-top: 10px;
        display: flex;
        gap: 6px;
      }}
      .top-subtab-btn {{
        flex: 1;
        justify-content: center;
        text-align: center;
        font-size: 11px;
        padding: 8px 6px;
        min-height: 38px;
        white-space: normal;
        line-height: 1.2;
      }}

      /* Mobile KPI Grid (2x2) */
      .kpi-grid {{
        grid-template-columns: 1fr 1fr;
        gap: 8px;
        margin-bottom: 16px;
      }}
      .kpi-card {{
        padding: 12px 12px;
        border-radius: 14px;
      }}
      .kpi-title {{ font-size: 10.5px; }}
      .kpi-icon {{ width: 26px; height: 26px; font-size: 12px; }}
      .kpi-val {{ font-size: 19px; margin-bottom: 4px; }}
      .kpi-val .unit {{ font-size: 10px; }}
      .kpi-sub {{ font-size: 10px; line-height: 1.3; }}
      .progress-bar-container {{ height: 4px; margin-top: 8px; }}

      /* Mobile Charts */
      .chart-card {{ padding: 14px 12px; border-radius: 16px; }}
      .card-title {{ font-size: 13.5px; }}
      .chart-card canvas {{ height: 240px !important; }}
      .chart-legend {{ gap: 8px; }}
      .legend-item {{ font-size: 10.5px; }}

      /* Mobile Navigation Tabs Bar */
      .tabs-wrapper {{ padding: 14px 10px !important; border-radius: 16px; }}
      .tabs-nav {{
        display: flex;
        flex-wrap: nowrap;
        overflow-x: auto;
        -webkit-overflow-scrolling: touch;
        scrollbar-width: none;
        gap: 6px;
        padding-bottom: 8px;
      }}
      .tabs-nav::-webkit-scrollbar {{ display: none; }}
      .tab-btn {{
        white-space: nowrap;
        font-size: 11px;
        padding: 8px 12px;
        min-height: 40px;
      }}

      /* Mobile Subtabs inside Branch Focus */
      .subtabs-bar {{
        width: 100%;
        display: flex;
        gap: 4px;
        padding: 4px;
        border-radius: 12px;
      }}
      .subtab-btn {{
        flex: 1;
        justify-content: center;
        text-align: center;
        font-size: 11px;
        padding: 8px 6px;
        min-height: 38px;
      }}

      /* Mobile Controls inside Tables */
      .table-controls-bar {{
        flex-direction: column;
        align-items: stretch;
        gap: 10px;
      }}
      .filter-group {{ width: 100%; }}
      .filter-select {{ width: 100%; min-height: 40px; font-size: 12px; }}
      .search-input-wrapper {{ max-width: 100% !important; width: 100%; }}
      .search-input {{ min-height: 40px; font-size: 12px; }}

      /* Mobile Data Tables */
      .table-wrapper {{
        border-radius: 10px;
        max-height: 520px;
        overflow-y: auto;
      }}
      .data-table th {{
        font-size: 10.5px;
        padding: 8px 8px;
        white-space: nowrap;
      }}
      .data-table td {{
        font-size: 10.5px;
        padding: 7px 8px;
      }}
      .data-table th:first-child,
      .data-table td:first-child {{
        position: sticky;
        left: 0;
        z-index: 4;
        background: #0a111e;
        box-shadow: 3px 0 8px rgba(0,0,0,0.5);
      }}

      /* Mobile Pagination */
      .pagination-bar {{
        flex-direction: column;
        gap: 10px;
        align-items: center;
      }}
      .pagination-bar .page-info {{ font-size: 11px; }}
      .page-btn {{ padding: 8px 14px; font-size: 12px; min-height: 38px; }}
    }}

    /* Live Sync Badge & Toast Notification */
    .sync-status-badge {{
      display: inline-flex;
      align-items: center;
      gap: 6px;
      background: rgba(16, 185, 129, 0.12);
      border: 1px solid rgba(16, 185, 129, 0.3);
      color: #34d399;
      padding: 6px 12px;
      border-radius: 20px;
      font-size: 11.5px;
      font-weight: 600;
    }}
    .spin {{
      animation: spinAnim 0.7s linear infinite;
    }}
    @keyframes spinAnim {{
      100% {{ transform: rotate(360deg); }}
    }}
    .toast-notification {{
      position: fixed;
      top: 24px;
      right: 24px;
      z-index: 99999;
      background: rgba(15, 23, 42, 0.95);
      border: 1px solid rgba(16, 185, 129, 0.4);
      color: #fff;
      padding: 12px 18px;
      border-radius: 12px;
      box-shadow: 0 10px 30px rgba(0,0,0,0.5);
      backdrop-filter: blur(12px);
      display: flex;
      align-items: center;
      gap: 10px;
      font-size: 12.5px;
      font-weight: 600;
      opacity: 0;
      transform: translateY(-20px);
      transition: all 0.3s cubic-bezier(0.16, 1, 0.3, 1);
      pointer-events: none;
    }}
    .toast-notification.show {{
      opacity: 1;
      transform: translateY(0);
      pointer-events: auto;
    }}
    /* Progress Action Bar Chart Section */
    .prog-analytics-card {{
      margin-bottom: 22px;
      padding: 22px 24px;
      border: 1px solid rgba(255, 255, 255, 0.08);
      border-radius: var(--radius-md);
      background: linear-gradient(145deg, rgba(30, 41, 59, 0.75), rgba(15, 23, 42, 0.9));
      backdrop-filter: blur(16px);
      box-shadow: 0 10px 30px rgba(0, 0, 0, 0.35);
    }}

    .prog-analytics-header {{
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      margin-bottom: 18px;
      gap: 16px;
      flex-wrap: wrap;
    }}

    .prog-kpi-chips {{
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
      align-items: center;
    }}

    .prog-chip {{
      background: rgba(255, 255, 255, 0.04);
      border: 1px solid rgba(255, 255, 255, 0.08);
      padding: 6px 14px;
      border-radius: 8px;
      font-size: 12px;
      display: flex;
      align-items: center;
      gap: 8px;
    }}

    .prog-chip strong {{
      font-family: 'JetBrains Mono', monospace;
      color: #fff;
    }}

    .prog-charts-grid {{
      display: grid;
      grid-template-columns: 1.15fr 1fr;
      gap: 20px;
    }}

    @media (max-width: 992px) {{
      .prog-charts-grid {{
        grid-template-columns: 1fr;
      }}
    }}

    .prog-chart-box {{
      background: rgba(15, 23, 42, 0.55);
      border: 1px solid rgba(255, 255, 255, 0.05);
      border-radius: 12px;
      padding: 18px 20px;
      display: flex;
      flex-direction: column;
      gap: 12px;
    }}

    .prog-chart-title {{
      font-size: 13.5px;
      font-weight: 700;
      color: #e2e8f0;
      display: flex;
      align-items: center;
      justify-content: space-between;
      margin-bottom: 4px;
    }}

    /* Horizontal Bar Row */
    .hbar-item {{
      display: flex;
      flex-direction: column;
      gap: 4px;
    }}

    .hbar-header {{
      display: flex;
      justify-content: space-between;
      font-size: 12px;
      color: var(--text-muted);
    }}

    .hbar-label {{
      font-weight: 600;
      display: flex;
      align-items: center;
      gap: 6px;
    }}

    .hbar-value {{
      font-family: 'JetBrains Mono', monospace;
      font-weight: 700;
      color: #fff;
    }}

    .hbar-track {{
      width: 100%;
      height: 8px;
      background: rgba(255, 255, 255, 0.06);
      border-radius: 6px;
      overflow: hidden;
      position: relative;
    }}

    .hbar-fill {{
      height: 100%;
      border-radius: 6px;
      transition: width 0.7s cubic-bezier(0.4, 0, 0.2, 1);
    }}

    /* Branch Comparison Bar Chart */
    .branch-bar-item {{
      display: flex;
      align-items: center;
      gap: 12px;
      padding: 8px 12px;
      border-radius: 8px;
      transition: all 0.2s ease;
      cursor: pointer;
      background: rgba(255, 255, 255, 0.02);
      border: 1px solid rgba(255, 255, 255, 0.04);
    }}

    .branch-bar-item:hover {{
      background: rgba(255, 255, 255, 0.06);
      border-color: rgba(255, 255, 255, 0.12);
      transform: translateX(2px);
    }}

    .branch-bar-name {{
      width: 82px;
      font-size: 12px;
      font-weight: 700;
      color: #fff;
      display: flex;
      align-items: center;
      gap: 6px;
    }}

    .branch-bar-track-wrap {{
      flex: 1;
      display: flex;
      flex-direction: column;
      gap: 3px;
    }}

    .branch-bar-track {{
      width: 100%;
      height: 9px;
      background: rgba(255, 255, 255, 0.06);
      border-radius: 6px;
      overflow: hidden;
      display: flex;
    }}

    .branch-bar-filled {{
      height: 100%;
      border-radius: 6px;
      transition: width 0.7s cubic-bezier(0.4, 0, 0.2, 1);
    }}

    .branch-bar-meta {{
      display: flex;
      justify-content: space-between;
      font-size: 11px;
      color: var(--text-dim);
      font-family: 'JetBrains Mono', monospace;
    }}

    .branch-bar-pct {{
      width: 52px;
      text-align: right;
      font-size: 12px;
      font-weight: 700;
      font-family: 'JetBrains Mono', monospace;
    }}

    /* Priority Analytics & Badges */
    .priority-analytics-card {{
      background: linear-gradient(135deg, rgba(30, 41, 59, 0.75), rgba(15, 23, 42, 0.85));
      border: 1px solid rgba(245, 158, 11, 0.3);
      border-radius: 16px;
      padding: 22px 24px;
      margin-bottom: 24px;
      box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.3);
      position: relative;
      overflow: hidden;
    }}
    .priority-analytics-card::before {{
      content: '';
      position: absolute;
      top: 0;
      left: 0;
      width: 100%;
      height: 3px;
      background: linear-gradient(90deg, #f59e0b, #38bdf8, #ef4444, #10b981);
    }}
    .priority-kpi-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(210px, 1fr));
      gap: 14px;
      margin-top: 16px;
      margin-bottom: 20px;
    }}
    .priority-kpi-item {{
      background: rgba(15, 23, 42, 0.65);
      border: 1px solid rgba(255, 255, 255, 0.08);
      border-radius: 12px;
      padding: 14px 16px;
      transition: all 0.2s ease;
      cursor: pointer;
    }}
    .priority-kpi-item:hover {{
      border-color: rgba(245, 158, 11, 0.45);
      background: rgba(30, 41, 59, 0.75);
      transform: translateY(-2px);
      box-shadow: 0 8px 16px rgba(0,0,0,0.25);
    }}
    .priority-filter-bar {{
      display: flex;
      align-items: center;
      gap: 8px;
      flex-wrap: wrap;
      margin-bottom: 12px;
      padding: 10px 14px;
      background: rgba(15, 23, 42, 0.7);
      border: 1px solid rgba(255, 255, 255, 0.08);
      border-radius: 12px;
    }}
    .priority-pill {{
      padding: 6px 14px;
      border-radius: 20px;
      font-size: 12px;
      font-weight: 600;
      cursor: pointer;
      background: rgba(255, 255, 255, 0.05);
      color: var(--text-muted);
      border: 1px solid rgba(255, 255, 255, 0.12);
      transition: all 0.2s ease;
      display: inline-flex;
      align-items: center;
      gap: 6px;
    }}
    .priority-pill:hover {{
      background: rgba(255, 255, 255, 0.12);
      color: #fff;
      border-color: rgba(255, 255, 255, 0.25);
      transform: translateY(-1px);
    }}
    .priority-pill.active {{
      background: linear-gradient(135deg, #f59e0b, #d97706);
      color: #0f172a;
      font-weight: 700;
      border-color: #f59e0b;
      box-shadow: 0 4px 12px rgba(245, 158, 11, 0.35);
    }}
    .prio-bar-wrapper {{
      background: rgba(15, 23, 42, 0.5);
      border: 1px solid rgba(255, 255, 255, 0.06);
      border-radius: 10px;
      padding: 12px 14px;
      transition: all 0.2s ease;
    }}
    .prio-bar-wrapper:hover {{
      background: rgba(30, 41, 59, 0.6);
      border-color: rgba(255, 255, 255, 0.15);
      transform: translateX(2px);
    }}

    /* Theme Toggle Button */
    .btn-theme-toggle {{
      background: rgba(255, 255, 255, 0.1);
      border-color: rgba(255, 255, 255, 0.22);
      color: #f8fafc;
      font-weight: 700;
    }}
    .btn-theme-toggle:hover {{
      background: rgba(255, 255, 255, 0.18);
      border-color: rgba(255, 255, 255, 0.35);
      color: #ffffff;
    }}

    /* =========================================================================
       VERSI PUTIH (LIGHT THEME) OVERRIDES
       ========================================================================= */
    html[data-theme="light"] {{
      --bg-primary: #f8fafc;
      --bg-secondary: #f1f5f9;
      --bg-card: #ffffff;
      --bg-card-hover: #f8fafc;
      --border-color: #e2e8f0;
      --border-accent: rgba(225, 29, 72, 0.35);
      --text-main: #0f172a;
      --text-muted: #475569;
      --text-dim: #64748b;
      --tsel-red: #e11d48;
      --tsel-red-glow: rgba(225, 29, 72, 0.15);
      --shadow-lg: 0 10px 25px -5px rgba(15, 23, 42, 0.08), 0 4px 10px -4px rgba(15, 23, 42, 0.04);
    }}
    html[data-theme="light"] body {{
      background: #f1f5f9;
      color: var(--text-main);
    }}
    html[data-theme="light"] .glass {{
      background: #ffffff;
      backdrop-filter: none;
      -webkit-backdrop-filter: none;
      border: 1px solid #e2e8f0;
      box-shadow: var(--shadow-lg);
    }}
    html[data-theme="light"] .brand-logo {{
      box-shadow: 0 4px 12px rgba(225, 29, 72, 0.14), 0 0 0 1px rgba(15, 23, 42, 0.08);
    }}
    html[data-theme="light"] .brand-text h1 {{
      background: none;
      -webkit-background-clip: unset;
      -webkit-text-fill-color: #0f172a;
      color: #0f172a;
    }}
    html[data-theme="light"] .badge-live {{
      background: #ecfdf5;
      border-color: #a7f3d0;
      color: #047857;
    }}
    html[data-theme="light"] .sync-status-badge {{
      background: #ecfdf5;
      border-color: #a7f3d0;
      color: #047857;
    }}
    html[data-theme="light"] #sync-status-label {{
      color: #334155;
    }}
    html[data-theme="light"] #refresh-timer {{
      color: #0284c7 !important;
    }}
    html[data-theme="light"] .meta-tag {{
      background: #f8fafc;
      border-color: #cbd5e1;
      color: #334155;
    }}
    html[data-theme="light"] .btn-action {{
      background: #ffffff;
      border-color: #cbd5e1;
      color: #1e293b;
      box-shadow: 0 1px 2px rgba(15, 23, 42, 0.05);
    }}
    html[data-theme="light"] .btn-action:hover {{
      background: #f8fafc;
      border-color: #94a3b8;
      color: #0f172a;
    }}
    html[data-theme="light"] .btn-theme-toggle {{
      background: #0f172a;
      border-color: #0f172a;
      color: #ffffff;
    }}
    html[data-theme="light"] .btn-theme-toggle:hover {{
      background: #1e293b;
      border-color: #1e293b;
      color: #ffffff;
    }}
    html[data-theme="light"] .btn-action[style*="background:rgba(255, 42, 68"],
    html[data-theme="light"] .btn-action[style*="background:rgba(255,42,68"] {{
      background: #ffe4e6 !important;
      border-color: #fda4af !important;
      color: #be123c !important;
    }}
    html[data-theme="light"] .filter-pill {{
      background: #f8fafc;
      border-color: #cbd5e1;
      color: #334155;
    }}
    html[data-theme="light"] .filter-pill:hover {{
      background: #e2e8f0;
      color: #0f172a;
    }}
    html[data-theme="light"] .filter-pill.active {{
      background: linear-gradient(135deg, #e11d48, #be123c);
      border-color: transparent;
      color: #ffffff;
    }}
    html[data-theme="light"] .branch-tab-switcher {{
      border-left-color: #cbd5e1;
    }}
    html[data-theme="light"] .top-subtab-btn {{
      background: #f8fafc;
      border-color: #cbd5e1;
      color: #334155;
    }}
    html[data-theme="light"] .top-subtab-btn:hover {{
      background: #e2e8f0;
      border-color: #94a3b8;
      color: #0f172a;
    }}
    html[data-theme="light"] .top-subtab-btn.active {{
      background: linear-gradient(135deg, #e11d48, #be123c);
      border-color: #e11d48;
      color: #ffffff;
    }}
    html[data-theme="light"] .target-badge {{
      background: #fef3c7;
      border-color: #fcd34d;
      color: #b45309;
    }}
    html[data-theme="light"] .target-badge[style*="color:#34d399"] {{
      background: #d1fae5 !important;
      border-color: #6ee7b7 !important;
      color: #047857 !important;
    }}
    html[data-theme="light"] .kpi-card:hover {{
      border-color: #94a3b8;
    }}
    html[data-theme="light"] .kpi-val {{
      color: #0f172a;
    }}
    html[data-theme="light"] .kpi-sub .trend-up {{
      color: #059669;
    }}
    html[data-theme="light"] .kpi-sub .trend-down {{
      color: #dc2626;
    }}
    html[data-theme="light"] .progress-bar-container {{
      background: #e2e8f0;
    }}
    html[data-theme="light"] .card-title,
    html[data-theme="light"] .insights-title {{
      color: #0f172a;
    }}
    html[data-theme="light"] .grid-line {{
      stroke: rgba(15, 23, 42, 0.08);
    }}
    html[data-theme="light"] .axis-label {{
      fill: #475569;
    }}
    html[data-theme="light"] .legend-color[style*="background:#38bdf8"] {{
      background: #0284c7 !important;
    }}
    html[data-theme="light"] .legend-color[style*="background:#f59e0b"] {{
      background: #d97706 !important;
    }}
    html[data-theme="light"] .legend-color[style*="background:#a855f7"] {{
      background: #7c3aed !important;
    }}
    html[data-theme="light"] .legend-color[style*="dashed #fff"] {{
      border-color: #0f172a !important;
    }}
    html[data-theme="light"] #side-card-alert {{
      background: #fff1f2 !important;
      border-color: #fecdd3 !important;
    }}
    html[data-theme="light"] #side-alert-title {{
      color: #be123c !important;
    }}
    html[data-theme="light"] #side-alert-desc {{
      color: #334155 !important;
    }}
    html[data-theme="light"] .subtabs-bar {{
      background: #f1f5f9;
      border-color: #cbd5e1;
    }}
    html[data-theme="light"] .subtab-btn {{
      background: #ffffff;
      border-color: #cbd5e1;
      color: #334155;
    }}
    html[data-theme="light"] .subtab-btn:hover {{
      background: #e2e8f0;
      color: #0f172a;
    }}
    html[data-theme="light"] .subtab-btn.active {{
      background: linear-gradient(135deg, #e11d48, #be123c);
      border-color: transparent;
      color: #ffffff;
    }}
    html[data-theme="light"] .tabs-nav {{
      background: #f1f5f9;
      border-color: #cbd5e1;
    }}
    html[data-theme="light"] .tab-btn {{
      color: #475569;
    }}
    html[data-theme="light"] .tab-btn:hover {{
      color: #0f172a;
      background: #e2e8f0;
    }}
    html[data-theme="light"] .tab-btn.active {{
      background: linear-gradient(135deg, #f59e0b, #d97706);
      color: #ffffff;
      box-shadow: 0 4px 12px rgba(217, 119, 6, 0.25);
    }}
    html[data-theme="light"] .tab-btn[onclick*="tab-aug-db"]:not(.active) {{
      background: #ecfeff !important;
      color: #0e7490 !important;
      border-color: #67e8f9 !important;
    }}
    html[data-theme="light"] #kab-drilldown-box {{
      background: #f8fafc !important;
      border-color: #fda4af !important;
    }}
    html[data-theme="light"] .table-container {{
      border-color: #cbd5e1;
      background: #ffffff;
    }}
    html[data-theme="light"] .table-controls-bar {{
      background: #f8fafc;
      border-color: #e2e8f0;
    }}
    html[data-theme="light"] .search-input,
    html[data-theme="light"] .filter-select {{
      background: #ffffff;
      border-color: #cbd5e1;
      color: #0f172a;
    }}
    html[data-theme="light"] .search-input::placeholder {{
      color: #64748b;
    }}
    html[data-theme="light"] .filter-select option {{
      background: #ffffff;
      color: #0f172a;
    }}
    html[data-theme="light"] select.filter-select[style*="color:#fcd34d"] {{
      background: #fffbeb !important;
      border-color: #f59e0b !important;
      color: #b45309 !important;
    }}
    html[data-theme="light"] select.filter-select[style*="color:#7dd3fc"] {{
      background: #f0f9ff !important;
      border-color: #0284c7 !important;
      color: #0369a1 !important;
    }}
    html[data-theme="light"] select.filter-select[style*="color:#38bdf8"] {{
      background: #e0f2fe !important;
      border-color: #0284c7 !important;
      color: #0369a1 !important;
    }}
    html[data-theme="light"] select.filter-select[style*="color:#d8b4fe"] {{
      background: #f5f3ff !important;
      border-color: #7c3aed !important;
      color: #6d28d9 !important;
    }}
    html[data-theme="light"] select.filter-select[style*="color:#34d399"] {{
      background: #ecfdf5 !important;
      border-color: #059669 !important;
      color: #047857 !important;
    }}
    html[data-theme="light"] select.filter-select[style*="color:#67e8f9"] {{
      background: #ecfeff !important;
      border-color: #0891b2 !important;
      color: #0e7490 !important;
    }}
    html[data-theme="light"] select.filter-select[style*="color:#fde047"] {{
      background: #fefce8 !important;
      border-color: #ca8a04 !important;
      color: #a16207 !important;
    }}
    html[data-theme="light"] select.filter-select[style*="background:rgba(239,68,68"] {{
      background: #fff1f2 !important;
      border-color: #e11d48 !important;
      color: #be123c !important;
    }}

    /* Table headers, rows, sticky & tinted headers */
    html[data-theme="light"] table.data-table th {{
      background: #f1f5f9;
      color: #334155;
      border-bottom: 1px solid #cbd5e1;
    }}
    html[data-theme="light"] table.data-table th.new-col-th {{
      background: #ecfeff;
      border-bottom: 2px solid #0891b2;
      color: #0e7490;
    }}
    html[data-theme="light"] table.data-table th[style*="rgba(239,68,68"] {{
      background: #fee2e2 !important;
      border-bottom: 2px solid #dc2626 !important;
      color: #b91c1c !important;
    }}
    html[data-theme="light"] table.data-table th[style*="rgba(245,158,11"] {{
      background: #fef3c7 !important;
      border-bottom: 2px solid #d97706 !important;
      color: #b45309 !important;
    }}
    html[data-theme="light"] table.data-table th[style*="rgba(234,179,8"] {{
      background: #fef9c3 !important;
      border-bottom: 2px solid #ca8a04 !important;
      color: #854d0e !important;
    }}
    html[data-theme="light"] table.data-table th[style*="rgba(16,185,129"] {{
      background: #d1fae5 !important;
      border-bottom: 2px solid #059669 !important;
      color: #047857 !important;
    }}
    html[data-theme="light"] table.data-table th[style*="rgba(167,139,250"] {{
      background: #ede9fe !important;
      border-bottom: 2px solid #7c3aed !important;
      color: #6d28d9 !important;
    }}
    html[data-theme="light"] table.data-table th[style*="rgba(56,189,248"] {{
      background: #e0f2fe !important;
      border-bottom: 2px solid #0284c7 !important;
      color: #0369a1 !important;
    }}
    html[data-theme="light"] table.data-table th[style*="rgba(59,130,246"] {{
      background: #dbeafe !important;
      border-bottom: 2px solid #2563eb !important;
      color: #1d4ed8 !important;
    }}
    html[data-theme="light"] table.data-table th[style*="rgba(14,165,233"] {{
      background: #e0f2fe !important;
      border-bottom: 2px solid #0284c7 !important;
      color: #0369a1 !important;
    }}
    html[data-theme="light"] table.data-table th[style*="rgba(6,182,212"] {{
      background: #cffafe !important;
      border-bottom: 2px solid #0891b2 !important;
      color: #0e7490 !important;
    }}
    html[data-theme="light"] table.data-table td {{
      color: #0f172a;
      border-bottom: 1px solid #e2e8f0;
    }}
    html[data-theme="light"] table.data-table td.new-col-td {{
      background: #f8fafc;
    }}
    html[data-theme="light"] table.data-table tbody tr:hover {{
      background: #f1f5f9;
    }}
    html[data-theme="light"] table.data-table tr.total-row {{
      background: #fff1f2;
      border-top: 2px solid #fda4af;
    }}

    /* Badges */
    html[data-theme="light"] .badge-success {{ background: #dcfce7; color: #15803d; border-color: #86efac; }}
    html[data-theme="light"] .badge-danger {{ background: #fee2e2; color: #b91c1c; border-color: #fca5a5; }}
    html[data-theme="light"] .badge-warning {{ background: #fef3c7; color: #b45309; border-color: #fcd34d; }}
    html[data-theme="light"] .badge-info {{ background: #e0f2fe; color: #0369a1; border-color: #7dd3fc; }}
    html[data-theme="light"] .badge-purple {{ background: #f3e8ff; color: #7e22ce; border-color: #d8b4fe; }}
    html[data-theme="light"] .badge-diamond {{ background: #e0f2fe; color: #0369a1; border-color: #0284c7; }}
    html[data-theme="light"] .badge-platinum {{ background: #f3e8ff; color: #6d28d9; border-color: #7c3aed; }}
    html[data-theme="light"] .badge-gold {{ background: #fef3c7; color: #b45309; border-color: #d97706; }}
    html[data-theme="light"] .badge-silver {{ background: #f1f5f9; color: #334155; border-color: #94a3b8; }}
    html[data-theme="light"] .badge-bronze {{ background: #ffedd5; color: #9a3412; border-color: #ea580c; }}

    /* Priority Analytics, Program Analytics, Gap Cards, Insights */
    html[data-theme="light"] .priority-analytics-card {{
      background: linear-gradient(135deg, #fffbeb 0%, #ffffff 100%);
      border-color: #fcd34d;
      box-shadow: var(--shadow-lg);
    }}
    html[data-theme="light"] .priority-kpi-item {{
      background: #ffffff;
      border-color: #e2e8f0;
    }}
    html[data-theme="light"] .priority-kpi-item:hover {{
      background: #f8fafc;
      border-color: #f59e0b;
      box-shadow: 0 6px 14px rgba(15, 23, 42, 0.06);
    }}
    html[data-theme="light"] .priority-filter-bar {{
      background: #f8fafc;
      border-color: #e2e8f0;
    }}
    html[data-theme="light"] .priority-pill {{
      background: #ffffff;
      border-color: #cbd5e1;
      color: #334155;
    }}
    html[data-theme="light"] .priority-pill:hover {{
      background: #f1f5f9;
      border-color: #94a3b8;
      color: #0f172a;
    }}
    html[data-theme="light"] .priority-pill.active {{
      background: linear-gradient(135deg, #f59e0b, #d97706);
      color: #ffffff;
      border-color: #d97706;
    }}
    html[data-theme="light"] .prio-bar-wrapper {{
      background: #ffffff;
      border-color: #e2e8f0;
    }}
    html[data-theme="light"] .prio-bar-wrapper:hover {{
      background: #f8fafc;
      border-color: #cbd5e1;
    }}
    html[data-theme="light"] .prog-analytics-card {{
      background: linear-gradient(145deg, #ffffff, #f8fafc);
      border-color: #cbd5e1;
      box-shadow: var(--shadow-lg);
    }}
    html[data-theme="light"] .prog-chip {{
      background: #f8fafc;
      border-color: #cbd5e1;
      color: #334155;
    }}
    html[data-theme="light"] .prog-chip span,
    html[data-theme="light"] .prog-chip strong {{
      color: #0f172a;
    }}
    html[data-theme="light"] .prog-chart-box {{
      background: #ffffff;
      border-color: #e2e8f0;
    }}
    html[data-theme="light"] .prog-chart-title {{
      color: #0f172a;
    }}
    html[data-theme="light"] .hbar-item {{
      background: #f8fafc;
      border-color: #e2e8f0;
    }}
    html[data-theme="light"] .hbar-item:hover {{
      background: #f1f5f9;
      border-color: #cbd5e1;
    }}
    html[data-theme="light"] .hbar-label {{ color: #1e293b; }}
    html[data-theme="light"] .hbar-track {{ background: #e2e8f0; }}
    html[data-theme="light"] .hbar-value {{ color: #0f172a; }}
    html[data-theme="light"] .branch-bar-name {{ color: #0f172a; }}
    html[data-theme="light"] .branch-bar-track {{ background: #e2e8f0; }}
    html[data-theme="light"] .branch-bar-meta {{ color: #475569; }}
    html[data-theme="light"] .gap-card {{
      background: #f8fafc;
      border-color: #e2e8f0;
    }}
    html[data-theme="light"] .gap-tier,
    html[data-theme="light"] .gap-count {{
      color: #0f172a;
    }}
    html[data-theme="light"] .insights-card {{
      background: #fff1f2;
      border-color: #fecdd3;
      border-left: 4px solid #e11d48;
    }}
    html[data-theme="light"] .insights-list {{ color: #334155; }}
    html[data-theme="light"] .insights-list li strong {{ color: #0f172a; }}
    html[data-theme="light"] .page-btn {{
      background: #ffffff;
      border-color: #cbd5e1;
      color: #0f172a;
    }}
    html[data-theme="light"] .page-btn:hover:not(:disabled) {{
      background: #f1f5f9;
    }}
    html[data-theme="light"] .mini-bar-bg {{
      background: #e2e8f0;
    }}
    html[data-theme="light"] .custom-tooltip {{
      background: #ffffff;
      border-color: #cbd5e1;
      color: #0f172a;
      box-shadow: 0 10px 25px rgba(15, 23, 42, 0.14);
    }}
    html[data-theme="light"] .toast-notification {{
      background: #ffffff;
      border-color: #10b981;
      color: #0f172a;
      box-shadow: 0 10px 25px rgba(15, 23, 42, 0.14);
    }}
    html[data-theme="light"] .mobile-swipe-hint {{
      color: #0369a1;
      background: #e0f2fe;
      border-color: #7dd3fc;
    }}

    /* High-contrast light overrides for dynamically generated / inline pastel colors */
    html[data-theme="light"] [style*="color:#fff"]:not([style*="linear-gradient"]),
    html[data-theme="light"] [style*="color: #fff"]:not([style*="linear-gradient"]) {{
      color: #0f172a !important;
    }}
    html[data-theme="light"] [style*="color:#fca5a5"],
    html[data-theme="light"] [style*="color:#f87171"],
    html[data-theme="light"] [style*="color:#ff6b7d"] {{
      color: #b91c1c !important;
    }}
    html[data-theme="light"] [style*="color:#fcd34d"],
    html[data-theme="light"] [style*="color:#fbbf24"],
    html[data-theme="light"] [style*="color:#fde047"],
    html[data-theme="light"] [style*="color:#eab308"] {{
      color: #b45309 !important;
    }}
    html[data-theme="light"] [style*="color:#6ee7b7"],
    html[data-theme="light"] [style*="color:#34d399"] {{
      color: #047857 !important;
    }}
    html[data-theme="light"] [style*="color:#7dd3fc"],
    html[data-theme="light"] [style*="color:#38bdf8"] {{
      color: #0369a1 !important;
    }}
    html[data-theme="light"] [style*="color:#93c5fd"],
    html[data-theme="light"] [style*="color:#60a5fa"] {{
      color: #1d4ed8 !important;
    }}
    html[data-theme="light"] [style*="color:#67e8f9"],
    html[data-theme="light"] [style*="color:#22d3ee"],
    html[data-theme="light"] [style*="color:#5eead4"] {{
      color: #0e7490 !important;
    }}
    html[data-theme="light"] [style*="color:#c4b5fd"],
    html[data-theme="light"] [style*="color:#d8b4fe"],
    html[data-theme="light"] [style*="color:#c084fc"],
    html[data-theme="light"] [style*="color:#a78bfa"] {{
      color: #6d28d9 !important;
    }}
    html[data-theme="light"] [style*="color:#f472b6"] {{
      color: #be185d !important;
    }}
    html[data-theme="light"] [style*="color:#fb923c"] {{
      color: #c2410c !important;
    }}
    html[data-theme="light"] [style*="background:rgba(255,255,255,0.06)"],
    html[data-theme="light"] [style*="background:rgba(255,255,255,0.08)"] {{
      background: #e2e8f0 !important;
    }}
    html[data-theme="light"] [style*="background:rgba(255,255,255,0.03)"] {{
      background: #f8fafc !important;
      border-color: #e2e8f0 !important;
    }}
    html[data-theme="light"] [style*="rgba(255,255,255,0.06)"][style*="color:#94a3b8"],
    html[data-theme="light"] [style*="rgba(255,255,255,0.06)"][style*="var(--text-muted)"] {{
      color: #334155 !important;
    }}
    html[data-theme="light"] [style*="border-top:1px dashed rgba(255,255,255"] {{
      border-top-color: #cbd5e1 !important;
    }}
    html[data-theme="light"] [style*="background:rgba(239,68,68,0.25)"] {{
      background: #fee2e2 !important;
      border-color: #fca5a5 !important;
    }}
    html[data-theme="light"] [style*="background:rgba(245,158,11,0.25)"] {{
      background: #fef3c7 !important;
      border-color: #fcd34d !important;
    }}
    html[data-theme="light"] [style*="background:rgba(168,85,247,0.25)"] {{
      background: #f3e8ff !important;
      border-color: #d8b4fe !important;
    }}
    html[data-theme="light"] [style*="background:rgba(56,189,248,0.25)"] {{
      background: #e0f2fe !important;
      border-color: #7dd3fc !important;
    }}
    html[data-theme="light"] [style*="background:rgba(6,182,212,0.15)"] {{
      background: #ecfeff !important;
      border-color: #67e8f9 !important;
    }}

    @media (max-width: 1024px) {{
      html[data-theme="light"] .data-table th:first-child {{
        background: #f1f5f9 !important;
        box-shadow: 3px 0 8px rgba(15, 23, 42, 0.08);
      }}
      html[data-theme="light"] .data-table td:first-child {{
        background: #ffffff !important;
        box-shadow: 3px 0 8px rgba(15, 23, 42, 0.08);
      }}
    }}
    @media (max-width: 768px) {{
      html[data-theme="light"] .branch-tab-switcher {{
        border-top-color: #e2e8f0;
      }}
      html[data-theme="light"] .data-table th:first-child {{
        background: #f1f5f9 !important;
        box-shadow: 3px 0 8px rgba(15, 23, 42, 0.08);
      }}
      html[data-theme="light"] .data-table td:first-child {{
        background: #ffffff !important;
        box-shadow: 3px 0 8px rgba(15, 23, 42, 0.08);
      }}
    }}
    @media print {{
      html[data-theme="light"] body {{
        background: #ffffff !important;
      }}
    }}

    /* Period Switcher (Agustus 2026 <-> September 2026) */
    .period-switcher-group {{
      display: inline-flex;
      align-items: center;
      background: rgba(15, 23, 42, 0.8);
      border: 1px solid rgba(56, 189, 248, 0.35);
      border-radius: 30px;
      padding: 4px;
      gap: 4px;
      box-shadow: 0 4px 14px rgba(0, 0, 0, 0.3);
    }}
    .period-switcher-label {{
      font-size: 11px;
      font-weight: 700;
      color: var(--text-muted);
      padding: 0 10px 0 12px;
      text-transform: uppercase;
      letter-spacing: 0.6px;
      white-space: nowrap;
    }}
    .period-switch-btn {{
      background: transparent;
      border: 1px solid transparent;
      color: var(--text-muted);
      padding: 6px 14px;
      border-radius: 24px;
      font-size: 12.5px;
      font-weight: 700;
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      gap: 7px;
      transition: all 0.22s ease;
      white-space: nowrap;
    }}
    .period-switch-btn:hover {{
      color: #fff;
      background: rgba(255, 255, 255, 0.07);
    }}
    .period-switch-btn.active {{
      background: linear-gradient(135deg, #0ea5e9, #2563eb);
      color: #ffffff;
      border-color: rgba(125, 211, 252, 0.5);
      box-shadow: 0 4px 14px rgba(14, 165, 233, 0.4);
    }}
    .period-pill-count {{
      font-size: 10.5px;
      font-family: 'JetBrains Mono', monospace;
      padding: 2px 7px;
      border-radius: 12px;
      background: rgba(255, 255, 255, 0.14);
      color: #e2e8f0;
      font-weight: 700;
    }}
    .period-switch-btn.active .period-pill-count {{
      background: rgba(15, 23, 42, 0.35);
      color: #fff;
    }}
    html[data-theme="light"] .period-switcher-group {{
      background: #f1f5f9;
      border-color: #cbd5e1;
      box-shadow: 0 2px 8px rgba(15, 23, 42, 0.06);
    }}
    html[data-theme="light"] .period-switcher-label {{
      color: #475569;
    }}
    html[data-theme="light"] .period-switch-btn {{
      color: #334155;
    }}
    html[data-theme="light"] .period-switch-btn:hover {{
      background: #e2e8f0;
      color: #0f172a;
    }}
    html[data-theme="light"] .period-switch-btn.active {{
      background: linear-gradient(135deg, #0284c7, #1d4ed8);
      color: #ffffff;
      border-color: transparent;
      box-shadow: 0 4px 12px rgba(2, 132, 199, 0.3);
    }}
    html[data-theme="light"] .period-pill-count {{
      background: #e2e8f0;
      color: #334155;
    }}
    html[data-theme="light"] .period-switch-btn.active .period-pill-count {{
      background: rgba(255, 255, 255, 0.25);
      color: #ffffff;
    }}

  </style>
</head>
<body>

<div class="dashboard-container">

  <!-- Header with updated title -->
  <header class="glass">
    <div class="brand-section">
      <div class="brand-logo" title="SQA PUMA Regional 11 Papua Maluku"><img src="{logo_data_uri}" alt="Logo SQA PUMA" style="width:100%;height:100%;object-fit:contain;display:block;"></div>
      <div class="brand-text">
        <h1>SQA Puma Update Site LOSS</h1>
        <p>
          <span>Regional 11 Papua Maluku (Area 4)</span>
          <span>•</span>
          <span class="badge-live"><span class="pulse-dot"></span> <span id="header-baseline-text">BASELINE SEPTEMBER 2026</span></span>
          <span>•</span>
          <span>Target H2: 93.0%</span>
        </p>
      </div>
    </div>
    <div class="header-actions" style="flex-wrap:wrap">
      <!-- Switch Periode Data (Pilihan A: Agustus 2026 <-> September 2026) -->
      <div class="period-switcher-group" title="Klik untuk beralih antara Data Agustus 2026 (Sheet Aug 26) dan Data September 2026 (Sheet SEP '26)">
        <span class="period-switcher-label">📅 Periode Data:</span>
        <button type="button" class="period-switch-btn" id="btn-period-aug" onclick="switchPeriod('AUG26')">
          <span>Agustus '26</span>
          <span class="period-pill-count" id="pill-loss-aug">{aug_meta['branchStats']['ALL']['loss']} Loss</span>
        </button>
        <button type="button" class="period-switch-btn active" id="btn-period-sep" onclick="switchPeriod('SEP26')">
          <span>🔥 September '26</span>
          <span class="period-pill-count" id="pill-loss-sep">{sep_meta['branchStats']['ALL']['loss']} Loss</span>
        </button>
      </div>

      <!-- Live Sync & Auto-Refresh Indicator -->
      <div class="sync-status-badge">
        <span class="pulse-dot"></span>
        <span id="sync-status-label">Auto-Refresh: <strong id="refresh-timer" style="color:#38bdf8">30s</strong></span>
      </div>
      <button class="btn-action" id="btn-refresh-manual" onclick="manualTriggerRefresh()" title="Refresh Data Sekarang">
        <svg id="refresh-icon" width="14" height="14" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"/></svg>
        <span>Refresh</span>
      </button>
      <div class="meta-tag" id="last-updated-badge">Update: {current_time_str}</div>
      <button class="btn-action btn-theme-toggle" id="btn-theme-toggle" onclick="toggleTheme()" title="Ganti Tema Gelap / Putih">
        <span>☀️ Versi Putih</span>
      </button>
      <button class="btn-action" onclick="window.print()">
        <svg width="15" height="15" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17 17h2a2 2 0 002-2v-4a2 2 0 00-2-2H5a2 2 0 00-2 2v4a2 2 0 002 2h2m2 4h6a2 2 0 002-2v-4a2 2 0 00-2-2H9a2 2 0 00-2 2v4a2 2 0 002 2zm8-12V5a2 2 0 00-2-2H9a2 2 0 00-2 2v4h10z"></path></svg>
        Print / PDF
      </button>
    </div>
  </header>

  <!-- Branch Filter Bar -->
  <div class="filter-bar glass">
    <div class="filter-group">
      <span class="filter-label">Branch Focus:</span>
      <button class="filter-pill active" id="branch-pill-ALL" onclick="setBranchFilter('ALL')">All Papua Maluku ({sep_meta['branchStats']['ALL']['total']})</button>
      <button class="filter-pill" id="branch-pill-AMBON" onclick="setBranchFilter('AMBON')">Ambon ({sep_meta['branchStats']['AMBON']['total']})</button>
      <button class="filter-pill" id="branch-pill-JAYAPURA" onclick="setBranchFilter('JAYAPURA')">Jayapura ({sep_meta['branchStats']['JAYAPURA']['total']})</button>
      <button class="filter-pill" id="branch-pill-SORONG" onclick="setBranchFilter('SORONG')">Sorong ({sep_meta['branchStats']['SORONG']['total']})</button>
      <button class="filter-pill" id="branch-pill-TIMIKA" onclick="setBranchFilter('TIMIKA')">Timika ({sep_meta['branchStats']['TIMIKA']['total']})</button>
    </div>

    <!-- TAB LANGSUNG DI BRANCH FOCUS (BUKAN DI BAWAH) -->
    <div class="branch-tab-switcher">
      <span style="font-size:11px;color:var(--text-muted);font-weight:700;text-transform:uppercase;letter-spacing:0.5px;">Tab Focus:</span>
      <button class="top-subtab-btn active" data-subtab="subtab-kab-distribusi" onclick="switchBranchSubTab('subtab-kab-distribusi', true)">
        🏙️ Sebaran Kota/Kabupaten
      </button>
      <button class="top-subtab-btn" data-subtab="subtab-prog-tracker" onclick="switchBranchSubTab('subtab-prog-tracker', true)">
        📌 Update Progress & Action Tracker
      </button>
    </div>

    <div style="display:flex;align-items:center;gap:10px;margin-left:auto;">
      <span style="font-size:13px;color:var(--text-muted)">Target Profitability:</span>
      <span class="target-badge">H1: 92.8%</span>
      <span class="target-badge" style="background:rgba(16,185,129,0.15);border-color:rgba(16,185,129,0.3);color:#34d399">H2: 93.0%</span>
    </div>
  </div>

  <!-- KPI Grid (All rounded except percentages) -->
  <div class="kpi-grid">
    <div class="kpi-card glass">
      <div class="kpi-header">
        <span class="kpi-title">Total Active Sites</span>
        <div class="kpi-icon" style="background:rgba(59,130,246,0.15);color:#60a5fa">📡</div>
      </div>
      <div class="kpi-val" id="kpi-total-sites">{sep_meta['branchStats']['ALL']['total']} <span class="unit">sites</span></div>
      <div class="kpi-sub" id="kpi-total-sub">
        <span class="trend-up" style="color:#38bdf8">Sheet SEP '26</span> dibanding Agustus 2026 ({aug_meta['branchStats']['ALL']['total']})
      </div>
      <div class="progress-bar-container">
        <div class="progress-bar-fill" style="width: 100%; background: #3b82f6;"></div>
      </div>
    </div>

    <div class="kpi-card glass">
      <div class="kpi-header">
        <span class="kpi-title" id="kpi-profit-title">Profit Sites (Sep 26)</span>
        <div class="kpi-icon" style="background:rgba(16,185,129,0.15);color:#34d399">💰</div>
      </div>
      <div class="kpi-val" id="kpi-profit-sites" style="color:#34d399">{sep_meta['branchStats']['ALL']['profit']} <span class="unit">sites</span></div>
      <div class="kpi-sub">
        <span class="trend-up" id="kpi-profit-pct">{sep_meta['branchStats']['ALL']['pPct']}</span> <span id="kpi-profit-sub-label">dari populasi (↑ +22 site vs Agu)</span>
      </div>
      <div class="progress-bar-container">
        <div class="progress-bar-fill" id="kpi-profit-bar" style="width: {sep_meta['branchStats']['ALL']['pPct']}; background: #10b981;"></div>
      </div>
    </div>

    <div class="kpi-card glass">
      <div class="kpi-header">
        <span class="kpi-title" id="kpi-loss-title">Non-Profit / Loss Sites (Sep 26)</span>
        <div class="kpi-icon" style="background:rgba(239,68,68,0.15);color:#f87171">⚠️</div>
      </div>
      <div class="kpi-val" id="kpi-loss-sites" style="color:#f87171">{sep_meta['branchStats']['ALL']['loss']} <span class="unit">sites</span></div>
      <div class="kpi-sub">
        <span class="trend-down" id="kpi-loss-pct">{sep_meta['branchStats']['ALL']['lPct']}</span> <span id="kpi-loss-sub-label">rasio loss (↓ -25 site membaik vs Agu)</span>
      </div>
      <div class="progress-bar-container">
        <div class="progress-bar-fill" id="kpi-loss-bar" style="width: {sep_meta['branchStats']['ALL']['lPct']}; background: #ef4444;"></div>
      </div>
    </div>

    <div class="kpi-card glass">
      <div class="kpi-header">
        <span class="kpi-title">H2 Target Gap Recovery</span>
        <div class="kpi-icon" style="background:rgba(245,158,11,0.15);color:#fbbf24">🎯</div>
      </div>
      <div class="kpi-val" id="kpi-gap-sites" style="color:#fbbf24">{sep_meta['branchStats']['ALL']['gap']} <span class="unit">sites</span></div>
      <div class="kpi-sub">
        Site perlu diselamatkan tembus <strong style="color:#fff;margin-left:4px">93.0%</strong>
      </div>
      <div class="progress-bar-container">
        <div class="progress-bar-fill" style="width: 85%; background: #f59e0b;"></div>
      </div>
    </div>
  </div>

  <!-- Visual Charts Section (Right Card per City/Kabupaten) -->
  <div class="section-grid-2">
    <div class="chart-card glass">
      <div class="card-title-bar">
        <div class="card-title">
          <span class="dot"></span>
          <span id="trend-chart-title">Tren % Profit Margin (Januari – September 2026)</span>
        </div>
        <div class="chart-legend">
          <div class="legend-item"><span class="legend-color" style="background:#ff2a44"></span> R11 Total</div>
          <div class="legend-item"><span class="legend-color" style="background:#f59e0b"></span> Ambon</div>
          <div class="legend-item"><span class="legend-color" style="background:#38bdf8"></span> Jayapura</div>
          <div class="legend-item"><span class="legend-color" style="background:#10b981"></span> Sorong</div>
          <div class="legend-item"><span class="legend-color" style="background:#a855f7"></span> Timika</div>
          <div class="legend-item"><span class="legend-color" style="background:#64748b;border:1px dashed #fff"></span> Target 93%</div>
        </div>
      </div>

      <div class="chart-svg-container" id="trend-chart-container"></div>
    </div>

    <!-- Right Card: DYNAMIC KABUPATEN / CITY DISTRIBUTION -->
    <div class="chart-card glass">
      <div class="card-title-bar">
        <div class="card-title">
          <span class="dot" style="background:#f87171"></span>
          <span id="side-card-title">Top Kota/Kabupaten Site Loss (Papua Maluku)</span>
        </div>
        <span class="badge badge-danger" id="side-card-badge">Total: {sep_meta['branchStats']['ALL']['loss']} Sites</span>
      </div>

      <div id="side-card-content" style="display:flex;flex-direction:column;gap:12px;margin-top:8px;">
        <!-- Injected dynamically by updateSideCard() -->
      </div>

      <div id="side-card-alert" style="margin-top:20px;padding:12px 14px;background:rgba(239,68,68,0.08);border:1px solid rgba(239,68,68,0.2);border-radius:10px;">
        <div style="font-size:12px;color:#fca5a5;font-weight:600;margin-bottom:3px" id="side-alert-title">⚠️ Focus Area: NOP Ambon</div>
        <div style="font-size:12px;color:var(--text-muted);line-height:1.5" id="side-alert-desc">
          Ambon memiliki rasio loss 15.30% (achv 84.70% vs target 93%). Untuk mencapai target Regional, <strong>71 site di Ambon</strong> wajib dikonversi menjadi profit.
        </div>
      </div>
    </div>
  </div>

  <!-- Tabs Navigation & Tables -->
  <div class="tabs-wrapper glass" style="padding:24px;">
    <div class="tabs-nav">
      <!-- TAB BRANCH FOCUS (DEFAULT ACTIVE) -->
      <button class="tab-btn active" onclick="switchTab('tab-kabupaten')">
        🏙️ Branch Focus: Kota/Kabupaten & Progress Action <span class="badge badge-danger" style="font-size:10px;margin-left:4px">Per-City & Action</span>
      </button>
      <button class="tab-btn" id="tab-btn-target" onclick="switchTab('tab-target')">🎯 Target vs Realisasi September 2026</button>
      <button class="tab-btn" id="tab-btn-db" onclick="switchTab('tab-aug-db')" style="background:rgba(6,182,212,0.15);color:#38bdf8;border-color:rgba(6,182,212,0.4)">
        <span id="tab-btn-db-label">📋 Database Site Sheet SEP '26</span> <span class="badge badge-info" style="font-size:10px;margin-left:4px">Lengkap 58 Kolom (Termasuk Breakdown, Nilai Cost, Plan Improv, Kategori)</span>
      </button>
      <button class="tab-btn" id="tab-btn-movement" onclick="switchTab('tab-movement')">📅 Pergerakan Bulanan (Jan – Sep 2026)</button>
      <button class="tab-btn" onclick="switchTab('tab-gap')">📊 Matriks Defisit Gap (<3M s/d >10M)</button>
      <button class="tab-btn" onclick="switchTab('tab-regional')">🌐 Pembanding Nasional</button>
    </div>

    <!-- Tab 1: Branch Focus: Site Loss & Progress Tracker (DI SINI TAB PROGRESSNYA!) -->
    <div id="tab-kabupaten" class="tab-pane active">
      <div class="card-title-bar">
        <div>
          <h3 style="font-size:17px;color:#fff;display:flex;align-items:center;gap:10px;">
            <span>Branch Focus: Distribusi Site Loss & Action Progress</span>
            <span class="badge badge-danger" id="kab-header-badge">All Branches</span>
          </h3>
          <p style="font-size:12px;color:var(--text-muted)">
            Pilih tampilan di bawah: <strong>Sebaran per Kota/Kabupaten</strong> atau <strong>Update Progress & Action Tracker</strong> untuk memantau aksi SQA, Sales, NOP, THI, BDBH, Availability, dan Status PL.
          </p>
        </div>
      </div>

      <!-- SUB-TABS INSIDE BRANCH FOCUS -->
      <div class="subtabs-bar">
        <button class="subtab-btn active" data-subtab="subtab-kab-distribusi" id="btn-subtab-distribusi" onclick="switchBranchSubTab('subtab-kab-distribusi')">
          🏙️ Sebaran per Kota / Kabupaten
        </button>
        <button class="subtab-btn" data-subtab="subtab-prog-tracker" id="btn-subtab-prog" onclick="switchBranchSubTab('subtab-prog-tracker')">
          📌 Update Progress & Action Tracker (SQA, Sales, NOP, THI, PL & NOS)
        </button>
      </div>

      <!-- SUBTAB PANE 1: DISTRIBUSI KOTA/KABUPATEN -->
      <div id="subtab-kab-distribusi" class="branch-subtab-pane active">
        <!-- Controls inside Tab Kabupaten -->
        <div class="table-controls-bar">
          <div class="filter-group">
            <span class="filter-label">Filter Branch:</span>
            <select id="kab-select-branch" class="filter-select" onchange="onKabBranchChange(this.value)">
              <option value="ALL">Semua Branch (57 Kab/Kota)</option>
              <option value="AMBON">Branch AMBON (11 Kab/Kota)</option>
              <option value="JAYAPURA">Branch JAYAPURA (16 Kab/Kota)</option>
              <option value="SORONG">Branch SORONG (17 Kab/Kota)</option>
              <option value="TIMIKA">Branch TIMIKA (13 Kab/Kota)</option>
            </select>
          </div>

          <div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap">
            <div class="search-input-wrapper" style="max-width:280px">
              <svg class="search-icon" width="16" height="16" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"></path></svg>
              <input type="text" id="kab-search-input" class="search-input" placeholder="Cari nama Kabupaten / Kota..." oninput="renderKabTable()">
            </div>

            <label style="font-size:12.5px;color:var(--text-muted);display:flex;align-items:center;gap:6px;cursor:pointer">
              <input type="checkbox" id="kab-filter-loss-only" checked onchange="renderKabTable()">
              <span>Hanya tampilkan Kabupaten dengan Site Loss (>0)</span>
            </label>

            <button class="btn-action" onclick="exportTableToCSV('kabupaten-table', 'Site_Loss_Per_Kabupaten.csv')">Export CSV</button>
          </div>
        </div>

        <!-- Kabupaten Table -->
        <div class="table-container" style="max-height:550px">
          <table class="data-table" id="kabupaten-table">
            <thead>
              <tr>
                <th>#</th>
                <th>Kabupaten / Kota</th>
                <th>Branch Induk</th>
                <th>Total Site</th>
                <th>Profit Site</th>
                <th>Loss Site</th>
                <th>Rasio % Loss</th>
                <th>< 3 Mio</th>
                <th>3 - 5 Mio</th>
                <th>5 - 10 Mio</th>
                <th>> 10 Mio (Kritis)</th>
                <th>Total Defisit / Margin (Rp)</th>
                <th>Aksi Drilldown</th>
              </tr>
            </thead>
            <tbody id="kab-table-body">
              <!-- Rendered by JS -->
            </tbody>
          </table>
        </div>

        <!-- Selected Kabupaten Drilldown Site List Box -->
        <div id="kab-drilldown-box" style="display:none;margin-top:24px;padding:20px;background:rgba(15,23,42,0.6);border:1px solid rgba(255,42,68,0.3);border-radius:12px;">
          <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:14px;flex-wrap:wrap;gap:10px">
            <div>
              <h4 id="drilldown-title" style="font-size:15px;color:#fff">Daftar Site Loss di Kabupaten: <span style="color:#ff2a44" id="drilldown-kab-name"></span></h4>
              <p style="font-size:12px;color:var(--text-muted)" id="drilldown-subtitle">Detail teknis, Kelas Site, Status OA, biaya, action plan SQA, Sales & NOP</p>
            </div>
            <button class="page-btn" onclick="closeDrilldown()" style="background:rgba(255,255,255,0.1)">✕ Tutup Detail</button>
          </div>
          <div class="table-container" style="max-height:380px">
            <table class="data-table" id="drilldown-table">
              <thead>
                <tr>
                  <th>Site ID</th>
                  <th>Site Name</th>
                  <th>Kelas Site</th>
                  <th>Status OA</th>
                  <th>Type Site</th>
                  <th>BTS Equip</th>
                  <th>Revenue (Rp)</th>
                  <th>Direct Cost (Rp)</th>
                  <th>Margin / Defisit (Rp)</th>
                  <th>GAP Dist</th>
                  <th>Inline SSGJ</th>
                  <th>Action SQA</th>
                  <th>Action Sales</th>
                  <th>Action NOP</th>
                  <th>Type Transport</th>
                  <th>Update Availability</th>
                  <th>Status PL Terakhir</th>
                  <th class="new-col-th" style="background:rgba(14,165,233,0.2);color:#7dd3fc">Breakdown Cost</th>
                  <th class="new-col-th" style="background:rgba(14,165,233,0.2);color:#7dd3fc">Nilai Cost</th>
                  <th class="new-col-th" style="background:rgba(14,165,233,0.15);color:#7dd3fc">Plan Improvement</th>
                  <th class="new-col-th" style="background:rgba(6,182,212,0.2);color:#67e8f9">Cek / Kategori Site</th>
                </tr>
              </thead>
              <tbody id="drilldown-tbody">
                <!-- Drilldown rows -->
              </tbody>
            </table>
          </div>
        </div>
      </div>

      <!-- SUBTAB PANE 2: UPDATE PROGRESS & ACTION TRACKER -->
      <div id="subtab-prog-tracker" class="branch-subtab-pane">

        <!-- DEDICATED PRIORITY ANALYTICS & DIAGRAM BATANG CARD -->
        <div class="priority-analytics-card">
          <div class="prog-analytics-header">
            <div>
              <div class="card-title" style="display:flex;align-items:center;gap:10px">
                <span class="dot" style="background:#f59e0b;box-shadow:0 0 10px #f59e0b"></span>
                <span style="font-size:16px;font-weight:800;letter-spacing:0.3px">Diagram Batang Site Priority (Sales & Network)</span>
                <span class="badge" style="background:rgba(245,158,11,0.2);color:#fcd34d;border:1px solid rgba(245,158,11,0.4);font-size:11px" id="priority-main-badge">78 Priority Sites (Papua Maluku)</span>
              </div>
              <p style="font-size:12px;color:var(--text-muted);margin-top:4px">
                Visualisasi diagram batang volume site <strong>Priority Sales & Priority Network</strong> beserta status progres pemenuhan action plan (<strong>Sales:</strong> Action Sales & NOP | <strong>Network:</strong> SQA, RHI, THI, Availability, PL/MBA & Cost).
              </p>
            </div>
          </div>

          <!-- Priority KPI Summary Cards -->
          <div class="priority-kpi-grid">
            <div class="priority-kpi-item" onclick="setPriorityFilter('SALES_ONLY')" title="Klik untuk filter tabel ke Priority Sales">
              <div style="display:flex;justify-content:space-between;align-items:center">
                <span style="font-size:12px;color:#fcd34d;font-weight:700">💼 Priority Sales</span>
                <span class="badge" style="background:rgba(245,158,11,0.25);color:#fcd34d" id="kpi-prio-sales-badge">41 Sites</span>
              </div>
              <div style="font-size:22px;font-weight:800;color:#fff;margin:6px 0 2px" id="kpi-prio-sales-val">41 <span style="font-size:12px;color:var(--text-dim);font-weight:normal">sites</span></div>
              <div style="font-size:11px;color:var(--text-muted);display:flex;justify-content:space-between">
                <span id="kpi-prio-sales-tiers">P1: 27 | P2: 13 | P3: 1</span>
                <span style="color:#34d399;font-weight:600" id="kpi-prio-sales-prog">Progress: 17</span>
              </div>
            </div>

            <div class="priority-kpi-item" onclick="setPriorityFilter('NET_ONLY')" title="Klik untuk filter tabel ke Priority Network">
              <div style="display:flex;justify-content:space-between;align-items:center">
                <span style="font-size:12px;color:#7dd3fc;font-weight:700">📡 Priority Network</span>
                <span class="badge" style="background:rgba(56,189,248,0.25);color:#7dd3fc" id="kpi-prio-net-badge">22 Sites</span>
              </div>
              <div style="font-size:22px;font-weight:800;color:#fff;margin:6px 0 2px" id="kpi-prio-net-val">22 <span style="font-size:12px;color:var(--text-dim);font-weight:normal">sites</span></div>
              <div style="font-size:11px;color:var(--text-muted);display:flex;justify-content:space-between">
                <span id="kpi-prio-net-tiers">P1: 2 | P2: 19 | P3: 1</span>
                <span style="color:#34d399;font-weight:600" id="kpi-prio-net-prog">Progress: 20</span>
              </div>
            </div>

            <div class="priority-kpi-item" onclick="setPriorityFilter('ANY_PRIO')" title="Klik untuk filter ke semua site priority">
              <div style="display:flex;justify-content:space-between;align-items:center">
                <span style="font-size:12px;color:#d8b4fe;font-weight:700">⭐ Total Site Priority</span>
                <span class="badge badge-purple" id="kpi-prio-total-badge">62 Sites</span>
              </div>
              <div style="font-size:22px;font-weight:800;color:#fff;margin:6px 0 2px" id="kpi-prio-total-val">62 <span style="font-size:12px;color:var(--text-dim);font-weight:normal">sites</span></div>
              <div style="font-size:11px;color:var(--text-muted);display:flex;justify-content:space-between">
                <span id="kpi-prio-total-ratio">Dari 217 Loss</span>
                <span style="color:#60a5fa;font-weight:600" id="kpi-prio-total-prog">Progress: 36</span>
              </div>
            </div>

            <div class="priority-kpi-item" onclick="setPriorityFilter('ACTION_DONE')" title="Klik untuk filter site priority yang sudah ada progress">
              <div style="display:flex;justify-content:space-between;align-items:center">
                <span style="font-size:12px;color:#34d399;font-weight:700">✅ Sudah Ada Action</span>
                <span class="badge badge-success" id="kpi-prio-done-pct">58.1%</span>
              </div>
              <div style="font-size:22px;font-weight:800;color:#34d399;margin:6px 0 2px" id="kpi-prio-done-val">36 <span style="font-size:12px;color:var(--text-dim);font-weight:normal">sites</span></div>
              <div style="font-size:11px;color:var(--text-muted)" id="kpi-prio-done-sub">Site priority dengan action terisi</div>
            </div>

            <div class="priority-kpi-item" onclick="setPriorityFilter('ACTION_PENDING')" title="Klik untuk filter site priority yang belum ada progress">
              <div style="display:flex;justify-content:space-between;align-items:center">
                <span style="font-size:12px;color:#f87171;font-weight:700">⏳ Belum Ada Action</span>
                <span class="badge badge-danger" id="kpi-prio-pending-pct">41.9%</span>
              </div>
              <div style="font-size:22px;font-weight:800;color:#f87171;margin:6px 0 2px" id="kpi-prio-pending-val">26 <span style="font-size:12px;color:var(--text-dim);font-weight:normal">sites</span></div>
              <div style="font-size:11px;color:var(--text-muted)">Memerlukan tindak lanjut PIC</div>
            </div>
          </div>

          <!-- Priority Diagram Batang Grid -->
          <div class="prog-charts-grid" style="margin-top:10px">
            <!-- Diagram Batang 1: Berapa Site Priority Sales & Network -->
            <div class="prog-chart-box">
              <div class="prog-chart-title">
                <span>📊 Diagram Batang: Komposisi Site Priority (Sales vs Network)</span>
                <span style="font-size:11px;color:var(--text-dim);font-weight:normal">Perbandingan Tier & Branch</span>
              </div>
              <div id="priority-volume-bars" style="display:flex;flex-direction:column;gap:12px;margin-top:10px">
                <!-- Rendered dynamically -->
              </div>
            </div>

            <!-- Diagram Batang 2: Progressnya Sudah Berapa Banyak -->
            <div class="prog-chart-box">
              <div class="prog-chart-title">
                <span>🚀 Diagram Batang: Progress Action Plan Site Priority</span>
                <span style="font-size:11px;color:var(--text-dim);font-weight:normal">Berapa site yang sudah ada progress</span>
              </div>
              <div id="priority-progress-bars" style="display:flex;flex-direction:column;gap:12px;margin-top:10px">
                <!-- Rendered dynamically -->
              </div>
            </div>
          </div>
        </div>


        <!-- DIAGRAM BATANG ANALISIS UPDATE PROGRESS & ACTION TRACKER -->
        <div class="prog-analytics-card">
          <div class="prog-analytics-header">
            <div>
              <div class="card-title">
                <span class="dot" style="background:#06b6d4"></span>
                <span>Diagram Batang Progres Input Action Tracker</span>
                <span class="badge badge-info" id="prog-chart-badge" style="font-size:11px;margin-left:6px">242 Loss Sites</span>
              </div>
              <p style="font-size:12px;color:var(--text-muted);margin-top:4px">
                Komparasi Rencana Aksi (Action to be Taken) vs Realisasi Progres (Update Action Taken) dan pemenuhan per Branch.
              </p>
            </div>
            <div class="prog-kpi-chips" id="prog-kpi-chips">
              <!-- Rendered dynamically -->
            </div>
          </div>

          <div class="prog-charts-grid">
            <!-- Diagram Batang 1: Rencana Aksi (Action to be Taken) -->
            <div class="prog-chart-box">
              <div class="prog-chart-title">
                <span>📋 Rencana Aksi (Action to be Taken)</span>
                <span style="font-size:11px;color:var(--text-dim);font-weight:normal">Jumlah site rencana yang perlu aksi</span>
              </div>
              <div id="prog-plan-bars" style="display:flex;flex-direction:column;gap:11px;margin-top:6px">
                <!-- Injected dynamically by renderProgCharts() -->
              </div>
            </div>

            <!-- Diagram Batang 2: Realisasi Progres (Update Action Taken) -->
            <div class="prog-chart-box">
              <div class="prog-chart-title">
                <span>✅ Realisasi Progres (Update Action Taken)</span>
                <span style="font-size:11px;color:var(--text-dim);font-weight:normal">Jumlah site yang sudah action / update</span>
              </div>
              <div id="prog-domain-bars" style="display:flex;flex-direction:column;gap:11px;margin-top:6px">
                <!-- Injected dynamically by renderProgCharts() -->
              </div>
            </div>
          </div>

          <!-- Diagram Batang 3: Komparasi per Branch -->
          <div class="prog-chart-box" style="margin-top:16px">
            <div class="prog-chart-title">
              <span>🏙️ Pemenuhan Action Plan per Branch</span>
              <span style="font-size:11px;color:var(--text-dim);font-weight:normal">Site yang memiliki aksi terisi • Klik bar untuk filter tabel</span>
            </div>
            <div id="prog-branch-bars" style="display:flex;flex-direction:column;gap:10px;margin-top:6px">
              <!-- Injected dynamically by renderProgCharts() -->
            </div>
          </div>
        </div>

        
        <!-- QUICK PRIORITY FILTER PILLS -->
        <div class="priority-filter-bar">
          <span style="font-size:12px;font-weight:700;color:#fff;display:flex;align-items:center;gap:6px">
            <span style="color:#f59e0b">⭐</span> Filter Priority:
          </span>
          <button class="priority-pill active" id="prio-pill-all" onclick="setPriorityFilter('ALL')">Semua Site (242 Loss)</button>
          <button class="priority-pill" id="prio-pill-any" onclick="setPriorityFilter('ANY_PRIO')">⭐ Semua Priority Sites (78)</button>
          <button class="priority-pill" id="prio-pill-sales" onclick="setPriorityFilter('SALES_ONLY')">💼 Priority Sales (56)</button>
          <button class="priority-pill" id="prio-pill-net" onclick="setPriorityFilter('NET_ONLY')">📡 Priority Network (23)</button>
          <button class="priority-pill" id="prio-pill-p1" onclick="setPriorityFilter('P1_ONLY')">🔥 Priority P1 (42)</button>
          <button class="priority-pill" id="prio-pill-p2" onclick="setPriorityFilter('P2_ONLY')">⚡ Priority P2 (35)</button>
          <button class="priority-pill" id="prio-pill-done" onclick="setPriorityFilter('ACTION_DONE')">✅ Sudah Ada Action (39)</button>
          <button class="priority-pill" id="prio-pill-pending" onclick="setPriorityFilter('ACTION_PENDING')">⏳ Belum Ada Action (39)</button>
        </div>

        <!-- QUICK NETWORK PART FILTER PILLS -->
        <div class="priority-filter-bar" style="margin-top:6px;background:rgba(168,85,247,0.06);border-color:rgba(168,85,247,0.25)">
          <span style="font-size:12px;font-weight:700;color:#d8b4fe;display:flex;align-items:center;gap:6px">
            <span>📡</span> Part Priority Net:
          </span>
          <button class="priority-pill active" id="prio-part-all" onclick="setNetPartFilter('ALL')">Semua Part Net</button>
          <button class="priority-pill" id="prio-part-sqa" onclick="setNetPartFilter('PART_SQA')">🎯 Part SQA (23)</button>
          <button class="priority-pill" id="prio-part-rhi" onclick="setNetPartFilter('PART_RHI')">📡 Part RHI (16)</button>
          <button class="priority-pill" id="prio-part-thi" onclick="setNetPartFilter('PART_THI')">📶 Part THI (0)</button>
          <button class="priority-pill" id="prio-part-avail" onclick="setNetPartFilter('PART_AVAIL')">⚡ Part Avail (11)</button>
          <button class="priority-pill" id="prio-part-pl" onclick="setNetPartFilter('PART_PL')">🔌 Part PL/MBA (12)</button>
          <button class="priority-pill" id="prio-part-cost" onclick="setNetPartFilter('PART_COST')">💰 Part Cost (0)</button>
        </div>

        <!-- Controls for Progress Tracker -->
        <div class="table-controls-bar">
          <div class="search-input-wrapper">
            <svg class="search-icon" width="16" height="16" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"></path></svg>
            <input type="text" id="prog-search-input" class="search-input" placeholder="Cari Site ID, Site Name, Kabupaten, Action..." oninput="onProgSearch()">
          </div>

          <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap">
            <select id="prog-filter-nop" class="filter-select" onchange="applyProgFilters()">
              <option value="ALL">Semua Branch</option>
              <option value="AMBON">Branch AMBON</option>
              <option value="JAYAPURA">Branch JAYAPURA</option>
              <option value="SORONG">Branch SORONG</option>
              <option value="TIMIKA">Branch TIMIKA</option>
            </select>

            
            <select id="prog-filter-priority-sales" class="filter-select" onchange="applyProgFilters()" style="border-color:rgba(245,158,11,0.4);background:rgba(245,158,11,0.08);color:#fcd34d">
              <option value="ALL">Semua Priority Sales</option>
              <option value="PRIORITY_ONLY">Hanya Priority Sales (P1/P2/P3)</option>
              <option value="P1">P1 Sales Priority (39)</option>
              <option value="P2">P2 Sales Priority (16)</option>
              <option value="P3">P3 Sales Priority (1)</option>
              <option value="Not Priority">Not Priority Sales</option>
            </select>

            <select id="prog-filter-priority-net" class="filter-select" onchange="applyProgFilters()" style="border-color:rgba(56,189,248,0.4);background:rgba(56,189,248,0.08);color:#7dd3fc">
              <option value="ALL">Semua Priority Network</option>
              <option value="PRIORITY_ONLY">Hanya Priority Network (P1/P2/P3)</option>
              <option value="P1">P1 Network Priority (3)</option>
              <option value="P2">P2 Network Priority (19)</option>
              <option value="P3">P3 Network Priority (1)</option>
              <option value="Not Priority">Not Priority Network</option>
            </select>

            <select id="prog-filter-net-part" class="filter-select" onchange="onNetPartDropdownChange()" style="border-color:rgba(168,85,247,0.4);background:rgba(168,85,247,0.08);color:#d8b4fe">
              <option value="ALL">Part Priority Net: Semua</option>
              <option value="PART_SQA">🎯 Part SQA (Col AH/AI)</option>
              <option value="PART_RHI">📡 Part RHI (Col AO/AP)</option>
              <option value="PART_THI">📶 Part THI (Col AR/AS)</option>
              <option value="PART_AVAIL">⚡ Part Avail (Col AW)</option>
              <option value="PART_PL">🔌 Part PL / MBA (Col AY)</option>
              <option value="PART_COST">💰 Part Cost (Col AZ/BA)</option>
              <option value="PART_ANY">✅ Sudah Ada Part Net (Any)</option>
              <option value="PART_NONE">⏳ Belum Ada Part Net</option>
            </select>

            <select id="prog-filter-pnl" class="filter-select" onchange="applyProgFilters()">
              <option value="Loss" selected>Loss Sites Only (242)</option>
              <option value="ALL">Semua Status (Profit/Loss)</option>
              <option value="Profit">Profit Sites Only (2,549)</option>
            </select>

            <select id="prog-filter-oa" class="filter-select" onchange="applyProgFilters()" style="border-color:rgba(16,185,129,0.35);background:rgba(16,185,129,0.06);color:#34d399">
              <option value="ALL">Semua Status OA</option>
              <option value="OA">OA Only</option>
              <option value="Not OA">Not OA Only</option>
            </select>

            <select id="prog-filter-plweek" class="filter-select" onchange="applyProgFilters()">
              <option value="ALL">Semua Status PL Terakhir</option>
              <option value="CLEAR">CLEAR</option>
              <option value="UNCLEAR">Belum CLEAR</option>
            </select>

            <select id="prog-filter-transport" class="filter-select" onchange="applyProgFilters()">
              <option value="ALL">Semua Type Transport</option>
              <option value="FO TELKOM">FO TELKOM</option>
              <option value="Radio IP Telkom">Radio IP Telkom</option>
              <option value="FO TSEL">FO TSEL</option>
            </select>

            <select id="prog-filter-cek" class="filter-select" onchange="applyProgFilters()" style="border-color:rgba(6,182,212,0.4);background:rgba(6,182,212,0.08);color:#67e8f9">
              <option value="ALL">Semua Kategori (Kolom BB)</option>
              <option value="Reguler">Reguler</option>
              <option value="Repeater/CDC/IDR/SPS">Repeater/CDC/IDR/SPS</option>
              <option value="Business & Commercial">Business & Commercial</option>
              <option value="N3T">N3T</option>
              <option value="B2B">B2B</option>
              <option value="VIP">VIP</option>
              <option value="Public Open Space & Tourism">Public Open Space & Tourism</option>
              <option value="Public Transportation">Public Transportation</option>
              <option value="Reguler / CDC">Reguler / CDC</option>
              <option value="TTC/MSC/STO">TTC/MSC/STO</option>
            </select>

            <!-- Filter Kolom AG: Site Loss Tracker -->
            <select id="prog-filter-ag" class="filter-select" onchange="applyProgFilters()" style="border-color:rgba(234,179,8,0.45);background:rgba(234,179,8,0.1);color:#fde047;font-weight:600">
              <option value="ALL">Semua Site Loss Tracker (Kolom AG)</option>
              <option value="Program H2">Program H2 (32 Site Loss)</option>
              <option value="0">0 (Non Program H2)</option>
              <option value="#N/A">#N/A</option>
            </select>

            <!-- Fitur Sort Besar ke Kecil (Shorting) -->
            <select id="prog-sort-select" class="filter-select" onchange="onProgSortChange(this.value)" style="border-color:rgba(56,189,248,0.45);background:rgba(56,189,248,0.12);color:#38bdf8;font-weight:700">
              <option value="NONE">⚡ Urutkan: Default</option>
              <option value="DEFISIT_DESC">🔻 Defisit Terbesar (Besar ke Kecil)</option>
              <option value="DEFISIT_ASC">🔺 Defisit Terkecil (Kecil ke Besar)</option>
              <option value="COST_DESC">🔻 Cost Terbesar (Besar ke Kecil)</option>
              <option value="REV_DESC">🔻 Revenue Terbesar (Besar ke Kecil)</option>
              <option value="NILAI_COST_DESC">🔻 Nilai Cost Terbesar (Kolom BA)</option>
              <option value="SITE_ID_ASC">🔤 Site ID (A - Z)</option>
            </select>

            <select id="prog-page-size" class="filter-select" onchange="changeProgPageSize(this.value)">
              <option value="25">25 per hal</option>
              <option value="50" selected>50 per hal</option>
              <option value="100">100 per hal</option>
              <option value="250">250 per hal</option>
            </select>

            <button class="btn-action" onclick="exportProgressToCSV()">
              <svg width="15" height="15" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"></path></svg>
              Export CSV
            </button>
          </div>
        </div>

        <!-- Progress Table Container -->
        <div class="table-container" style="max-height:600px">
          <table class="data-table" id="progress-database-table">
            <thead>
              <tr>
                <th>#</th>
                <th onclick="sortProgTable('site_id')" style="cursor:pointer;user-select:none" title="Klik untuk mengurutkan Site ID">Site ID <span id="th-sort-prog-site_id" style="font-size:10px;color:var(--text-dim)">⬍</span></th>
                <th onclick="sortProgTable('site_name')" style="cursor:pointer;user-select:none" title="Klik untuk mengurutkan Site Name">Site Name <span id="th-sort-prog-site_name" style="font-size:10px;color:var(--text-dim)">⬍</span></th>
                <th onclick="sortProgTable('branch')" style="cursor:pointer;user-select:none" title="Klik untuk mengurutkan Branch">Branch <span id="th-sort-prog-branch" style="font-size:10px;color:var(--text-dim)">⬍</span></th>
                <th onclick="sortProgTable('kabupaten')" style="cursor:pointer;user-select:none" title="Klik untuk mengurutkan Kabupaten">Kabupaten <span id="th-sort-prog-kabupaten" style="font-size:10px;color:var(--text-dim)">⬍</span></th>
                <th onclick="sortProgTable('pnl')" style="cursor:pointer;user-select:none" title="Klik untuk mengurutkan Status PnL">Status PnL <span id="th-sort-prog-pnl" style="font-size:10px;color:var(--text-dim)">⬍</span></th>
                <th onclick="sortProgTable('defisit')" style="cursor:pointer;user-select:none;background:rgba(239,68,68,0.2);color:#fca5a5" title="Klik untuk mengurutkan Defisit Terbesar (Besar ke Kecil)">Defisit / Margin (Rp) <span id="th-sort-prog-defisit" style="font-size:10px;color:#fca5a5">⬍</span></th>
                <th onclick="sortProgTable('cost')" style="cursor:pointer;user-select:none;background:rgba(245,158,11,0.18);color:#fcd34d" title="Klik untuk mengurutkan Cost Terbesar (Besar ke Kecil)">Cost (Rp) <span id="th-sort-prog-cost" style="font-size:10px;color:#fcd34d">⬍</span></th>
                <th>Kelas Site</th>
                <th>Status OA</th>
                <th onclick="sortProgTable('ag')" style="cursor:pointer;user-select:none;background:rgba(234,179,8,0.25);color:#fde047" title="Klik untuk mengurutkan Kolom AG (Site Loss Tracker)">Site Loss Tracker (AG) <span id="th-sort-prog-ag" style="font-size:10px;color:#fde047">⬍</span></th>
                <th class="new-col-th" style="background:rgba(245,158,11,0.25);color:#fcd34d">Priority Sales</th>
                <th class="new-col-th" style="background:rgba(56,189,248,0.25);color:#7dd3fc">Priority Network</th>
                <th class="new-col-th" style="background:rgba(239,68,68,0.2);color:#fca5a5">Action to be taken SQA</th>
                <th class="new-col-th" style="background:rgba(239,68,68,0.15);color:#fca5a5">Update action taken sqa</th>
                <th class="new-col-th" style="background:rgba(245,158,11,0.2);color:#fcd34d">Action to be taken Sales</th>
                <th class="new-col-th" style="background:rgba(245,158,11,0.15);color:#fcd34d">Update action taken sales</th>
                <th class="new-col-th" style="background:rgba(16,185,129,0.2);color:#6ee7b7">Action to be taken NOP</th>
                <th class="new-col-th" style="background:rgba(16,185,129,0.15);color:#6ee7b7">Update action taken NOP</th>
                <th class="new-col-th" style="background:rgba(167,139,250,0.2);color:#c4b5fd">update RHI Week terakhir</th>
                <th class="new-col-th" style="background:rgba(167,139,250,0.15);color:#c4b5fd">Root Cause RHI</th>
                <th class="new-col-th" style="background:rgba(167,139,250,0.12);color:#c4b5fd">propose action RHI</th>
                <th class="new-col-th" style="background:rgba(56,189,248,0.2);color:#7dd3fc">update THI Week terakhir</th>
                <th class="new-col-th" style="background:rgba(56,189,248,0.15);color:#7dd3fc">Root Cause THI</th>
                <th class="new-col-th" style="background:rgba(56,189,248,0.12);color:#7dd3fc">propose action THI</th>
                <th class="new-col-th">Typer Transport</th>
                <th class="new-col-th">utilisasi avg weekly BDBH</th>
                <th class="new-col-th">update avg status availability</th>
                <th class="new-col-th">plan action availability</th>
                <th class="new-col-th" style="background:rgba(59,130,246,0.25);color:#93c5fd">Status PL week terakhir</th>
                <th class="new-col-th" style="background:rgba(59,130,246,0.18);color:#93c5fd">propose action (MBA)</th>
                <th class="new-col-th" style="background:rgba(14,165,233,0.2);color:#7dd3fc">Breakdown cost terbesar</th>
                <th class="new-col-th" onclick="sortProgTable('nilai_cost')" style="cursor:pointer;user-select:none;background:rgba(14,165,233,0.2);color:#7dd3fc" title="Klik untuk mengurutkan Nilai Cost Terbesar">Nilai Cost <span id="th-sort-prog-nilai_cost" style="font-size:10px;color:#7dd3fc">⬍</span></th>
                <th class="new-col-th" style="background:rgba(14,165,233,0.15);color:#7dd3fc">plan improvement cost terbesar</th>
                <th class="new-col-th" style="background:rgba(6,182,212,0.2);color:#67e8f9">Cek / Kategori Site</th>
              </tr>
            </thead>
            <tbody id="prog-table-body">
              <!-- Rows rendered by JS -->
            </tbody>
          </table>
        </div>

        <!-- Pagination Footer for Progress Tab -->
        <div class="pagination-bar">
          <div id="prog-pagination-info">Menampilkan 1 - 50</div>
          <div class="pagination-controls">
            <button class="page-btn" id="prog-btn-first" onclick="goToProgPage(1)">« First</button>
            <button class="page-btn" id="prog-btn-prev" onclick="prevProgPage()">‹ Prev</button>
            <span style="font-family:'JetBrains Mono',monospace;font-size:12px;margin:0 8px">
              Hal <span id="prog-current-page-num" style="color:#fff;font-weight:bold">1</span> / <span id="prog-total-pages-num">1</span>
            </span>
            <button class="page-btn" id="prog-btn-next" onclick="nextProgPage()">Next ›</button>
            <button class="page-btn" id="prog-btn-last" onclick="goToLastProgPage()">Last »</button>
          </div>
        </div>
      </div>
    </div>

    <!-- Tab 2: Target vs Realisasi (Dynamic SEP '26 / Aug 26) -->
    <div id="tab-target" class="tab-pane">
      <div class="card-title-bar">
        <div>
          <h3 style="font-size:15px;color:#fff" id="target-tab-title">Tabel Pencapaian & Gap Recovery Target H2 (93.0%) — September 2026</h3>
          <p style="font-size:12px;color:var(--text-muted)" id="target-tab-desc">Evaluasi pencapaian profitabilitas per Branch berdasarkan Sheet SEP '26. Seluruh target site dan gap dibulatkan.</p>
        </div>
        <button class="btn-action" onclick="exportTableToCSV('target-table', 'Target_Achv_2026.csv')">Export CSV</button>
      </div>

      <div class="table-container">
        <table class="data-table" id="target-table">
          <thead>
            <tr>
              <th>Branch / NOP</th>
              <th>Total Site</th>
              <th>Profit Site</th>
              <th>Loss Site</th>
              <th>Target Site (93%)</th>
              <th>Target %</th>
              <th>Realisasi %</th>
              <th>Gap Site (Perlu Diclearkan)</th>
              <th>Status Evaluasi</th>
            </tr>
          </thead>
          <tbody id="target-table-tbody">
            <tr>
              <td><strong style="color:#f87171">AMBON</strong></td>
              <td>{sep_meta['branchStats']['AMBON']['total']}</td>
              <td>{sep_meta['branchStats']['AMBON']['profit']}</td>
              <td>{sep_meta['branchStats']['AMBON']['loss']}</td>
              <td><strong>{sep_meta['branchStats']['AMBON']['target']}</strong></td>
              <td><span class="badge badge-warning">93.0%</span></td>
              <td><strong style="color:#f87171">{sep_meta['branchStats']['AMBON']['pPct']}</strong></td>
              <td><span class="badge badge-danger">{sep_meta['branchStats']['AMBON']['gap']} Site</span></td>
              <td><span class="badge badge-danger">Belum Capai Target</span></td>
            </tr>
            <tr>
              <td><strong style="color:#38bdf8">JAYAPURA</strong></td>
              <td>{sep_meta['branchStats']['JAYAPURA']['total']}</td>
              <td>{sep_meta['branchStats']['JAYAPURA']['profit']}</td>
              <td>{sep_meta['branchStats']['JAYAPURA']['loss']}</td>
              <td><strong>{sep_meta['branchStats']['JAYAPURA']['target']}</strong></td>
              <td><span class="badge badge-warning">93.0%</span></td>
              <td><strong style="color:#34d399">{sep_meta['branchStats']['JAYAPURA']['pPct']}</strong></td>
              <td><span class="badge badge-success">{sep_meta['branchStats']['JAYAPURA']['gap']} Site</span></td>
              <td><span class="badge badge-success">Target Tercapai ✓</span></td>
            </tr>
            <tr>
              <td><strong style="color:#10b981">SORONG</strong></td>
              <td>{sep_meta['branchStats']['SORONG']['total']}</td>
              <td>{sep_meta['branchStats']['SORONG']['profit']}</td>
              <td>{sep_meta['branchStats']['SORONG']['loss']}</td>
              <td><strong>{sep_meta['branchStats']['SORONG']['target']}</strong></td>
              <td><span class="badge badge-warning">93.0%</span></td>
              <td><strong style="color:#34d399">{sep_meta['branchStats']['SORONG']['pPct']}</strong></td>
              <td><span class="badge badge-success">{sep_meta['branchStats']['SORONG']['gap']} Site</span></td>
              <td><span class="badge badge-success">Target Tercapai ✓</span></td>
            </tr>
            <tr>
              <td><strong style="color:#a855f7">TIMIKA</strong></td>
              <td>{sep_meta['branchStats']['TIMIKA']['total']}</td>
              <td>{sep_meta['branchStats']['TIMIKA']['profit']}</td>
              <td>{sep_meta['branchStats']['TIMIKA']['loss']}</td>
              <td><strong>{sep_meta['branchStats']['TIMIKA']['target']}</strong></td>
              <td><span class="badge badge-warning">93.0%</span></td>
              <td><strong style="color:#34d399">{sep_meta['branchStats']['TIMIKA']['pPct']}</strong></td>
              <td><span class="badge badge-success">{sep_meta['branchStats']['TIMIKA']['gap']} Site</span></td>
              <td><span class="badge badge-success">Target Tercapai ✓</span></td>
            </tr>
            <tr class="total-row">
              <td><strong>TOTAL R11 (PAPUA MALUKU)</strong></td>
              <td><strong>{sep_meta['branchStats']['ALL']['total']}</strong></td>
              <td><strong>{sep_meta['branchStats']['ALL']['profit']}</strong></td>
              <td><strong>{sep_meta['branchStats']['ALL']['loss']}</strong></td>
              <td><strong>{sep_meta['branchStats']['ALL']['target']}</strong></td>
              <td><span class="badge badge-warning">93.0%</span></td>
              <td><strong style="color:#fbbf24">{sep_meta['branchStats']['ALL']['pPct']}</strong></td>
              <td><span class="badge badge-danger">{sep_meta['branchStats']['ALL']['gap']} Site</span></td>
              <td><span class="badge badge-warning">Gap Defisit: {abs(sep_meta['branchStats']['ALL']['gap_num'])} Site</span></td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>

    <!-- Tab 3: Database Site (SEP '26 / Aug 26) -->
    <div id="tab-aug-db" class="tab-pane">
      <div class="card-title-bar">
        <div>
          <h3 style="font-size:16px;color:#fff;display:flex;align-items:center;gap:10px;">
            <span id="db-tab-title">Database Site Sheet SEP '26 (Lengkap 58 Kolom)</span>
            <span class="badge badge-info" style="font-size:11px" id="db-tab-badge">{sep_meta['branchStats']['ALL']['total']} Sites</span>
          </h3>
          <p style="font-size:12px;color:var(--text-muted)" id="db-tab-desc">
            Menampilkan seluruh data site sheet SEP '26 beserta <strong>Kelas Site</strong>, <strong>Status OA</strong>, <strong>Action SQA, Sales & NOP</strong>, performa THI, Transport, BDBH, Availability, dan Status PL.
          </p>
        </div>
        <button class="btn-action" onclick="exportFilteredAugToCSV()">
          <svg width="15" height="15" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"></path></svg>
          Export Filtered CSV
        </button>
      </div>

      <!-- Controls: Search & Filters -->
      <div class="table-controls-bar">
        <div class="search-input-wrapper">
          <svg class="search-icon" width="16" height="16" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"></path></svg>
          <input type="text" id="aug-search-input" class="search-input" placeholder="Cari Site ID, Site Name, NOP, Kabupaten, Kelas..." oninput="onAugSearch()">
        </div>

        <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap">
          <select id="aug-filter-nop" class="filter-select" onchange="applyAugFilters()">
            <option value="ALL">Semua NOP</option>
            <option value="NOP AMBON">NOP AMBON ({sep_meta['nopCounts']['NOP AMBON']})</option>
            <option value="NOP JAYAPURA">NOP JAYAPURA ({sep_meta['nopCounts']['NOP JAYAPURA']})</option>
            <option value="NOP SORONG">NOP SORONG ({sep_meta['nopCounts']['NOP SORONG']})</option>
            <option value="NOP TIMIKA">NOP TIMIKA ({sep_meta['nopCounts']['NOP TIMIKA']})</option>
          </select>

          
          <select id="aug-filter-priority-sales" class="filter-select" onchange="applyAugFilters()" style="border-color:rgba(245,158,11,0.4);background:rgba(245,158,11,0.08);color:#fcd34d">
            <option value="ALL">Semua Priority Sales</option>
            <option value="PRIORITY_ONLY">Hanya Priority (P1/P2/P3)</option>
            <option value="P1">P1 Sales Priority</option>
            <option value="P2">P2 Sales Priority</option>
            <option value="P3">P3 Sales Priority</option>
            <option value="Not Priority">Not Priority Sales</option>
          </select>

          <select id="aug-filter-priority-net" class="filter-select" onchange="applyAugFilters()" style="border-color:rgba(56,189,248,0.4);background:rgba(56,189,248,0.08);color:#7dd3fc">
            <option value="ALL">Semua Priority Network</option>
            <option value="PRIORITY_ONLY">Hanya Priority (P1/P2/P3)</option>
            <option value="P1">P1 Network Priority</option>
            <option value="P2">P2 Network Priority</option>
            <option value="P3">P3 Network Priority</option>
            <option value="Not Priority">Not Priority Network</option>
          </select>

          <select id="aug-filter-pnl" class="filter-select" onchange="applyAugFilters()">
            <option value="ALL">Semua Status (Profit/Loss)</option>
            <option value="Loss">Loss Only ({sep_meta['branchStats']['ALL']['loss']})</option>
            <option value="Profit">Profit Only ({sep_meta['branchStats']['ALL']['profit']})</option>
          </select>

          <select id="aug-filter-class" class="filter-select" onchange="applyAugFilters()">
            <option value="ALL">Semua Kelas Site</option>
            <option value="DIAMOND">DIAMOND</option>
            <option value="PLATINUM">PLATINUM</option>
            <option value="GOLD">GOLD</option>
            <option value="SILVER">SILVER</option>
            <option value="BRONZE">BRONZE</option>
          </select>

          <select id="aug-filter-oa" class="filter-select" onchange="applyAugFilters()">
            <option value="ALL">Semua Status OA</option>
            <option value="OA">OA ({sep_meta['oaCount']:,})</option>
            <option value="Not Oa">Not OA ({sep_meta['notOaCount']:,})</option>
          </select>

          <select id="aug-filter-gap" class="filter-select" onchange="applyAugFilters()">
            <option value="ALL">Semua GAP Dist</option>
            <option value="<3 Mio">< 3 Mio</option>
            <option value="3-5 Mio">3-5 Mio</option>
            <option value="5-10 Mio">5-10 Mio</option>
            <option value=">10 Mio">> 10 Mio (High Loss)</option>
            <option value="Profit">Profit</option>
          </select>

          <select id="aug-filter-cek" class="filter-select" onchange="applyAugFilters()" style="border-color:rgba(6,182,212,0.4);background:rgba(6,182,212,0.08);color:#67e8f9">
            <option value="ALL">Semua Kategori (Kolom BB)</option>
            <option value="Reguler">Reguler</option>
            <option value="Repeater/CDC/IDR/SPS">Repeater/CDC/IDR/SPS</option>
            <option value="Business & Commercial">Business & Commercial</option>
            <option value="N3T">N3T</option>
            <option value="B2B">B2B</option>
            <option value="VIP">VIP</option>
            <option value="Public Open Space & Tourism">Public Open Space & Tourism</option>
            <option value="Public Transportation">Public Transportation</option>
            <option value="Reguler / CDC">Reguler / CDC</option>
            <option value="TTC/MSC/STO">TTC/MSC/STO</option>
          </select>

          <!-- Filter Kolom AG: Site Loss Tracker -->
          <select id="aug-filter-ag" class="filter-select" onchange="applyAugFilters()" style="border-color:rgba(234,179,8,0.45);background:rgba(234,179,8,0.1);color:#fde047;font-weight:600">
            <option value="ALL">Semua Site Loss Tracker (Kolom AG)</option>
            <option value="Program H2">Program H2 ({sep_meta['h2AllCount']} Site)</option>
            <option value="0">0 (Non Program H2)</option>
            <option value="#N/A">#N/A</option>
          </select>

          <!-- Fitur Sort Besar ke Kecil (Shorting) -->
          <select id="aug-sort-select" class="filter-select" onchange="onAugSortChange(this.value)" style="border-color:rgba(56,189,248,0.45);background:rgba(56,189,248,0.12);color:#38bdf8;font-weight:700">
            <option value="NONE">⚡ Urutkan: Default</option>
            <option value="DEFISIT_DESC">🔻 Defisit Terbesar (Besar ke Kecil)</option>
            <option value="DEFISIT_ASC">🔺 Defisit Terkecil (Kecil ke Besar)</option>
            <option value="COST_DESC">🔻 Cost Terbesar (Besar ke Kecil)</option>
            <option value="REV_DESC">🔻 Revenue Terbesar (Besar ke Kecil)</option>
            <option value="NILAI_COST_DESC">🔻 Nilai Cost Terbesar (Kolom BA)</option>
            <option value="SITE_ID_ASC">🔤 Site ID (A - Z)</option>
          </select>

          <select id="aug-page-size" class="filter-select" onchange="changePageSize(this.value)">
            <option value="25">25 per hal</option>
            <option value="50" selected>50 per hal</option>
            <option value="100">100 per hal</option>
            <option value="250">250 per hal</option>
          </select>
        </div>
      </div>

      <!-- Detail Table Container -->
      <div class="table-container" style="max-height:680px">
        <table class="data-table" id="aug-database-table">
          <thead>
            <tr>
              <th>#</th>
              <th>Site ID</th>
              <th>Site Name</th>
              <th>NOP</th>
              <th>Branch</th>
              <th>Kabupaten</th>
              <th>Kelas Site</th>
              <th>Status OA</th>
              <th>Type Site</th>
              <th>BTS Equip</th>
              <th id="th-db-rev">Revenue Sep-26 (Rp)</th>
              <th>Direct Cost (Rp)</th>
              <th>Margin (Rp)</th>
              <th>% Margin</th>
              <th>Status PnL</th>
              <th>GAP Dist</th>
              <th>Inline SSGJ</th>
              <th>Program Group</th>
              <th class="new-col-th" style="background:rgba(245,158,11,0.25);color:#fcd34d">Priority Sales</th>
              <th class="new-col-th" style="background:rgba(56,189,248,0.25);color:#7dd3fc">Priority Network</th>
              <th>Intervention Flag</th>
              <th id="th-db-status">Status Sep</th>
              <th onclick="sortAugTable('ag')" style="cursor:pointer;user-select:none;background:rgba(234,179,8,0.25);color:#fde047" title="Klik untuk mengurutkan Kolom AG">Site Loss Tracker (AG) <span id="th-sort-aug-ag" style="font-size:10px;color:#fde047">⬍</span></th>
              <th>Payload</th>
              <!-- NEW COLUMNS HEADERS -->
              <th class="new-col-th" style="background:rgba(239,68,68,0.2);color:#fca5a5">Action to be taken SQA</th>
              <th class="new-col-th" style="background:rgba(239,68,68,0.15);color:#fca5a5">Update Action SQA</th>
              <th class="new-col-th" style="background:rgba(245,158,11,0.2);color:#fcd34d">Action to be taken Sales</th>
              <th class="new-col-th" style="background:rgba(245,158,11,0.15);color:#fcd34d">Update Action Sales</th>
              <th class="new-col-th" style="background:rgba(16,185,129,0.2);color:#6ee7b7">Action to be taken NOP</th>
              <th class="new-col-th" style="background:rgba(16,185,129,0.15);color:#6ee7b7">Update Action NOP</th>
              <th class="new-col-th" style="background:rgba(167,139,250,0.2);color:#c4b5fd">update RHI Week Terakhir</th>
              <th class="new-col-th" style="background:rgba(167,139,250,0.15);color:#c4b5fd">Root Cause RHI</th>
              <th class="new-col-th" style="background:rgba(167,139,250,0.12);color:#c4b5fd">propose action RHI</th>
              <th class="new-col-th" style="background:rgba(56,189,248,0.2);color:#7dd3fc">update THI Week Terakhir</th>
              <th class="new-col-th" style="background:rgba(56,189,248,0.15);color:#7dd3fc">Root Cause THI</th>
              <th class="new-col-th" style="background:rgba(56,189,248,0.12);color:#7dd3fc">propose action THI</th>
              <th class="new-col-th">Typer Transport</th>
              <th class="new-col-th">utilisasi avg weekly BDBH</th>
              <th class="new-col-th">update avg status availability</th>
              <th class="new-col-th">plan action availability</th>
              <th class="new-col-th" style="background:rgba(59,130,246,0.25);color:#93c5fd">Status PL week terakhir</th>
              <th class="new-col-th" style="background:rgba(59,130,246,0.18);color:#93c5fd">propose action (MBA)</th>
              <th class="new-col-th" style="background:rgba(14,165,233,0.2);color:#7dd3fc">Breakdown cost terbesar</th>
              <th class="new-col-th" style="background:rgba(14,165,233,0.2);color:#7dd3fc">Nilai Cost</th>
              <th class="new-col-th" style="background:rgba(14,165,233,0.15);color:#7dd3fc">plan improvement cost terbesar</th>
              <th class="new-col-th" style="background:rgba(6,182,212,0.2);color:#67e8f9">Cek / Kategori Site</th>
            </tr>
          </thead>
          <tbody id="aug-table-body">
            <!-- Rows rendered by JS -->
          </tbody>
        </table>
      </div>

      <!-- Pagination Footer -->
      <div class="pagination-bar">
        <div id="aug-pagination-info">Menampilkan 1 - 50 dari {sep_meta['branchStats']['ALL']['total']} site</div>
        <div class="pagination-controls">
          <button class="page-btn" id="btn-first" onclick="goToPage(1)">« First</button>
          <button class="page-btn" id="btn-prev" onclick="prevPage()">‹ Prev</button>
          <span style="font-family:'JetBrains Mono',monospace;font-size:12px;margin:0 8px">
            Hal <span id="current-page-num" style="color:#fff;font-weight:bold">1</span> / <span id="total-pages-num">56</span>
          </span>
          <button class="page-btn" id="btn-next" onclick="nextPage()">Next ›</button>
          <button class="page-btn" id="btn-last" onclick="goToLastPage()">Last »</button>
        </div>
      </div>
    </div>

    <!-- Tab 4: Monthly Movement Non-Profit Sites -->
    <div id="tab-movement" class="tab-pane">
      <div class="card-title-bar">
        <div>
          <h3 style="font-size:15px;color:#fff">Tracking Historis Jumlah Site Loss Bulanan (Jan – Sep 2026)</h3>
          <p style="font-size:12px;color:var(--text-muted)">Rekapitulasi Non-Profit Site Movement Bulanan Regional 11 Papua Maluku</p>
        </div>
        <button class="btn-action" onclick="exportTableToCSV('movement-table', 'Monthly_Loss_Movement_2026.csv')">Export CSV</button>
      </div>

      <div class="table-container">
        <table class="data-table" id="movement-table">
          <thead>
            <tr>
              <th>Branch</th>
              <th>Target H1</th>
              <th>Jan</th>
              <th>Feb</th>
              <th>Mar</th>
              <th>Apr</th>
              <th>Mei</th>
              <th>Jun</th>
              <th>Jul</th>
              <th>Agust</th>
              <th style="background:rgba(14,165,233,0.2);color:#38bdf8">Sept</th>
              <th>Tren (Agu ke Sep)</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td><strong style="color:#f87171">AMBON</strong></td>
              <td><span class="badge badge-warning">61</span></td>
              <td>90</td>
              <td>119</td>
              <td>63</td>
              <td>104</td>
              <td>123</td>
              <td>119</td>
              <td>147</td>
              <td>145</td>
              <td><strong style="color:#38bdf8">{sep_meta['branchStats']['AMBON']['loss']}</strong></td>
              <td><span class="badge badge-success">↓ -14 (Membaik)</span></td>
            </tr>
            <tr>
              <td><strong style="color:#38bdf8">JAYAPURA</strong></td>
              <td><span class="badge badge-warning">12</span></td>
              <td>32</td>
              <td>56</td>
              <td>32</td>
              <td>48</td>
              <td>54</td>
              <td>36</td>
              <td>41</td>
              <td>48</td>
              <td><strong style="color:#38bdf8">{sep_meta['branchStats']['JAYAPURA']['loss']}</strong></td>
              <td><span class="badge badge-info">= 0 (Stabil)</span></td>
            </tr>
            <tr>
              <td><strong style="color:#10b981">SORONG</strong></td>
              <td><span class="badge badge-warning">21</span></td>
              <td>19</td>
              <td>29</td>
              <td>20</td>
              <td>24</td>
              <td>28</td>
              <td>23</td>
              <td>31</td>
              <td>25</td>
              <td><strong style="color:#38bdf8">{sep_meta['branchStats']['SORONG']['loss']}</strong></td>
              <td><span class="badge badge-success">↓ -4 (Membaik)</span></td>
            </tr>
            <tr>
              <td><strong style="color:#a855f7">TIMIKA</strong></td>
              <td><span class="badge badge-warning">20</span></td>
              <td>17</td>
              <td>23</td>
              <td>21</td>
              <td>24</td>
              <td>28</td>
              <td>20</td>
              <td>21</td>
              <td>24</td>
              <td><strong style="color:#38bdf8">{sep_meta['branchStats']['TIMIKA']['loss']}</strong></td>
              <td><span class="badge badge-success">↓ -7 (Membaik)</span></td>
            </tr>
            <tr class="total-row">
              <td><strong>TOTAL REGIONAL</strong></td>
              <td><span class="badge badge-warning">114</span></td>
              <td><strong>158</strong></td>
              <td><strong>227</strong></td>
              <td><strong>136</strong></td>
              <td><strong>200</strong></td>
              <td><strong>233</strong></td>
              <td><strong>198</strong></td>
              <td><strong>240</strong></td>
              <td><strong>242</strong></td>
              <td><strong style="color:#38bdf8">{sep_meta['branchStats']['ALL']['loss']}</strong></td>
              <td><span class="badge badge-success">↓ -25 Site (Membaik)</span></td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>

    <!-- Tab 5: Defisit Gap Analysis -->
    <div id="tab-gap" class="tab-pane">
      <div class="card-title-bar">
        <div>
          <h3 style="font-size:15px;color:#fff" id="gap-tab-title">Kategori Besaran Defisit Site Loss (September 2026)</h3>
          <p style="font-size:12px;color:var(--text-muted)">Klasifikasi besaran kerugian bulanan per site (<3 Juta s/d >10 Juta per bulan)</p>
        </div>
        <button class="btn-action" onclick="exportTableToCSV('gap-table', 'Deficit_Gap_Distribution.csv')">Export CSV</button>
      </div>

      <div class="table-container">
        <table class="data-table" id="gap-table">
          <thead>
            <tr>
              <th>Branch / NOP</th>
              <th>< 3 Juta / bln</th>
              <th>3 - 5 Juta / bln</th>
              <th>5 - 10 Juta / bln</th>
              <th>> 10 Juta / bln (High Deficit)</th>
              <th>Total Site Loss</th>
            </tr>
          </thead>
          <tbody id="gap-table-tbody">
            <tr>
              <td><strong>AMBON</strong></td>
              <td>{sep_meta['gapMatrix']['AMBON']['lt3']}</td>
              <td>{sep_meta['gapMatrix']['AMBON']['m3_5']}</td>
              <td>{sep_meta['gapMatrix']['AMBON']['m5_10']}</td>
              <td><strong style="color:#f87171">{sep_meta['gapMatrix']['AMBON']['gt10']}</strong></td>
              <td><strong>{sep_meta['gapMatrix']['AMBON']['total']}</strong></td>
            </tr>
            <tr>
              <td><strong>JAYAPURA</strong></td>
              <td>{sep_meta['gapMatrix']['JAYAPURA']['lt3']}</td>
              <td>{sep_meta['gapMatrix']['JAYAPURA']['m3_5']}</td>
              <td>{sep_meta['gapMatrix']['JAYAPURA']['m5_10']}</td>
              <td><strong style="color:#f87171">{sep_meta['gapMatrix']['JAYAPURA']['gt10']}</strong></td>
              <td><strong>{sep_meta['gapMatrix']['JAYAPURA']['total']}</strong></td>
            </tr>
            <tr>
              <td><strong>SORONG</strong></td>
              <td>{sep_meta['gapMatrix']['SORONG']['lt3']}</td>
              <td>{sep_meta['gapMatrix']['SORONG']['m3_5']}</td>
              <td>{sep_meta['gapMatrix']['SORONG']['m5_10']}</td>
              <td><strong style="color:#f87171">{sep_meta['gapMatrix']['SORONG']['gt10']}</strong></td>
              <td><strong>{sep_meta['gapMatrix']['SORONG']['total']}</strong></td>
            </tr>
            <tr>
              <td><strong>TIMIKA</strong></td>
              <td>{sep_meta['gapMatrix']['TIMIKA']['lt3']}</td>
              <td>{sep_meta['gapMatrix']['TIMIKA']['m3_5']}</td>
              <td>{sep_meta['gapMatrix']['TIMIKA']['m5_10']}</td>
              <td><strong style="color:#f87171">{sep_meta['gapMatrix']['TIMIKA']['gt10']}</strong></td>
              <td><strong>{sep_meta['gapMatrix']['TIMIKA']['total']}</strong></td>
            </tr>
            <tr class="total-row">
              <td><strong>TOTAL R11</strong></td>
              <td><strong>{sep_meta['gapMatrix']['TOTAL']['lt3']} ({sep_meta['gapMatrix']['TOTAL']['lt3_pct']}%)</strong></td>
              <td><strong>{sep_meta['gapMatrix']['TOTAL']['m3_5']} ({sep_meta['gapMatrix']['TOTAL']['m3_5_pct']}%)</strong></td>
              <td><strong>{sep_meta['gapMatrix']['TOTAL']['m5_10']} ({sep_meta['gapMatrix']['TOTAL']['m5_10_pct']}%)</strong></td>
              <td><strong style="color:#f87171">{sep_meta['gapMatrix']['TOTAL']['gt10']} ({sep_meta['gapMatrix']['TOTAL']['gt10_pct']}%)</strong></td>
              <td><strong>{sep_meta['gapMatrix']['TOTAL']['total']} Sites</strong></td>
            </tr>
          </tbody>
        </table>
      </div>

      <div class="gap-grid" id="gap-cards-container">
        <div class="gap-card">
          <div class="gap-card-header">
            <span class="gap-tier">Tier 1: < 3 Juta / bln</span>
            <span class="badge badge-info">Low Deficit</span>
          </div>
          <div class="gap-count" id="gap-card-lt3-val">{sep_meta['gapMatrix']['TOTAL']['lt3']} <span style="font-size:14px;color:var(--text-muted)">sites</span></div>
          <div class="gap-pct" id="gap-card-lt3-pct">{sep_meta['gapMatrix']['TOTAL']['lt3_pct']}% dari total site loss. Cukup butuh dorongan penetrasi paket voucher lokal.</div>
        </div>

        <div class="gap-card">
          <div class="gap-card-header">
            <span class="gap-tier">Tier 2: 3 - 5 Juta / bln</span>
            <span class="badge badge-info">Medium-Low</span>
          </div>
          <div class="gap-count" id="gap-card-35-val">{sep_meta['gapMatrix']['TOTAL']['m3_5']} <span style="font-size:14px;color:var(--text-muted)">sites</span></div>
          <div class="gap-pct" id="gap-card-35-pct">{sep_meta['gapMatrix']['TOTAL']['m3_5_pct']}% dari total site loss. Potensi recovery melalui kolaborasi Sales & RNOP.</div>
        </div>

        <div class="gap-card">
          <div class="gap-card-header">
            <span class="gap-tier">Tier 3: 5 - 10 Juta / bln</span>
            <span class="badge badge-warning">Medium-High</span>
          </div>
          <div class="gap-count" id="gap-card-510-val">{sep_meta['gapMatrix']['TOTAL']['m5_10']} <span style="font-size:14px;color:var(--text-muted)">sites</span></div>
          <div class="gap-pct" id="gap-card-510-pct">{sep_meta['gapMatrix']['TOTAL']['m5_10_pct']}% dari total site loss. Membutuhkan evaluasi biaya sewa / daya genset.</div>
        </div>

        <div class="gap-card" style="border-color:rgba(239,68,68,0.4);background:rgba(239,68,68,0.06)">
          <div class="gap-card-header">
            <span class="gap-tier" style="color:#f87171">Tier 4: > 10 Juta / bln</span>
            <span class="badge badge-danger">Critical Deficit</span>
          </div>
          <div class="gap-count" style="color:#f87171" id="gap-card-gt10-val">{sep_meta['gapMatrix']['TOTAL']['gt10']} <span style="font-size:14px;color:var(--text-muted)">sites</span></div>
          <div class="gap-pct" id="gap-card-gt10-pct">{sep_meta['gapMatrix']['TOTAL']['gt10_pct']}% dari total loss! Menyerap kerugian terbesar (site repeater satelit/DMT/TBG).</div>
        </div>
      </div>
    </div>

    <!-- Tab 6: Pembanding Nasional -->
    <div id="tab-regional" class="tab-pane">
      <div class="card-title-bar">
        <div>
          <h3 style="font-size:15px;color:#fff">Pembanding Performa Lintas Area Nasional (Juli 2026 Baseline)</h3>
          <p style="font-size:12px;color:var(--text-muted)">Perbandingan rasio profitabilitas antar Area 1, 2, 3, dan 4 dari Pivot Master</p>
        </div>
      </div>

      <div class="table-container">
        <table class="data-table">
          <thead>
            <tr>
              <th>Area / Regional</th>
              <th>Loss Site</th>
              <th>Profit Site</th>
              <th>Total Site Terdata</th>
              <th>% Profit Margin</th>
              <th>Catatan Regional</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td><strong>AREA 1 (Sumatera)</strong></td>
              <td>9</td>
              <td>172</td>
              <td>181</td>
              <td><span class="badge badge-success">95.0%</span></td>
              <td>R01 Sumbagut, R02 Sumbagsel, R10 Sumbagteng</td>
            </tr>
            <tr>
              <td><strong>AREA 2 (Jabotabek & Jabar)</strong></td>
              <td>10</td>
              <td>288</td>
              <td>298</td>
              <td><span class="badge badge-success">96.6%</span></td>
              <td>R03 Jakarta Banten, R04 Jabar, R12 Eastern Jabotabek</td>
            </tr>
            <tr>
              <td><strong>AREA 3 (Jawa Bali Nusra)</strong></td>
              <td>1</td>
              <td>56</td>
              <td>57</td>
              <td><span class="badge badge-success">98.2%</span></td>
              <td>R05 Jateng, R06 Jatim, R07 Bali Nusra</td>
            </tr>
            <tr style="background:rgba(255,42,68,0.08)">
              <td><strong>AREA 4 (Pamasuka)</strong></td>
              <td><strong>204</strong></td>
              <td><strong>2,644</strong></td>
              <td><strong>2,848</strong></td>
              <td><span class="badge badge-warning">92.8%</span></td>
              <td>R08 Kalimantan (2 Loss), R09 Sulawesi (4 Loss), R11 Papua Maluku (198 Loss)</td>
            </tr>
            <tr class="total-row">
              <td><strong>TOTAL SAMPEL NASIONAL</strong></td>
              <td><strong>224</strong></td>
              <td><strong>3,160</strong></td>
              <td><strong>3,384</strong></td>
              <td><span class="badge badge-success">93.4%</span></td>
              <td>Konsentrasi site loss nasional terpusat di wilayah timur (Area 4 / R11)</td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  </div>

  <!-- Strategic Insights Card -->
  <div class="insights-card glass">
    <div class="insights-title">
      <svg width="20" height="20" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"></path></svg>
      <span id="insights-main-title">Executive Takeaways & Rencana Aksi Tim SQA PUMA (Update September 2026)</span>
    </div>
    <ul class="insights-list" id="insights-list-body">
      <li>
        <strong>Tiga Branch Berhasil Melampaui Target 93% (September 2026):</strong> NOP Timika (<strong>97.02%</strong>), NOP Sorong (<strong>96.79%</strong>), dan NOP Jayapura (<strong>93.21%</strong>) konsisten melampaui target H2. Total site loss Regional turun signifikan <strong>-25 site</strong> (dari 242 di Agustus menjadi <strong>217 site loss</strong> di September, capaian R11 naik ke <strong>92.22%</strong>).
      </li>
      <li>
        <strong>Konsentrasi Site Loss per Kabupaten (September 2026):</strong>
        <br>• <em>Branch Ambon (131 Loss):</em> <strong>Maluku Tengah (36 site loss)</strong>, <strong>Kota Ambon (25 site loss)</strong>, <strong>Seram Bagian Timur (16 site loss)</strong>, <strong>Maluku Tenggara (15 site loss)</strong>, dan <strong>Buru (13 site loss)</strong>.
        <br>• <em>Branch Jayapura (48 Loss):</em> <strong>Biak Numfor (15 site loss)</strong>, <strong>Kabupaten Jayapura (9 site loss)</strong>, dan <strong>Kota Jayapura (7 site loss)</strong>.
        <br>• <em>Branch Sorong (21 Loss):</em> <strong>Manokwari (7 site loss)</strong>, <strong>Fakfak (3 site loss)</strong>, dan <strong>Kota Sorong (3 site loss)</strong>.
        <br>• <em>Branch Timika (17 Loss):</em> <strong>Mimika (6 site loss)</strong>, <strong>Mappi (4 site loss)</strong>, dan <strong>Merauke (4 site loss)</strong>.
      </li>
      <li>
        <strong>Switch Periode & Action Tracking Real-Time:</strong> Gunakan tombol <strong>Periode Data (Agustus '26 / September '26)</strong> di bagian atas untuk membandingkan data bulan Agustus dan September secara instan. Seluruh keterangan lapangan tersinkronisasi langsung dengan sheet <code>SEP '26</code> dan <code>Aug 26</code>.
      </li>
    </ul>
  </div>

</div>

<!-- Tooltip element -->
<div id="tooltip" class="custom-tooltip"></div>

<script>
  const periodDatasets = {json_period_datasets};
  let activePeriod = 'SEP26';
  let augData = periodDatasets.SEP26.rows;
  let kabData = periodDatasets.SEP26.kabData;
  let branchKabSummary = periodDatasets.SEP26.branchKabSummary;

  const monthsAug = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agust'];
  const monthsSep = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agust', 'Sept'];
  let months = monthsSep;

  const trendDataAug = {{
    TOTAL: [93.97, 91.37, 94.84, 92.43, 91.06, 92.74, 91.24, 91.33],
    AMBON: [88.81, 82.83, 89.72, 84.76, 82.47, 85.61, 82.22, 82.94],
    JAYAPURA: [95.24, 89.45, 92.92, 90.75, 90.03, 94.80, 94.08, 93.23],
    SORONG: [96.82, 91.24, 92.99, 92.36, 91.72, 96.34, 94.90, 96.19],
    TIMIKA: [96.90, 92.46, 93.33, 92.81, 92.28, 96.49, 96.14, 95.83]
  }};
  const trendDataSep = {{
    TOTAL: [93.97, 91.37, 94.84, 92.43, 91.06, 92.74, 91.24, 91.33, periodDatasets.SEP26.branchStats.ALL.pPctNum],
    AMBON: [88.81, 82.83, 89.72, 84.76, 82.47, 85.61, 82.22, 82.94, periodDatasets.SEP26.branchStats.AMBON.pPctNum],
    JAYAPURA: [95.24, 89.45, 92.92, 90.75, 90.03, 94.80, 94.08, 93.23, periodDatasets.SEP26.branchStats.JAYAPURA.pPctNum],
    SORONG: [96.82, 91.24, 92.99, 92.36, 91.72, 96.34, 94.90, 96.19, periodDatasets.SEP26.branchStats.SORONG.pPctNum],
    TIMIKA: [96.90, 92.46, 93.33, 92.81, 92.28, 96.49, 96.14, 95.83, periodDatasets.SEP26.branchStats.TIMIKA.pPctNum]
  }};
  let trendData = trendDataSep;

  const branchColorsDark = {{
    TOTAL: '#ff2a44',
    AMBON: '#f59e0b',
    JAYAPURA: '#38bdf8',
    SORONG: '#10b981',
    TIMIKA: '#a855f7'
  }};
  const branchColorsLight = {{
    TOTAL: '#e11d48',
    AMBON: '#d97706',
    JAYAPURA: '#0284c7',
    SORONG: '#059669',
    TIMIKA: '#7c3aed'
  }};

  function isLightTheme() {{
    return document.documentElement.getAttribute('data-theme') === 'light';
  }}

  let activeBranch = 'ALL';

  function renderTrendChart() {{
    const container = document.getElementById('trend-chart-container');
    const width = container.clientWidth || 800;
    const height = 310;
    const padX = 55;
    const padY = 35;
    const chartW = width - padX * 2;
    const chartH = height - padY * 2;
    const lightMode = isLightTheme();
    const palette = lightMode ? branchColorsLight : branchColorsDark;
    const targetColor = lightMode ? '#d97706' : '#fbbf24';
    const pointStroke = lightMode ? '#ffffff' : '#090d16';

    const minVal = 80;
    const maxVal = 100;

    function getX(idx) {{
      return padX + (idx / (months.length - 1)) * chartW;
    }}

    function getY(val) {{
      return padY + chartH - ((val - minVal) / (maxVal - minVal)) * chartH;
    }}

    let svg = `<svg class="chart-svg" viewBox="0 0 ${{width}} ${{height}}">`;

    for (let v = 80; v <= 100; v += 5) {{
      const y = getY(v);
      svg += `<line x1="${{padX}}" y1="${{y}}" x2="${{padX + chartW}}" y2="${{y}}" class="grid-line" />`;
      svg += `<text x="${{padX - 10}}" y="${{y + 4}}" class="axis-label" text-anchor="end">${{v}}%</text>`;
    }}

    const yTarget = getY(93);
    svg += `<line x1="${{padX}}" y1="${{yTarget}}" x2="${{padX + chartW}}" y2="${{yTarget}}" stroke="${{targetColor}}" stroke-width="1.8" stroke-dasharray="6,4" opacity="${{lightMode ? '0.9' : '0.8'}}" />`;
    svg += `<text x="${{padX + chartW + 5}}" y="${{yTarget + 4}}" class="axis-label" fill="${{targetColor}}" font-weight="bold">Target 93%</text>`;

    months.forEach((m, i) => {{
      const x = getX(i);
      svg += `<text x="${{x}}" y="${{padY + chartH + 20}}" class="axis-label" text-anchor="middle">${{m}}</text>`;
    }});

    const branchesToDraw = activeBranch === 'ALL' 
      ? ['AMBON', 'JAYAPURA', 'SORONG', 'TIMIKA', 'TOTAL'] 
      : [activeBranch];

    branchesToDraw.forEach(b => {{
      const vals = trendData[b];
      const color = palette[b];
      const isTotal = (b === 'TOTAL');

      let pathD = '';
      let areaD = `M ${{getX(0)}} ${{padY + chartH}} `;

      vals.forEach((v, i) => {{
        const x = getX(i);
        const y = getY(v);
        if (i === 0) {{
          pathD += `M ${{x}} ${{y}} `;
          areaD += `L ${{x}} ${{y}} `;
        }} else {{
          pathD += `L ${{x}} ${{y}} `;
          areaD += `L ${{x}} ${{y}} `;
        }}
      }});

      areaD += `L ${{getX(vals.length - 1)}} ${{padY + chartH}} Z`;

      if (activeBranch !== 'ALL' || isTotal) {{
        svg += `<path d="${{areaD}}" fill="${{color}}" opacity="${{isTotal ? (lightMode ? '0.09' : '0.12') : (lightMode ? '0.14' : '0.2')}}" />`;
      }}

      svg += `<path d="${{pathD}}" stroke="${{color}}" class="chart-line" stroke-width="${{isTotal ? 4 : 2.5}}" />`;

      vals.forEach((v, i) => {{
        const x = getX(i);
        const y = getY(v);
        svg += `<circle cx="${{x}}" cy="${{y}}" r="${{isTotal ? 5 : 4}}" fill="${{color}}" stroke="${{pointStroke}}" stroke-width="2" class="data-point" 
                  onmouseover="showTooltip(event, '${{b}}', '${{months[i]}}', '${{v.toFixed(2)}}%')" 
                  onmouseout="hideTooltip()" />`;
      }});
    }});

    svg += `</svg>`;
    container.innerHTML = svg;
  }}

  // Tooltip
  const tooltip = document.getElementById('tooltip');
  function showTooltip(e, branch, month, value) {{
    tooltip.innerHTML = `<strong>${{branch}}</strong><br>${{month}} 2026: <span style="color:#34d399;font-weight:bold">${{value}}</span>`;
    tooltip.style.left = (e.pageX + 12) + 'px';
    tooltip.style.top = (e.pageY - 35) + 'px';
    tooltip.style.opacity = '1';
  }}

  function hideTooltip() {{
    tooltip.style.opacity = '0';
  }}

  // Switch Main Tabs
  function switchTab(tabId) {{
    document.querySelectorAll('.tabs-nav .tab-btn').forEach(btn => btn.classList.remove('active'));
    document.querySelectorAll('.tab-pane').forEach(p => p.classList.remove('active'));
    event.currentTarget.classList.add('active');
    document.getElementById(tabId).classList.add('active');

    if (tabId === 'tab-kabupaten') {{
      renderKabTable();
      renderProgTable();
    }} else if (tabId === 'tab-aug-db') {{
      renderAugTable();
    }}
  }}

  // Switch Sub-Tabs inside Branch Focus
  function switchBranchSubTab(subTabId) {{
    document.querySelectorAll('.subtabs-bar .subtab-btn').forEach(btn => btn.classList.remove('active'));
    document.querySelectorAll('.branch-subtab-pane').forEach(p => p.classList.remove('active'));
    event.currentTarget.classList.add('active');
    document.getElementById(subTabId).classList.add('active');

    if (subTabId === 'subtab-kab-distribusi') {{
      renderKabTable();
    }} else if (subTabId === 'subtab-prog-tracker') {{
      renderProgTable();
    }}
  }}

  // --- DYNAMIC SIDE CARD (PER KOTA/KABUPATEN) ---
  function updateSideCard(branch) {{
    const titleElem = document.getElementById('side-card-title');
    const badgeElem = document.getElementById('side-card-badge');
    const contentElem = document.getElementById('side-card-content');
    const alertTitle = document.getElementById('side-alert-title');
    const alertDesc = document.getElementById('side-alert-desc');

    const summary = branchKabSummary[branch] || branchKabSummary['ALL'];
    const totalLoss = summary.total_loss;
    const isSep = (activePeriod === 'SEP26');
    const curStats = periodDatasets[activePeriod].branchStats;

    if (branch === 'ALL') {{
      titleElem.innerText = `Top Kota/Kabupaten Site Loss (${{isSep ? 'Sep 26' : 'Agu 26'}})`;
      badgeElem.innerText = `Total: ${{totalLoss}} Sites`;
      alertTitle.innerText = `⚠️ Top Critical Kabupaten se-Regional (${{isSep ? 'September 2026' : 'Agustus 2026'}})`;
      const top4 = summary.kabs.slice(0, 4).map(k => `<strong>${{k.name}} (${{k.loss}} site)</strong>`).join(', ');
      alertDesc.innerHTML = `Konsentrasi site loss terbesar berada di ${{top4}}.`;
    }} else {{
      titleElem.innerText = `Sebaran Site Loss Kota/Kabupaten — Branch ${{branch}} (${{isSep ? 'Sep 26' : 'Agu 26'}})`;
      badgeElem.innerText = `Total: ${{totalLoss}} Sites Loss`;
      const bStat = curStats[branch];
      const top2 = summary.kabs.slice(0, 2).map(k => `<strong>${{k.name}} (${{k.loss}} site)</strong>`).join(' dan ');
      if (branch === 'AMBON') {{
        alertTitle.innerText = '⚠️ Focus Area: NOP Ambon';
        alertDesc.innerHTML = `Ambon memiliki <strong>${{bStat.loss}} site loss</strong> (achv ${{bStat.pPct}}). Terkonsentrasi di ${{top2}}. Perlu rescue <strong>${{Math.abs(bStat.gap_num)}} site</strong> untuk capai target 93%.`;
      }} else {{
        alertTitle.innerText = `🎯 Focus Area: Branch ${{branch}}`;
        alertDesc.innerHTML = `${{branch}} memiliki <strong>${{bStat.loss}} site loss</strong> (achv ${{bStat.pPct}} - Target Tercapai ✓). Konsentrasi loss utama di ${{top2}}.`;
      }}
    }}

    let html = '';
    const colors = isLightTheme()
      ? ['#dc2626', '#e11d48', '#ea580c', '#d97706', '#0284c7', '#059669', '#7c3aed', '#475569']
      : ['#ef4444', '#f87171', '#fb923c', '#f59e0b', '#38bdf8', '#34d399', '#a855f7', '#64748b'];

    summary.kabs.slice(0, 7).forEach((k, idx) => {{
      const c = colors[idx % colors.length];
      html += `<div>
        <div style="display:flex;justify-content:space-between;font-size:12px;margin-bottom:4px;">
          <span style="font-weight:700;color:${{c}}">${{k.name}} (${{k.loss}} site)</span>
          <span style="color:var(--text-muted);font-family:'JetBrains Mono'">${{k.pct_of_branch_loss}}% dari loss</span>
        </div>
        <div style="height:7px;background:rgba(255,255,255,0.06);border-radius:4px;overflow:hidden">
          <div style="width:${{k.pct_of_branch_loss}}%;height:100%;background:${{c}};border-radius:4px;"></div>
        </div>
      </div>`;
    }});

    contentElem.innerHTML = html;
  }}

  function renderPeriodStaticTables() {{
    const isSep = (activePeriod === 'SEP26');
    const ds = periodDatasets[activePeriod];
    const bs = ds.branchStats;
    const gm = ds.gapMatrix;
    const pLabel = isSep ? 'September 2026' : 'Agustus 2026';
    const sLabel = isSep ? "Sheet SEP '26" : 'Sheet Aug 26';

    // 1. Target vs Realisasi Table
    const tTitle = document.getElementById('target-tab-title');
    if (tTitle) tTitle.innerText = `Tabel Pencapaian & Gap Recovery Target H2 (93.0%) — ${{pLabel}}`;
    const tDesc = document.getElementById('target-tab-desc');
    if (tDesc) tDesc.innerText = `Evaluasi pencapaian profitabilitas per Branch berdasarkan ${{sLabel}}. Seluruh target site dan gap dibulatkan.`;

    const tBody = document.getElementById('target-table-tbody');
    if (tBody) {{
      const brConfigs = [
        {{ key: 'AMBON', color: '#f87171' }},
        {{ key: 'JAYAPURA', color: '#38bdf8' }},
        {{ key: 'SORONG', color: '#10b981' }},
        {{ key: 'TIMIKA', color: '#a855f7' }}
      ];
      let rowsHtml = '';
      brConfigs.forEach(bc => {{
        const st = bs[bc.key];
        const achieved = st.gap_num >= 0;
        rowsHtml += `<tr>
          <td><strong style="color:${{bc.color}}">${{bc.key}}</strong></td>
          <td>${{st.total}}</td>
          <td>${{st.profit}}</td>
          <td>${{st.loss}}</td>
          <td><strong>${{st.target}}</strong></td>
          <td><span class="badge badge-warning">93.0%</span></td>
          <td><strong style="color:${{achieved ? '#34d399' : '#f87171'}}">${{st.pPct}}</strong></td>
          <td><span class="badge ${{achieved ? 'badge-success' : 'badge-danger'}}">${{st.gap}} Site</span></td>
          <td><span class="badge ${{achieved ? 'badge-success' : 'badge-danger'}}">${{achieved ? 'Target Tercapai ✓' : 'Belum Capai Target'}}</span></td>
        </tr>`;
      }});
      const stAll = bs.ALL;
      rowsHtml += `<tr class="total-row">
        <td><strong>TOTAL R11 (PAPUA MALUKU)</strong></td>
        <td><strong>${{stAll.total}}</strong></td>
        <td><strong>${{stAll.profit}}</strong></td>
        <td><strong>${{stAll.loss}}</strong></td>
        <td><strong>${{stAll.target}}</strong></td>
        <td><span class="badge badge-warning">93.0%</span></td>
        <td><strong style="color:#fbbf24">${{stAll.pPct}}</strong></td>
        <td><span class="badge ${{stAll.gap_num >= 0 ? 'badge-success' : 'badge-danger'}}">${{stAll.gap}} Site</span></td>
        <td><span class="badge badge-warning">Gap Defisit: ${{Math.abs(stAll.gap_num)}} Site</span></td>
      </tr>`;
      tBody.innerHTML = rowsHtml;
    }}

    // 2. Gap Matrix Table & Cards
    const gTitle = document.getElementById('gap-tab-title');
    if (gTitle) gTitle.innerText = `Kategori Besaran Defisit Site Loss (${{pLabel}})`;
    const gBody = document.getElementById('gap-table-tbody');
    if (gBody) {{
      let gHtml = '';
      ['AMBON', 'JAYAPURA', 'SORONG', 'TIMIKA'].forEach(br => {{
        const g = gm[br];
        gHtml += `<tr>
          <td><strong>${{br}}</strong></td>
          <td>${{g.lt3}}</td>
          <td>${{g.m3_5}}</td>
          <td>${{g.m5_10}}</td>
          <td><strong style="color:#f87171">${{g.gt10}}</strong></td>
          <td><strong>${{g.total}}</strong></td>
        </tr>`;
      }});
      const gt = gm.TOTAL;
      gHtml += `<tr class="total-row">
        <td><strong>TOTAL R11</strong></td>
        <td><strong>${{gt.lt3}} (${{gt.lt3_pct}}%)</strong></td>
        <td><strong>${{gt.m3_5}} (${{gt.m3_5_pct}}%)</strong></td>
        <td><strong>${{gt.m5_10}} (${{gt.m5_10_pct}}%)</strong></td>
        <td><strong style="color:#f87171">${{gt.gt10}} (${{gt.gt10_pct}}%)</strong></td>
        <td><strong>${{gt.total}} Sites</strong></td>
      </tr>`;
      gBody.innerHTML = gHtml;

      const elLt3Val = document.getElementById('gap-card-lt3-val');
      if (elLt3Val) elLt3Val.innerHTML = `${{gt.lt3}} <span style="font-size:14px;color:var(--text-muted)">sites</span>`;
      const elLt3Pct = document.getElementById('gap-card-lt3-pct');
      if (elLt3Pct) elLt3Pct.innerText = `${{gt.lt3_pct}}% dari total site loss. Cukup butuh dorongan penetrasi paket voucher lokal.`;

      const el35Val = document.getElementById('gap-card-35-val');
      if (el35Val) el35Val.innerHTML = `${{gt.m3_5}} <span style="font-size:14px;color:var(--text-muted)">sites</span>`;
      const el35Pct = document.getElementById('gap-card-35-pct');
      if (el35Pct) el35Pct.innerText = `${{gt.m3_5_pct}}% dari total site loss. Potensi recovery melalui kolaborasi Sales & RNOP.`;

      const el510Val = document.getElementById('gap-card-510-val');
      if (el510Val) el510Val.innerHTML = `${{gt.m5_10}} <span style="font-size:14px;color:var(--text-muted)">sites</span>`;
      const el510Pct = document.getElementById('gap-card-510-pct');
      if (el510Pct) el510Pct.innerText = `${{gt.m5_10_pct}}% dari total site loss. Membutuhkan evaluasi biaya sewa / daya genset.`;

      const elGt10Val = document.getElementById('gap-card-gt10-val');
      if (elGt10Val) elGt10Val.innerHTML = `${{gt.gt10}} <span style="font-size:14px;color:var(--text-muted)">sites</span>`;
      const elGt10Pct = document.getElementById('gap-card-gt10-pct');
      if (elGt10Pct) elGt10Pct.innerText = `${{gt.gt10_pct}}% dari total loss! Menyerap kerugian terbesar (site repeater satelit/DMT/TBG).`;
    }}

    // 3. Database Tab Labels & Dropdowns
    const tabBtnTarget = document.getElementById('tab-btn-target');
    if (tabBtnTarget) tabBtnTarget.innerText = `🎯 Target vs Realisasi ${{pLabel}}`;
    const tabBtnDbLabel = document.getElementById('tab-btn-db-label');
    if (tabBtnDbLabel) tabBtnDbLabel.innerText = `📋 Database Site ${{sLabel}}`;
    const dbTitle = document.getElementById('db-tab-title');
    if (dbTitle) dbTitle.innerText = `Database Site ${{sLabel}} (Lengkap 58 Kolom)`;
    const dbBadge = document.getElementById('db-tab-badge');
    if (dbBadge) dbBadge.innerText = `${{bs.ALL.total}} Sites`;
    const dbDesc = document.getElementById('db-tab-desc');
    if (dbDesc) dbDesc.innerHTML = `Menampilkan seluruh data site ${{sLabel}} beserta <strong>Kelas Site</strong>, <strong>Status OA</strong>, <strong>Action SQA, Sales & NOP</strong>, performa THI, Transport, BDBH, Availability, dan Status PL.`;
    const thRev = document.getElementById('th-db-rev');
    if (thRev) thRev.innerText = isSep ? 'Revenue Sep-26 (Rp)' : 'Revenue Aug-26 (Rp)';
    const thSt = document.getElementById('th-db-status');
    if (thSt) thSt.innerText = isSep ? 'Status Sep' : 'Status Aug';

    // Update NOP, PnL, OA, AG dropdown option texts
    const progPnl = document.getElementById('prog-filter-pnl');
    if (progPnl && progPnl.options.length >= 3) {{
      progPnl.options[0].text = `Loss Sites Only (${{bs.ALL.loss}})`;
      progPnl.options[2].text = `Profit Sites Only (${{bs.ALL.profit}})`;
    }}
    const progAg = document.getElementById('prog-filter-ag');
    if (progAg && progAg.options.length >= 2) {{
      progAg.options[1].text = `Program H2 (${{ds.h2LossCount}} Site Loss)`;
    }}
    const augNop = document.getElementById('aug-filter-nop');
    if (augNop && augNop.options.length >= 5) {{
      augNop.options[1].text = `NOP AMBON (${{ds.nopCounts['NOP AMBON']}})`;
      augNop.options[2].text = `NOP JAYAPURA (${{ds.nopCounts['NOP JAYAPURA']}})`;
      augNop.options[3].text = `NOP SORONG (${{ds.nopCounts['NOP SORONG']}})`;
      augNop.options[4].text = `NOP TIMIKA (${{ds.nopCounts['NOP TIMIKA']}})`;
    }}
    const augPnl = document.getElementById('aug-filter-pnl');
    if (augPnl && augPnl.options.length >= 3) {{
      augPnl.options[1].text = `Loss Only (${{bs.ALL.loss}})`;
      augPnl.options[2].text = `Profit Only (${{bs.ALL.profit}})`;
    }}
    const augOa = document.getElementById('aug-filter-oa');
    if (augOa && augOa.options.length >= 3) {{
      augOa.options[1].text = `OA (${{ds.oaCount.toLocaleString()}})`;
      augOa.options[2].text = `Not OA (${{ds.notOaCount.toLocaleString()}})`;
    }}
    const augAg = document.getElementById('aug-filter-ag');
    if (augAg && augAg.options.length >= 2) {{
      augAg.options[1].text = `Program H2 (${{ds.h2AllCount}} Site)`;
    }}
  }}

  function switchPeriod(periodKey, silent = false) {{
    if (!periodDatasets[periodKey]) return;
    activePeriod = periodKey;
    const isSep = (periodKey === 'SEP26');
    const ds = periodDatasets[periodKey];

    augData = ds.rows;
    kabData = ds.kabData;
    branchKabSummary = ds.branchKabSummary;
    months = isSep ? monthsSep : monthsAug;
    trendData = isSep ? trendDataSep : trendDataAug;

    // Update Period Switcher Buttons
    const btnAug = document.getElementById('btn-period-aug');
    const btnSep = document.getElementById('btn-period-sep');
    if (btnAug) btnAug.classList.toggle('active', !isSep);
    if (btnSep) btnSep.classList.toggle('active', isSep);

    // Update Header Badge & Trend Chart Title
    const hdrText = document.getElementById('header-baseline-text');
    if (hdrText) hdrText.innerText = isSep ? 'BASELINE SEPTEMBER 2026' : 'BASELINE AGUSTUS 2026';
    const trTitle = document.getElementById('trend-chart-title');
    if (trTitle) trTitle.innerText = isSep
      ? 'Tren % Profit Margin (Januari – September 2026)'
      : 'Tren % Profit Margin (Januari – Agustus 2026)';

    // Update Branch Filter Pills Counts
    const bs = ds.branchStats;
    const pillAll = document.getElementById('branch-pill-ALL');
    if (pillAll) pillAll.innerText = `All Papua Maluku (${{bs.ALL.total}})`;
    const pillAmb = document.getElementById('branch-pill-AMBON');
    if (pillAmb) pillAmb.innerText = `Ambon (${{bs.AMBON.total}})`;
    const pillJay = document.getElementById('branch-pill-JAYAPURA');
    if (pillJay) pillJay.innerText = `Jayapura (${{bs.JAYAPURA.total}})`;
    const pillSor = document.getElementById('branch-pill-SORONG');
    if (pillSor) pillSor.innerText = `Sorong (${{bs.SORONG.total}})`;
    const pillTim = document.getElementById('branch-pill-TIMIKA');
    if (pillTim) pillTim.innerText = `Timika (${{bs.TIMIKA.total}})`;

    // Update KPI Card Titles & Subtexts
    const kpiProfTitle = document.getElementById('kpi-profit-title');
    if (kpiProfTitle) kpiProfTitle.innerText = isSep ? 'Profit Sites (Sep 26)' : 'Profit Sites (Aug 26)';
    const kpiLossTitle = document.getElementById('kpi-loss-title');
    if (kpiLossTitle) kpiLossTitle.innerText = isSep ? 'Non-Profit / Loss Sites (Sep 26)' : 'Non-Profit / Loss Sites (Aug 26)';
    const kpiTotSub = document.getElementById('kpi-total-sub');
    if (kpiTotSub) {{
      kpiTotSub.innerHTML = isSep
        ? `<span class="trend-up" style="color:#38bdf8">Sheet SEP '26</span> dibanding Agustus 2026 (${{periodDatasets.AUG26.branchStats.ALL.total}})`
        : `<span class="trend-up">↑ +51</span> dibanding Juli 2026 (2,740)`;
    }}
    const kpiProfSub = document.getElementById('kpi-profit-sub-label');
    if (kpiProfSub) kpiProfSub.innerText = isSep ? 'dari populasi (↑ +22 site vs Agu)' : 'dari total populasi';
    const kpiLossSub = document.getElementById('kpi-loss-sub-label');
    if (kpiLossSub) kpiLossSub.innerText = isSep ? 'rasio loss (↓ -25 site membaik vs Agu)' : 'ratio kerugian';

    renderPeriodStaticTables();
    setBranchFilter(activeBranch, true);
    applyProgFilters();
    applyAugFilters();

    try {{
      const u = new URL(window.location.href);
      u.searchParams.set('period', isSep ? 'sep' : 'aug');
      window.history.replaceState({{}}, '', u.toString());
    }} catch (e) {{}}

    if (!silent) {{
      showToast(
        isSep
          ? `📅 Beralih ke Data Terbaru: September 2026 (Sheet SEP '26 — ${{bs.ALL.loss}} Site Loss)`
          : `📅 Beralih ke Data Historis: Agustus 2026 (Sheet Aug 26 — ${{bs.ALL.loss}} Site Loss)`,
        true
      );
    }}
  }}

  // Filter Pills (Top Bar)
  function setBranchFilter(branch, skipPillEvent = false) {{
    activeBranch = branch;
    document.querySelectorAll('.filter-pill').forEach(pill => pill.classList.remove('active'));
    const targetPill = document.getElementById(`branch-pill-${{branch}}`);
    if (targetPill) targetPill.classList.add('active');
    else if (!skipPillEvent && typeof event !== 'undefined' && event && event.currentTarget) {{
      event.currentTarget.classList.add('active');
    }}

    const kpiSites = document.getElementById('kpi-total-sites');
    const kpiProfit = document.getElementById('kpi-profit-sites');
    const kpiLoss = document.getElementById('kpi-loss-sites');
    const kpiGap = document.getElementById('kpi-gap-sites');
    const profitPct = document.getElementById('kpi-profit-pct');
    const lossPct = document.getElementById('kpi-loss-pct');
    const profitBar = document.getElementById('kpi-profit-bar');
    const lossBar = document.getElementById('kpi-loss-bar');

    const branchStats = periodDatasets[activePeriod].branchStats;
    const cur = branchStats[branch] || branchStats.ALL;
    kpiSites.innerHTML = `${{cur.total}} <span class="unit">sites</span>`;
    kpiProfit.innerHTML = `${{cur.profit}} <span class="unit">sites</span>`;
    kpiLoss.innerHTML = `${{cur.loss}} <span class="unit">sites</span>`;
    kpiGap.innerHTML = `${{cur.gap}} <span class="unit">sites</span>`;
    profitPct.innerText = cur.pPct;
    lossPct.innerText = cur.lPct;
    profitBar.style.width = cur.pPct;
    lossBar.style.width = cur.lPct;

    updateSideCard(branch);
    renderPriorityCharts();

    const kabSelect = document.getElementById('kab-select-branch');
    if (kabSelect) {{
      kabSelect.value = branch;
      renderKabTable();
    }}

    const progNop = document.getElementById('prog-filter-nop');
    if (progNop) {{
      progNop.value = branch;
      applyProgFilters();
    }}

    const nopSelect = document.getElementById('aug-filter-nop');
    if (nopSelect) {{
      if (branch === 'ALL') nopSelect.value = 'ALL';
      else if (branch === 'AMBON') nopSelect.value = 'NOP AMBON';
      else if (branch === 'JAYAPURA') nopSelect.value = 'NOP JAYAPURA';
      else if (branch === 'SORONG') nopSelect.value = 'NOP SORONG';
      else if (branch === 'TIMIKA') nopSelect.value = 'NOP TIMIKA';
      applyAugFilters();
    }}

    renderTrendChart();
  }}

  // --- KABUPATEN TAB LOGIC ---
  function onKabBranchChange(val) {{
    renderKabTable();
  }}

  function renderKabTable() {{
    const branchVal = document.getElementById('kab-select-branch').value;
    const q = (document.getElementById('kab-search-input').value || '').toLowerCase().trim();
    const lossOnly = document.getElementById('kab-filter-loss-only').checked;
    const badge = document.getElementById('kab-header-badge');
    badge.innerText = branchVal === 'ALL' ? 'Semua Branch' : `Branch ${{branchVal}}`;

    const filteredKabs = kabData.filter(k => {{
      if (branchVal !== 'ALL' && k.branch !== branchVal) return false;
      if (lossOnly && k.loss === 0) return false;
      if (q && !k.kabupaten.toLowerCase().includes(q)) return false;
      return true;
    }});

    const tbody = document.getElementById('kab-table-body');
    if (filteredKabs.length === 0) {{
      tbody.innerHTML = `<tr><td colspan="13" style="text-align:center;padding:35px;color:var(--text-muted)">Tidak ada Kabupaten/Kota yang sesuai dengan kriteria filter</td></tr>`;
      return;
    }}

    let html = '';
    let totSites = 0, totProfit = 0, totLoss = 0, totLt3 = 0, tot3_5 = 0, tot5_10 = 0, totGt10 = 0, totMargin = 0;

    filteredKabs.forEach((k, idx) => {{
      totSites += k.total;
      totProfit += k.profit;
      totLoss += k.loss;
      totLt3 += k.tier_lt3;
      tot3_5 += k.tier_3_5;
      tot5_10 += k.tier_5_10;
      totGt10 += k.tier_gt10;
      totMargin += k.margin;

      const lossPctColor = k.pct_loss > 20 ? '#ef4444' : (k.pct_loss > 10 ? '#f59e0b' : '#38bdf8');
      const barFillColor = k.pct_loss > 20 ? '#ef4444' : (k.pct_loss > 10 ? '#f59e0b' : '#10b981');

      html += `<tr>
        <td style="color:var(--text-dim);font-family:'JetBrains Mono'">${{idx + 1}}</td>
        <td><strong style="color:#fff">${{k.kabupaten}}</strong></td>
        <td><span class="badge" style="background:rgba(255,255,255,0.06);color:#94a3b8">${{k.branch}}</span></td>
        <td style="font-weight:600">${{k.total}}</td>
        <td style="font-weight:600;color:#34d399">${{k.profit}}</td>
        <td style="font-weight:700;color:#f87171">${{k.loss}}</td>
        <td>
          <span style="font-weight:700;color:${{lossPctColor}}">${{k.pct_loss}}%</span>
          <span class="mini-bar-bg"><span class="mini-bar-fill" style="width:${{k.pct_loss}}%;background:${{barFillColor}}"></span></span>
        </td>
        <td style="font-weight:600">${{k.tier_lt3 || '-'}}</td>
        <td style="font-weight:600">${{k.tier_3_5 || '-'}}</td>
        <td style="font-weight:600">${{k.tier_5_10 || '-'}}</td>
        <td style="font-weight:700;color:${{k.tier_gt10 > 0 ? '#f87171' : 'inherit'}}">${{k.tier_gt10 || '-'}}</td>
        <td style="text-align:right;font-weight:600;color:${{k.margin < 0 ? '#f87171' : '#34d399'}}">${{Math.round(k.margin).toLocaleString()}}</td>
        <td>
          <button class="page-btn" onclick="openKabDrilldown('${{k.kabupaten}}', '${{k.branch}}')" style="padding:4px 10px;font-size:11.5px;background:rgba(255,42,68,0.15);border-color:rgba(255,42,68,0.35);color:#ff6b7d">
            🔍 Detail Site Loss (${{k.loss}})
          </button>
        </td>
      </tr>`;
    }});

    // Add Total Row
    html += `<tr class="total-row">
      <td colspan="3"><strong>TOTAL FILTERED (${{filteredKabs.length}} KAB/KOTA)</strong></td>
      <td><strong>${{totSites}}</strong></td>
      <td style="color:#34d399"><strong>${{totProfit}}</strong></td>
      <td style="color:#f87171"><strong>${{totLoss}}</strong></td>
      <td><strong>${{totSites ? (totLoss/totSites*100).toFixed(1) : 0}}%</strong></td>
      <td><strong>${{totLt3}}</strong></td>
      <td><strong>${{tot3_5}}</strong></td>
      <td><strong>${{tot5_10}}</strong></td>
      <td style="color:#f87171"><strong>${{totGt10}}</strong></td>
      <td style="text-align:right;color:${{totMargin < 0 ? '#f87171' : '#34d399'}}"><strong>${{Math.round(totMargin).toLocaleString()}}</strong></td>
      <td>-</td>
    </tr>`;

    tbody.innerHTML = html;
  }}

  // Kabupaten Drilldown with Site Class, Status OA, Action SQA, Sales & NOP
  function openKabDrilldown(kabName, branchName) {{
    const box = document.getElementById('kab-drilldown-box');
    document.getElementById('drilldown-kab-name').innerText = kabName;
    document.getElementById('drilldown-subtitle').innerText = `Menampilkan seluruh site berstatus Loss di Kabupaten ${{kabName}} (${{branchName}})`;

    const kabLossSites = augData.filter(r => r[4] === kabName && r[13] === 'Loss');
    const tbody = document.getElementById('drilldown-tbody');

    if (kabLossSites.length === 0) {{
      tbody.innerHTML = `<tr><td colspan="21" style="text-align:center;padding:25px;color:var(--text-muted)">Tidak ada site loss di Kabupaten ini</td></tr>`;
    }} else {{
      let html = '';
      kabLossSites.forEach(s => {{
        const sClass = s[7] || '-';
        let classBadge = `<span class="badge badge-silver">${{sClass}}</span>`;
        if (sClass === 'DIAMOND') classBadge = `<span class="badge badge-diamond">DIAMOND</span>`;
        else if (sClass === 'PLATINUM') classBadge = `<span class="badge badge-platinum">PLATINUM</span>`;
        else if (sClass === 'GOLD') classBadge = `<span class="badge badge-gold">GOLD</span>`;
        else if (sClass === 'BRONZE') classBadge = `<span class="badge badge-bronze">BRONZE</span>`;

        const sOA = s[8] || '-';
        const oaBadge = (sOA === 'OA') 
          ? `<span class="badge badge-success">OA</span>` 
          : `<span class="badge badge-danger">Not OA</span>`;

        html += `<tr>
          <td><strong style="color:#f87171;font-family:'JetBrains Mono'">${{s[0]}}</strong></td>
          <td title="${{s[1]}}">${{s[1]}}</td>
          <td>${{classBadge}}</td>
          <td>${{oaBadge}}</td>
          <td>${{s[5] || '-'}}</td>
          <td>${{s[6] || '-'}}</td>
          <td style="text-align:right">${{Math.round(s[9]).toLocaleString()}}</td>
          <td style="text-align:right">${{Math.round(s[10]).toLocaleString()}}</td>
          <td style="text-align:right;color:#f87171;font-weight:700">${{Math.round(s[11]).toLocaleString()}}</td>
          <td><span class="badge badge-danger">${{s[14]}}</span></td>
          <td>${{s[15]}}</td>
          <td style="color:#fca5a5">${{s[20] || '-'}}</td>
          <td style="color:#fcd34d">${{s[22] || '-'}}</td>
          <td style="color:#6ee7b7">${{s[24] || '-'}}</td>
          <td><span class="badge badge-purple">${{s[32] || '-'}}</span></td>
          <td>${{s[34] || '-'}}</td>
          <td style="color:#93c5fd">${{s[36] || '-'}}</td>
          <td style="color:#38bdf8">${{s[38] || '-'}}</td>
          <td style="text-align:right;color:#38bdf8;font-family:'JetBrains Mono'">${{s[44] ? ('Rp ' + Number(s[44]).toLocaleString()) : '-'}}</td>
          <td style="color:#38bdf8">${{s[39] || '-'}}</td>
          <td>${{s[43] ? ('<span class="badge" style="background:rgba(6,182,212,0.15);color:#22d3ee;border:1px solid rgba(6,182,212,0.3)">' + s[43] + '</span>') : '-'}}</td>
        </tr>`;
      }});
      tbody.innerHTML = html;
    }}

    box.style.display = 'block';
    box.scrollIntoView({{ behavior: 'smooth' }});
  }}

  function closeDrilldown() {{
    document.getElementById('kab-drilldown-box').style.display = 'none';
  }}

  // --- SUBTAB PROGRESS & ACTION TRACKER INSIDE BRANCH FOCUS ---
  let filteredProgRows = augData.filter(r => r[13] === 'Loss');
  let curProgPage = 1;
  let progPageSize = 50;
  let activePriorityPill = 'ALL';

  function getPriorityBadge(val, type) {{
    if (!val || val === 'Not Priority' || val === '-' || val === '#N/A' || val === 'None') {{
      return `<span style="color:var(--text-dim);font-size:11px">Not Priority</span>`;
    }}
    if (val === 'P1') {{
      return `<span class="badge" style="background:rgba(239,68,68,0.25);color:#fca5a5;border:1px solid rgba(239,68,68,0.5);font-weight:700">🔥 P1</span>`;
    }}
    if (val === 'P2') {{
      return `<span class="badge" style="background:rgba(245,158,11,0.25);color:#fcd34d;border:1px solid rgba(245,158,11,0.5);font-weight:700">⚡ P2</span>`;
    }}
    if (val === 'P3') {{
      return `<span class="badge" style="background:rgba(168,85,247,0.25);color:#d8b4fe;border:1px solid rgba(168,85,247,0.5);font-weight:700">📌 P3</span>`;
    }}
    return `<span class="badge badge-info">${{val}}</span>`;
  }}

  function isFilledVal(v) {{
    if (v === null || v === undefined) return false;
    const s = String(v).trim();
    return s !== '' && s !== '-' && s !== '0' && s !== '#N/A' && s !== 'None';
  }}

  function hasSalesPrioAction(r) {{
    return isFilledVal(r[22]) || isFilledVal(r[23]) || isFilledVal(r[24]) || isFilledVal(r[25]);
  }}
  function hasNetPartSQA(r) {{ return isFilledVal(r[20]) || isFilledVal(r[21]); }}
  function hasNetPartRHI(r) {{ return isFilledVal(r[27]) || isFilledVal(r[28]); }}
  function hasNetPartTHI(r) {{ return isFilledVal(r[30]) || isFilledVal(r[31]); }}
  function hasNetPartAvail(r) {{ return isFilledVal(r[35]); }}
  function hasNetPartPL(r) {{ return isFilledVal(r[37]); }}
  function hasNetPartCost(r) {{ return isFilledVal(r[38]) || isFilledVal(r[39]); }}
  function hasNetPrioAction(r) {{
    return hasNetPartSQA(r) || hasNetPartRHI(r) || hasNetPartTHI(r) || hasNetPartAvail(r) || hasNetPartPL(r) || hasNetPartCost(r);
  }}
  function prioSiteHasProg(r) {{
    const isSales = (r[17] === 'P1' || r[17] === 'P2' || r[17] === 'P3');
    const isNet = (r[18] === 'P1' || r[18] === 'P2' || r[18] === 'P3');
    if (isSales && hasSalesPrioAction(r)) return true;
    if (isNet && hasNetPrioAction(r)) return true;
    return false;
  }}

  function setPriorityFilter(prioType) {{
    activePriorityPill = prioType;
    document.querySelectorAll('.priority-pill').forEach(btn => {{
      if (!btn.id.startsWith('prio-part-')) btn.classList.remove('active');
    }});
    const btnId = 
      prioType === 'ALL' ? 'prio-pill-all' :
      prioType === 'ANY_PRIO' ? 'prio-pill-any' :
      prioType === 'SALES_ONLY' ? 'prio-pill-sales' :
      prioType === 'NET_ONLY' ? 'prio-pill-net' :
      prioType === 'P1_ONLY' ? 'prio-pill-p1' :
      prioType === 'P2_ONLY' ? 'prio-pill-p2' :
      prioType === 'ACTION_DONE' ? 'prio-pill-done' :
      prioType === 'ACTION_PENDING' ? 'prio-pill-pending' : '';
    const btn = document.getElementById(btnId);
    if (btn) btn.classList.add('active');

    // If switching away from NET_ONLY, reset net part dropdown and pills
    if (prioType !== 'NET_ONLY') {{
      const el = document.getElementById('prog-filter-net-part');
      if (el) el.value = 'ALL';
      document.querySelectorAll('[id^="prio-part-"]').forEach(b => b.classList.remove('active'));
      const allPartBtn = document.getElementById('prio-part-all');
      if (allPartBtn) allPartBtn.classList.add('active');
    }}
    applyProgFilters();
  }}

  function setNetPartFilter(part) {{
    const el = document.getElementById('prog-filter-net-part');
    if (el) el.value = part;
    
    document.querySelectorAll('[id^="prio-part-"]').forEach(b => b.classList.remove('active'));
    const pillId = 
      part === 'ALL' ? 'prio-part-all' :
      part === 'PART_SQA' ? 'prio-part-sqa' :
      part === 'PART_RHI' ? 'prio-part-rhi' :
      part === 'PART_THI' ? 'prio-part-thi' :
      part === 'PART_AVAIL' ? 'prio-part-avail' :
      part === 'PART_PL' ? 'prio-part-pl' :
      part === 'PART_COST' ? 'prio-part-cost' : '';
    const pillBtn = document.getElementById(pillId);
    if (pillBtn) pillBtn.classList.add('active');

    if (part !== 'ALL') {{
      activePriorityPill = 'NET_ONLY';
      document.querySelectorAll('.priority-pill').forEach(btn => {{
        if (!btn.id.startsWith('prio-part-')) btn.classList.remove('active');
      }});
      const netBtn = document.getElementById('prio-pill-net');
      if (netBtn) netBtn.classList.add('active');
    }}

    applyProgFilters();
  }}

  function onNetPartDropdownChange() {{
    const el = document.getElementById('prog-filter-net-part');
    const val = el ? el.value : 'ALL';
    document.querySelectorAll('[id^="prio-part-"]').forEach(b => b.classList.remove('active'));
    const pillId = 
      val === 'ALL' ? 'prio-part-all' :
      val === 'PART_SQA' ? 'prio-part-sqa' :
      val === 'PART_RHI' ? 'prio-part-rhi' :
      val === 'PART_THI' ? 'prio-part-thi' :
      val === 'PART_AVAIL' ? 'prio-part-avail' :
      val === 'PART_PL' ? 'prio-part-pl' :
      val === 'PART_COST' ? 'prio-part-cost' : '';
    const pillBtn = document.getElementById(pillId);
    if (pillBtn) pillBtn.classList.add('active');

    if (val !== 'ALL') {{
      activePriorityPill = 'NET_ONLY';
      document.querySelectorAll('.priority-pill').forEach(btn => {{
        if (!btn.id.startsWith('prio-part-')) btn.classList.remove('active');
      }});
      const netBtn = document.getElementById('prio-pill-net');
      if (netBtn) netBtn.classList.add('active');
    }}
    applyProgFilters();
  }}

  // --- GLOBAL SORTING STATE & ENGINE (SHORT BESAR KE KECIL) ---
  let progSortCol = null;
  let progSortDir = 'DESC'; // 'DESC' (Besar ke Kecil) or 'ASC' (Kecil ke Besar)

  let augSortCol = null;
  let augSortDir = 'DESC';

  function sortRows(rows, col, dir) {{
    if (!col || col === 'NONE') return rows;
    const sorted = [...rows];
    sorted.sort((a, b) => {{
      let va, vb;
      if (col === 'defisit') {{
        // Defisit Terbesar (Besar ke Kecil): Magnitude kerugian Math.abs(margin)
        va = Math.abs(Number(a[11]) || 0);
        vb = Math.abs(Number(b[11]) || 0);
        return dir === 'DESC' ? (vb - va) : (va - vb);
      }} else if (col === 'margin') {{
        va = Number(a[11]) || 0;
        vb = Number(b[11]) || 0;
        return dir === 'DESC' ? (vb - va) : (va - vb);
      }} else if (col === 'cost') {{
        va = Number(a[10]) || 0;
        vb = Number(b[10]) || 0;
        return dir === 'DESC' ? (vb - va) : (va - vb);
      }} else if (col === 'rev') {{
        va = Number(a[9]) || 0;
        vb = Number(b[9]) || 0;
        return dir === 'DESC' ? (vb - va) : (va - vb);
      }} else if (col === 'nilai_cost') {{
        va = Number(a[44]) || 0;
        vb = Number(b[44]) || 0;
        return dir === 'DESC' ? (vb - va) : (va - vb);
      }} else if (col === 'site_id') {{
        va = String(a[0] || '').toLowerCase();
        vb = String(b[0] || '').toLowerCase();
        return dir === 'DESC' ? vb.localeCompare(va) : va.localeCompare(vb);
      }} else if (col === 'site_name') {{
        va = String(a[1] || '').toLowerCase();
        vb = String(b[1] || '').toLowerCase();
        return dir === 'DESC' ? vb.localeCompare(va) : va.localeCompare(vb);
      }} else if (col === 'branch') {{
        va = String(a[3] || '').toLowerCase();
        vb = String(b[3] || '').toLowerCase();
        return dir === 'DESC' ? vb.localeCompare(va) : va.localeCompare(vb);
      }} else if (col === 'kabupaten') {{
        va = String(a[4] || '').toLowerCase();
        vb = String(b[4] || '').toLowerCase();
        return dir === 'DESC' ? vb.localeCompare(va) : va.localeCompare(vb);
      }} else if (col === 'pnl') {{
        va = String(a[13] || '').toLowerCase();
        vb = String(b[13] || '').toLowerCase();
        return dir === 'DESC' ? vb.localeCompare(va) : va.localeCompare(vb);
      }} else if (col === 'ag') {{
        va = String(a[42] || '').toLowerCase();
        vb = String(b[42] || '').toLowerCase();
        return dir === 'DESC' ? vb.localeCompare(va) : va.localeCompare(vb);
      }}
      return 0;
    }});
    return sorted;
  }}

  function updateSortIndicators(prefix, col, dir) {{
    const arrow = dir === 'DESC' ? '▼' : '▲';
    const cols = ['defisit', 'cost', 'rev', 'nilai_cost', 'site_id', 'site_name', 'branch', 'kabupaten', 'pnl', 'ag'];
    cols.forEach(c => {{
      const el = document.getElementById(`th-sort-${{prefix}}-${{c}}`);
      if (el) {{
        if (c === col) {{
          el.innerText = arrow;
          el.style.color = '#38bdf8';
          el.style.fontWeight = 'bold';
        }} else {{
          el.innerText = '⬍';
          el.style.color = 'var(--text-dim)';
          el.style.fontWeight = 'normal';
        }}
      }}
    }});
  }}

  function onProgSortChange(val) {{
    if (val === 'NONE') {{
      progSortCol = null;
      progSortDir = 'DESC';
    }} else if (val === 'DEFISIT_DESC') {{
      progSortCol = 'defisit';
      progSortDir = 'DESC';
    }} else if (val === 'DEFISIT_ASC') {{
      progSortCol = 'defisit';
      progSortDir = 'ASC';
    }} else if (val === 'COST_DESC') {{
      progSortCol = 'cost';
      progSortDir = 'DESC';
    }} else if (val === 'REV_DESC') {{
      progSortCol = 'rev';
      progSortDir = 'DESC';
    }} else if (val === 'NILAI_COST_DESC') {{
      progSortCol = 'nilai_cost';
      progSortDir = 'DESC';
    }} else if (val === 'SITE_ID_ASC') {{
      progSortCol = 'site_id';
      progSortDir = 'ASC';
    }}
    applyProgFilters();
  }}

  function sortProgTable(colKey) {{
    if (progSortCol === colKey) {{
      progSortDir = (progSortDir === 'DESC') ? 'ASC' : 'DESC';
    }} else {{
      progSortCol = colKey;
      progSortDir = 'DESC'; // Default to Besar ke Kecil!
    }}
    const sel = document.getElementById('prog-sort-select');
    if (sel) {{
      if (colKey === 'defisit') sel.value = progSortDir === 'DESC' ? 'DEFISIT_DESC' : 'DEFISIT_ASC';
      else if (colKey === 'cost') sel.value = 'COST_DESC';
      else if (colKey === 'rev') sel.value = 'REV_DESC';
      else if (colKey === 'nilai_cost') sel.value = 'NILAI_COST_DESC';
      else if (colKey === 'site_id') sel.value = 'SITE_ID_ASC';
      else sel.value = 'NONE';
    }}
    applyProgFilters();
  }}

  function onProgSearch() {{
    applyProgFilters();
  }}

  function applyProgFilters() {{
    const q = (document.getElementById('prog-search-input').value || '').toLowerCase().trim();
    const branchVal = document.getElementById('prog-filter-nop').value;
    const pnlVal = document.getElementById('prog-filter-pnl').value;
    const plweekVal = document.getElementById('prog-filter-plweek').value;
    const transVal = document.getElementById('prog-filter-transport').value;
    const oaVal = document.getElementById('prog-filter-oa') ? document.getElementById('prog-filter-oa').value : 'ALL';
    const agVal = document.getElementById('prog-filter-ag') ? document.getElementById('prog-filter-ag').value : 'ALL';
    const salesPrioVal = document.getElementById('prog-filter-priority-sales') ? document.getElementById('prog-filter-priority-sales').value : 'ALL';
    const netPrioVal = document.getElementById('prog-filter-priority-net') ? document.getElementById('prog-filter-priority-net').value : 'ALL';
    const netPartVal = document.getElementById('prog-filter-net-part') ? document.getElementById('prog-filter-net-part').value : 'ALL';
    const cekVal = document.getElementById('prog-filter-cek') ? document.getElementById('prog-filter-cek').value : 'ALL';

    filteredProgRows = augData.filter(r => {{
      // 1. Text Search across identifiers, priorities and action columns
      if (q) {{
        const text = (r[0] + ' ' + r[1] + ' ' + r[2] + ' ' + r[3] + ' ' + r[4] + ' ' + r[17] + ' ' + r[18] + ' ' + r[20] + ' ' + r[21] + ' ' + r[22] + ' ' + r[23] + ' ' + r[24] + ' ' + r[25] + ' ' + r[26] + ' ' + r[27] + ' ' + r[28] + ' ' + r[29] + ' ' + r[30] + ' ' + r[31] + ' ' + r[32] + ' ' + r[33] + ' ' + r[34] + ' ' + r[35] + ' ' + r[36] + ' ' + r[37] + ' ' + r[38] + ' ' + r[39] + ' ' + (r[43] || '')).toLowerCase();
        if (!text.includes(q)) return false;
      }}

      // 2. Branch & PnL
      if (branchVal !== 'ALL' && r[3] !== branchVal) return false;
      if (pnlVal !== 'ALL' && r[13] !== pnlVal) return false;
      if (cekVal !== 'ALL' && r[43] !== cekVal) return false;

      // 3. Priority Checks
      const isSalesPrio = (r[17] === 'P1' || r[17] === 'P2' || r[17] === 'P3');
      const isNetPrio = (r[18] === 'P1' || r[18] === 'P2' || r[18] === 'P3');
      const isAnyPrio = isSalesPrio || isNetPrio;

      const hasSalesAct = hasSalesPrioAction(r);
      const hasNetAct = hasNetPrioAction(r);
      const siteHasProg = prioSiteHasProg(r);

      // Pill filter evaluation
      if (activePriorityPill === 'ANY_PRIO' && !isAnyPrio) return false;
      if (activePriorityPill === 'SALES_ONLY' && !isSalesPrio) return false;
      if (activePriorityPill === 'NET_ONLY' && !isNetPrio) return false;
      if (activePriorityPill === 'P1_ONLY' && r[17] !== 'P1' && r[18] !== 'P1') return false;
      if (activePriorityPill === 'P2_ONLY' && r[17] !== 'P2' && r[18] !== 'P2') return false;
      if (activePriorityPill === 'ACTION_DONE') {{
        if (!isAnyPrio || !siteHasProg) return false;
      }}
      if (activePriorityPill === 'ACTION_PENDING') {{
        if (!isAnyPrio || siteHasProg) return false;
      }}

      // Dropdown Priority Filters
      if (salesPrioVal !== 'ALL') {{
        if (salesPrioVal === 'PRIORITY_ONLY' && !isSalesPrio) return false;
        if (salesPrioVal !== 'PRIORITY_ONLY' && r[17] !== salesPrioVal) return false;
      }}
      if (netPrioVal !== 'ALL') {{
        if (netPrioVal === 'PRIORITY_ONLY' && !isNetPrio) return false;
        if (netPrioVal !== 'PRIORITY_ONLY' && r[18] !== netPrioVal) return false;
      }}

      // Dropdown Priority Network Part Filter
      if (netPartVal !== 'ALL') {{
        if (!isNetPrio) return false;
        if (netPartVal === 'PART_SQA' && !hasNetPartSQA(r)) return false;
        if (netPartVal === 'PART_RHI' && !hasNetPartRHI(r)) return false;
        if (netPartVal === 'PART_THI' && !hasNetPartTHI(r)) return false;
        if (netPartVal === 'PART_AVAIL' && !hasNetPartAvail(r)) return false;
        if (netPartVal === 'PART_PL' && !hasNetPartPL(r)) return false;
        if (netPartVal === 'PART_COST' && !hasNetPartCost(r)) return false;
        if (netPartVal === 'PART_ANY' && !hasNetAct) return false;
        if (netPartVal === 'PART_NONE' && hasNetAct) return false;
      }}

      // 4. Other Filters
      if (plweekVal !== 'ALL') {{
        if (plweekVal === 'CLEAR' && r[36] !== 'CLEAR') return false;
        if (plweekVal === 'UNCLEAR' && r[36] === 'CLEAR') return false;
      }}
      if (oaVal !== 'ALL') {{
        if (oaVal === 'OA' && r[8] !== 'OA') return false;
        if (oaVal === 'Not OA' && r[8] === 'OA') return false;
      }}
      if (transVal !== 'ALL' && r[32] !== transVal) return false;
      if (agVal !== 'ALL') {{
        if (agVal === 'Program H2' && r[42] !== 'Program H2') return false;
        if (agVal === '0' && r[42] !== '0') return false;
        if (agVal === '#N/A' && r[42] !== '#N/A') return false;
      }}
      return true;
    }});

    // Apply Sorting (Besar ke Kecil / Shorting)
    if (progSortCol && progSortCol !== 'NONE') {{
      filteredProgRows = sortRows(filteredProgRows, progSortCol, progSortDir);
    }}

    curProgPage = 1;
    renderProgTable();
  }}

  function changeProgPageSize(val) {{
    progPageSize = parseInt(val) || 50;
    curProgPage = 1;
    renderProgTable();
  }}

  function goToProgPage(p) {{
    curProgPage = p;
    renderProgTable();
  }}

  function prevProgPage() {{
    if (curProgPage > 1) {{
      curProgPage--;
      renderProgTable();
    }}
  }}

  function nextProgPage() {{
    const maxPage = Math.ceil(filteredProgRows.length / progPageSize) || 1;
    if (curProgPage < maxPage) {{
      curProgPage++;
      renderProgTable();
    }}
  }}

  function goToLastProgPage() {{
    const maxPage = Math.ceil(filteredProgRows.length / progPageSize) || 1;
    curProgPage = maxPage;
    renderProgTable();
  }}

  function renderProgTable() {{
    const tbody = document.getElementById('prog-table-body');
    const total = filteredProgRows.length;
    const totalPages = Math.ceil(total / progPageSize) || 1;

    if (curProgPage > totalPages) curProgPage = totalPages;
    if (curProgPage < 1) curProgPage = 1;

    const startIdx = (curProgPage - 1) * progPageSize;
    const endIdx = Math.min(startIdx + progPageSize, total);

    document.getElementById('prog-current-page-num').innerText = curProgPage;
    document.getElementById('prog-total-pages-num').innerText = totalPages;
    document.getElementById('prog-btn-prev').disabled = (curProgPage === 1);
    document.getElementById('prog-btn-first').disabled = (curProgPage === 1);
    document.getElementById('prog-btn-next').disabled = (curProgPage === totalPages);
    document.getElementById('prog-btn-last').disabled = (curProgPage === totalPages);

    document.getElementById('prog-pagination-info').innerText = total === 0 
      ? 'Tidak ada data site yang sesuai' 
      : `Menampilkan ${{startIdx + 1}} - ${{endIdx}} dari ${{total.toLocaleString()}} site`;

    if (total === 0) {{
      tbody.innerHTML = `<tr><td colspan="32" style="text-align:center;padding:40px;color:var(--text-muted)">Tidak ada data site yang sesuai filter</td></tr>`;
      renderProgCharts();
      renderPriorityCharts();
      return;
    }}

    let rowsHtml = '';
    for (let i = startIdx; i < endIdx; i++) {{
      const r = filteredProgRows[i];
      const isLoss = (r[13] === 'Loss');
      const pnlBadge = isLoss 
        ? `<span class="badge badge-danger">Loss</span>` 
        : `<span class="badge badge-success">Profit</span>`;

      const sClass = r[7] || '-';
      let classBadge = `<span class="badge badge-silver">${{sClass}}</span>`;
      if (sClass === 'DIAMOND') classBadge = `<span class="badge badge-diamond">DIAMOND</span>`;
      else if (sClass === 'PLATINUM') classBadge = `<span class="badge badge-platinum">PLATINUM</span>`;
      else if (sClass === 'GOLD') classBadge = `<span class="badge badge-gold">GOLD</span>`;
      else if (sClass === 'BRONZE') classBadge = `<span class="badge badge-bronze">BRONZE</span>`;

      const sOA = r[8] || '-';
      const oaBadge = (sOA === 'OA') 
        ? `<span class="badge badge-success">OA</span>` 
        : `<span class="badge badge-danger">Not OA</span>`;

      const plWeekBadge = (r[36] === 'CLEAR')
        ? `<span class="badge badge-success">CLEAR</span>`
        : (r[36] ? `<span class="badge badge-warning">${{r[36]}}</span>` : '-');

      const rhiWeekBadge = (r[26] === 'Pass')
        ? `<span class="badge badge-success">Pass</span>`
        : (r[26] === 'Fail' ? `<span class="badge badge-danger">Fail</span>` : (r[26] || '-'));

      const thiWeekBadge = (r[29] === 'Pass')
        ? `<span class="badge badge-success">Pass</span>`
        : (r[29] === 'Fail' ? `<span class="badge badge-danger">Fail</span>` : (r[29] || '-'));

      const availBadge = (r[34] === 'Achieved')
        ? `<span class="badge badge-success">Achieved</span>`
        : (r[34] || '-');

      rowsHtml += `<tr>
        <td style="color:var(--text-dim);font-family:'JetBrains Mono'">${{i + 1}}</td>
        <td><strong style="color:${{isLoss ? '#f87171' : '#38bdf8'}};font-family:'JetBrains Mono'">${{r[0]}}</strong></td>
        <td title="${{r[1]}}" style="max-width:210px;overflow:hidden;text-overflow:ellipsis">${{r[1]}}</td>
        <td><span class="badge" style="background:rgba(255,255,255,0.06);color:#94a3b8">${{r[3]}}</span></td>
        <td>${{r[4]}}</td>
        <td>${{pnlBadge}}</td>
        <td style="text-align:right;color:${{r[11] < 0 ? '#f87171' : '#34d399'}};font-weight:700;font-family:'JetBrains Mono'">${{Math.round(r[11]).toLocaleString()}}</td>
        <td style="text-align:right;font-family:'JetBrains Mono'">${{Math.round(r[10]).toLocaleString()}}</td>
        <td>${{classBadge}}</td>
        <td>${{oaBadge}}</td>
        <td class="new-col-td">${{r[42] === 'Program H2' ? '<span class="badge" style="background:rgba(245,158,11,0.2);color:#fbbf24;border:1px solid rgba(245,158,11,0.4)">Program H2</span>' : (r[42] || '-')}}</td>
        <!-- Priority Sales & Priority Network BADGES -->
        <td>${{getPriorityBadge(r[17], 'sales')}}</td>
        <td>${{getPriorityBadge(r[18], 'net')}}</td>
        <!-- SQA (AH, AI) -->
        <td class="new-col-td" style="color:#fca5a5">${{r[20] || '-'}}</td>
        <td class="new-col-td" style="color:#fca5a5">${{r[21] || '-'}}</td>
        <!-- Sales (AJ, AK) -->
        <td class="new-col-td" style="color:#fcd34d">${{r[22] || '-'}}</td>
        <td class="new-col-td" style="color:#fcd34d">${{r[23] || '-'}}</td>
        <!-- NOP (AL, AM) -->
        <td class="new-col-td" style="color:#6ee7b7">${{r[24] || '-'}}</td>
        <td class="new-col-td" style="color:#6ee7b7">${{r[25] || '-'}}</td>
        <!-- RHI (AN, AO, AP) -->
        <td class="new-col-td" style="font-family:'JetBrains Mono'">${{rhiWeekBadge}}</td>
        <td class="new-col-td" style="color:#c4b5fd">${{r[27] || '-'}}</td>
        <td class="new-col-td" style="color:#c4b5fd">${{r[28] || '-'}}</td>
        <!-- THI (AQ, AR, AS) -->
        <td class="new-col-td" style="font-family:'JetBrains Mono'">${{thiWeekBadge}}</td>
        <td class="new-col-td" style="color:#7dd3fc">${{r[30] || '-'}}</td>
        <td class="new-col-td" style="color:#7dd3fc">${{r[31] || '-'}}</td>
        <!-- Transport & BDBH (AT, AU) -->
        <td class="new-col-td">${{r[32] ? '<span class="badge badge-purple">' + r[32] + '</span>' : '-'}}</td>
        <td class="new-col-td" style="font-family:'JetBrains Mono'">${{r[33] || '-'}}</td>
        <!-- Availability (AV, AW) -->
        <td class="new-col-td">${{availBadge}}</td>
        <td class="new-col-td" style="color:#fbbf24">${{r[35] || '-'}}</td>
        <!-- Status PL & MBA (AX, AY) -->
        <td class="new-col-td">${{plWeekBadge}}</td>
        <td class="new-col-td" style="color:#93c5fd">${{r[37] || '-'}}</td>
        <!-- NOS (AZ, BA, BB, BC) -->
        <td class="new-col-td" style="color:#38bdf8">${{r[38] || '-'}}</td>
        <td class="new-col-td" style="text-align:right;color:#38bdf8;font-family:'JetBrains Mono'">${{r[44] ? ('Rp ' + Number(r[44]).toLocaleString()) : '-'}}</td>
        <td class="new-col-td" style="color:#38bdf8">${{r[39] || '-'}}</td>
        <!-- Cek / Kategori Site -->
        <td class="new-col-td">${{r[43] ? ('<span class="badge" style="background:rgba(6,182,212,0.15);color:#22d3ee;border:1px solid rgba(6,182,212,0.3)">' + r[43] + '</span>') : '-'}}</td>
      </tr>`;
    }}

    tbody.innerHTML = rowsHtml;
    updateSortIndicators('prog', progSortCol, progSortDir);
    renderProgCharts();
    renderPriorityCharts();
  }}

  function renderPriorityCharts() {{
    const containerVol = document.getElementById('priority-volume-bars');
    const containerProg = document.getElementById('priority-progress-bars');
    if (!containerVol || !containerProg) return;

    const targetRows = (activeBranch === 'ALL')
      ? augData
      : augData.filter(r => r[3] === activeBranch);

    const prioSalesRows = targetRows.filter(r => r[17] === 'P1' || r[17] === 'P2' || r[17] === 'P3');
    const prioNetRows = targetRows.filter(r => r[18] === 'P1' || r[18] === 'P2' || r[18] === 'P3');
    const allPrioRows = targetRows.filter(r => (r[17] === 'P1' || r[17] === 'P2' || r[17] === 'P3') || (r[18] === 'P1' || r[18] === 'P2' || r[18] === 'P3'));

    const sP1 = prioSalesRows.filter(r => r[17] === 'P1').length;
    const sP2 = prioSalesRows.filter(r => r[17] === 'P2').length;
    const sP3 = prioSalesRows.filter(r => r[17] === 'P3').length;
    const sTotal = prioSalesRows.length;

    const nP1 = prioNetRows.filter(r => r[18] === 'P1').length;
    const nP2 = prioNetRows.filter(r => r[18] === 'P2').length;
    const nP3 = prioNetRows.filter(r => r[18] === 'P3').length;
    const nTotal = prioNetRows.length;

    const allTotal = allPrioRows.length;

    // Priority Sales Action: Sales & NOP
    const sProg = prioSalesRows.filter(hasSalesPrioAction).length;
    const sProgPct = sTotal > 0 ? (sProg / sTotal * 100).toFixed(1) : '0.0';

    // Priority Network Action: SQA, RHI, THI, Avail, PL/MBA, Cost
    const nProg = prioNetRows.filter(hasNetPrioAction).length;
    const nProgPct = nTotal > 0 ? (nProg / nTotal * 100).toFixed(1) : '0.0';

    // Overall Priority Action Progress
    const allProg = allPrioRows.filter(prioSiteHasProg).length;
    const allProgPct = allTotal > 0 ? (allProg / allTotal * 100).toFixed(1) : '0.0';
    const allPending = allTotal - allProg;
    const allPendingPct = allTotal > 0 ? (allPending / allTotal * 100).toFixed(1) : '0.0';

    // Priority Network Parts Breakdown:
    const netSqa = prioNetRows.filter(hasNetPartSQA).length;
    const netRhi = prioNetRows.filter(hasNetPartRHI).length;
    const netThi = prioNetRows.filter(hasNetPartTHI).length;
    const netAvail = prioNetRows.filter(hasNetPartAvail).length;
    const netPl = prioNetRows.filter(hasNetPartPL).length;
    const netCost = prioNetRows.filter(hasNetPartCost).length;

    // Priority Sales Breakdown:
    const sSales = prioSalesRows.filter(r => isFilledVal(r[22]) || isFilledVal(r[23])).length;
    const sNop = prioSalesRows.filter(r => isFilledVal(r[24]) || isFilledVal(r[25])).length;

    // KPI Cards Updates
    const badge = document.getElementById('priority-main-badge');
    if (badge) badge.innerText = `${{allTotal}} Priority Sites (${{activeBranch === 'ALL' ? 'Papua Maluku' : activeBranch}})`;

    const elSBadge = document.getElementById('kpi-prio-sales-badge');
    if (elSBadge) elSBadge.innerText = `${{sTotal}} Sites`;
    const elSTiers = document.getElementById('kpi-prio-sales-tiers');
    if (elSTiers) elSTiers.innerText = `P1: ${{sP1}} | P2: ${{sP2}} | P3: ${{sP3}}`;
    const elSVal = document.getElementById('kpi-prio-sales-val');
    if (elSVal) elSVal.innerHTML = `${{sTotal}} <span style="font-size:12px;color:var(--text-dim);font-weight:normal">sites</span>`;
    const elSProg = document.getElementById('kpi-prio-sales-prog');
    if (elSProg) elSProg.innerText = `Progress: ${{sProg}} (${{sProgPct}}%)`;

    const elNBadge = document.getElementById('kpi-prio-net-badge');
    if (elNBadge) elNBadge.innerText = `${{nTotal}} Sites`;
    const elNTiers = document.getElementById('kpi-prio-net-tiers');
    if (elNTiers) elNTiers.innerText = `P1: ${{nP1}} | P2: ${{nP2}} | P3: ${{nP3}}`;
    const elNVal = document.getElementById('kpi-prio-net-val');
    if (elNVal) elNVal.innerHTML = `${{nTotal}} <span style="font-size:12px;color:var(--text-dim);font-weight:normal">sites</span>`;
    const elNProg = document.getElementById('kpi-prio-net-prog');
    if (elNProg) elNProg.innerText = `Progress: ${{nProg}} (${{nProgPct}}%)`;

    const lossRowsCount = targetRows.filter(r => r[13] === 'Loss').length;
    const prioLossPct = lossRowsCount > 0 ? (allTotal / lossRowsCount * 100).toFixed(1) : '0.0';
    const elTBadge = document.getElementById('kpi-prio-total-badge');
    if (elTBadge) elTBadge.innerText = `${{allTotal}} Sites`;
    const elTRatio = document.getElementById('kpi-prio-total-ratio');
    if (elTRatio) elTRatio.innerText = `${{prioLossPct}}% dari ${{lossRowsCount}} Loss`;
    const elTVal = document.getElementById('kpi-prio-total-val');
    if (elTVal) elTVal.innerHTML = `${{allTotal}} <span style="font-size:12px;color:var(--text-dim);font-weight:normal">sites</span>`;
    const elTProg = document.getElementById('kpi-prio-total-prog');
    if (elTProg) elTProg.innerText = `Progress: ${{allProg}} (${{allProgPct}}%)`;

    const elDoneVal = document.getElementById('kpi-prio-done-val');
    if (elDoneVal) elDoneVal.innerHTML = `${{allProg}} <span style="font-size:12px;color:var(--text-dim);font-weight:normal">sites</span>`;
    const elDonePct = document.getElementById('kpi-prio-done-pct');
    if (elDonePct) elDonePct.innerText = `${{allProgPct}}%`;
    const elDoneSub = document.getElementById('kpi-prio-done-sub');
    if (elDoneSub) elDoneSub.innerText = `NOP (${{sNop}}), RHI (${{netRhi}}), MBA (${{netPl}}), Sales (${{sSales}})`;

    const elPendVal = document.getElementById('kpi-prio-pending-val');
    if (elPendVal) elPendVal.innerHTML = `${{allPending}} <span style="font-size:12px;color:var(--text-dim);font-weight:normal">sites</span>`;
    const elPendPct = document.getElementById('kpi-prio-pending-pct');
    if (elPendPct) elPendPct.innerText = `${{allPendingPct}}%`;

    // Update Quick Priority Filter Pills & Part Pills dynamically
    const allP1 = allPrioRows.filter(r => r[17] === 'P1' || r[18] === 'P1').length;
    const allP2 = allPrioRows.filter(r => r[17] === 'P2' || r[18] === 'P2').length;
    const pAllBtn = document.getElementById('prio-pill-all');
    if (pAllBtn) pAllBtn.innerText = `Semua Site (${{lossRowsCount}} Loss)`;
    const pAnyBtn = document.getElementById('prio-pill-any');
    if (pAnyBtn) pAnyBtn.innerText = `⭐ Semua Priority Sites (${{allTotal}})`;
    const pSalesBtn = document.getElementById('prio-pill-sales');
    if (pSalesBtn) pSalesBtn.innerText = `💼 Priority Sales (${{sTotal}})`;
    const pNetBtn = document.getElementById('prio-pill-net');
    if (pNetBtn) pNetBtn.innerText = `📡 Priority Network (${{nTotal}})`;
    const pP1Btn = document.getElementById('prio-pill-p1');
    if (pP1Btn) pP1Btn.innerText = `🔥 Priority P1 (${{allP1}})`;
    const pP2Btn = document.getElementById('prio-pill-p2');
    if (pP2Btn) pP2Btn.innerText = `⚡ Priority P2 (${{allP2}})`;
    const pDoneBtn = document.getElementById('prio-pill-done');
    if (pDoneBtn) pDoneBtn.innerText = `✅ Sudah Ada Action (${{allProg}})`;
    const pPendBtn = document.getElementById('prio-pill-pending');
    if (pPendBtn) pPendBtn.innerText = `⏳ Belum Ada Action (${{allPending}})`;

    const ptSqa = document.getElementById('prio-part-sqa');
    if (ptSqa) ptSqa.innerText = `🎯 Part SQA (${{netSqa}})`;
    const ptRhi = document.getElementById('prio-part-rhi');
    if (ptRhi) ptRhi.innerText = `📡 Part RHI (${{netRhi}})`;
    const ptThi = document.getElementById('prio-part-thi');
    if (ptThi) ptThi.innerText = `📶 Part THI (${{netThi}})`;
    const ptAvail = document.getElementById('prio-part-avail');
    if (ptAvail) ptAvail.innerText = `⚡ Part Avail (${{netAvail}})`;
    const ptPl = document.getElementById('prio-part-pl');
    if (ptPl) ptPl.innerText = `🔌 Part PL/MBA (${{netPl}})`;
    const ptCost = document.getElementById('prio-part-cost');
    if (ptCost) ptCost.innerText = `💰 Part Cost (${{netCost}})`;

    // 1. RENDER DIAGRAM BATANG 1: VOLUME & DISTRIBUSI
    let volHtml = `
      <div class="prio-bar-wrapper" onclick="setPriorityFilter('SALES_ONLY')" style="cursor:pointer" title="Klik untuk filter ke Priority Sales">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px">
          <div style="display:flex;align-items:center;gap:8px">
            <span style="width:10px;height:10px;border-radius:3px;background:#f59e0b;display:inline-block"></span>
            <span style="font-size:13px;font-weight:700;color:#fff">Priority Sales</span>
            <span class="badge" style="background:rgba(245,158,11,0.2);color:#fcd34d;font-size:11px">${{sTotal}} Sites</span>
          </div>
          <div style="font-size:12px;color:var(--text-muted)">
            <span style="color:#ef4444;font-weight:700">P1: ${{sP1}}</span> | 
            <span style="color:#f59e0b;font-weight:700">P2: ${{sP2}}</span> | 
            <span style="color:#a855f7;font-weight:700">P3: ${{sP3}}</span>
          </div>
        </div>
        <div style="height:26px;background:rgba(255,255,255,0.06);border-radius:8px;overflow:hidden;display:flex;position:relative">
          <div style="width:${{sTotal > 0 ? (sP1 / sTotal * 100) : 0}}%;background:linear-gradient(90deg, #ef4444, #dc2626);display:flex;align-items:center;justify-content:center;color:#fff;font-size:11px;font-weight:700" title="P1: ${{sP1}} sites">${{sP1 > 2 ? 'P1: ' + sP1 : ''}}</div>
          <div style="width:${{sTotal > 0 ? (sP2 / sTotal * 100) : 0}}%;background:linear-gradient(90deg, #f59e0b, #d97706);display:flex;align-items:center;justify-content:center;color:#fff;font-size:11px;font-weight:700" title="P2: ${{sP2}} sites">${{sP2 > 2 ? 'P2: ' + sP2 : ''}}</div>
          <div style="width:${{sTotal > 0 ? (sP3 / sTotal * 100) : 0}}%;background:linear-gradient(90deg, #8b5cf6, #7c3aed);display:flex;align-items:center;justify-content:center;color:#fff;font-size:11px;font-weight:700" title="P3: ${{sP3}} sites">${{sP3 > 0 ? sP3 : ''}}</div>
        </div>
      </div>

      <div class="prio-bar-wrapper" onclick="setPriorityFilter('NET_ONLY')" style="cursor:pointer;margin-top:10px" title="Klik untuk filter ke Priority Network">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px">
          <div style="display:flex;align-items:center;gap:8px">
            <span style="width:10px;height:10px;border-radius:3px;background:#38bdf8;display:inline-block"></span>
            <span style="font-size:13px;font-weight:700;color:#fff">Priority Network</span>
            <span class="badge" style="background:rgba(56,189,248,0.2);color:#7dd3fc;font-size:11px">${{nTotal}} Sites</span>
          </div>
          <div style="font-size:12px;color:var(--text-muted)">
            <span style="color:#ef4444;font-weight:700">P1: ${{nP1}}</span> | 
            <span style="color:#f59e0b;font-weight:700">P2: ${{nP2}}</span> | 
            <span style="color:#a855f7;font-weight:700">P3: ${{nP3}}</span>
          </div>
        </div>
        <div style="height:26px;background:rgba(255,255,255,0.06);border-radius:8px;overflow:hidden;display:flex;position:relative">
          <div style="width:${{nTotal > 0 ? (nP1 / nTotal * 100) : 0}}%;background:linear-gradient(90deg, #ef4444, #dc2626);display:flex;align-items:center;justify-content:center;color:#fff;font-size:11px;font-weight:700" title="P1: ${{nP1}} sites">${{nP1 > 1 ? 'P1: ' + nP1 : ''}}</div>
          <div style="width:${{nTotal > 0 ? (nP2 / nTotal * 100) : 0}}%;background:linear-gradient(90deg, #38bdf8, #0284c7);display:flex;align-items:center;justify-content:center;color:#fff;font-size:11px;font-weight:700" title="P2: ${{nP2}} sites">${{nP2 > 2 ? 'P2: ' + nP2 : ''}}</div>
          <div style="width:${{nTotal > 0 ? (nP3 / nTotal * 100) : 0}}%;background:linear-gradient(90deg, #8b5cf6, #7c3aed);display:flex;align-items:center;justify-content:center;color:#fff;font-size:11px;font-weight:700" title="P3: ${{nP3}} sites">${{nP3 > 0 ? nP3 : ''}}</div>
        </div>
      </div>

      <div style="margin-top:14px;padding-top:12px;border-top:1px dashed rgba(255,255,255,0.1)">
        <div style="font-size:11px;color:var(--text-muted);font-weight:700;margin-bottom:8px;text-transform:uppercase;letter-spacing:0.5px">Distribusi Site Priority per Branch:</div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px">
    `;

    const brList = [
      {{ name: 'AMBON', label: 'Ambon', color: '#f59e0b' }},
      {{ name: 'JAYAPURA', label: 'Jayapura', color: '#38bdf8' }},
      {{ name: 'SORONG', label: 'Sorong', color: '#10b981' }},
      {{ name: 'TIMIKA', label: 'Timika', color: '#a855f7' }}
    ];

    brList.forEach(b => {{
      const bSites = allPrioRows.filter(r => r[3] === b.name);
      const bSales = bSites.filter(r => r[17] === 'P1' || r[17] === 'P2' || r[17] === 'P3').length;
      const bNet = bSites.filter(r => r[18] === 'P1' || r[18] === 'P2' || r[18] === 'P3').length;
      volHtml += `
        <div style="background:rgba(255,255,255,0.03);border:1px solid rgba(255,255,255,0.06);border-radius:8px;padding:8px 10px;cursor:pointer" onclick="setProgBranchFilter('${{b.name}}');setPriorityFilter('ANY_PRIO')">
          <div style="display:flex;justify-content:space-between;align-items:center;font-size:12px">
            <span style="font-weight:700;color:${{b.color}}">${{b.label}}</span>
            <span class="badge" style="background:rgba(255,255,255,0.08);color:#fff">${{bSites.length}} Sites</span>
          </div>
          <div style="font-size:11px;color:var(--text-dim);margin-top:2px">Sales: ${{bSales}} | Net: ${{bNet}}</div>
        </div>
      `;
    }});
    volHtml += `</div></div>`;
    containerVol.innerHTML = volHtml;

    // 2. RENDER DIAGRAM BATANG 2: STATUS PROGRESS & ACTION TRACKER
    let progHtml = `
      <div class="prio-bar-wrapper" onclick="setPriorityFilter('NET_ONLY')" style="cursor:pointer" title="Klik untuk filter ke Priority Network">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px">
          <span style="font-size:12px;font-weight:700;color:#7dd3fc">📡 Progress Priority Network (SQA, RHI, THI, Avail, PL, Cost)</span>
          <span style="font-size:12px;color:#34d399;font-weight:700">${{nProg}} / ${{nTotal}} Site Terisi (${{nProgPct}}%)</span>
        </div>
        <div style="height:18px;background:rgba(255,255,255,0.06);border-radius:6px;overflow:hidden;position:relative">
          <div style="width:${{nProgPct}}%;height:100%;background:linear-gradient(90deg, #10b981, #059669);border-radius:6px;transition:width 0.4s ease"></div>
        </div>
      </div>

      <div class="prio-bar-wrapper" onclick="setPriorityFilter('SALES_ONLY')" style="cursor:pointer;margin-top:10px" title="Klik untuk filter ke Priority Sales">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px">
          <span style="font-size:12px;font-weight:700;color:#fcd34d">💼 Progress Priority Sales (Action Sales & NOP)</span>
          <span style="font-size:12px;color:${{sProgPct > 50 ? '#34d399' : '#f59e0b'}};font-weight:700">${{sProg}} / ${{sTotal}} Site Terisi (${{sProgPct}}%)</span>
        </div>
        <div style="height:18px;background:rgba(255,255,255,0.06);border-radius:6px;overflow:hidden;position:relative">
          <div style="width:${{sProgPct}}%;height:100%;background:linear-gradient(90deg, #f59e0b, #d97706);border-radius:6px;transition:width 0.4s ease"></div>
        </div>
      </div>

      <div class="prio-bar-wrapper" onclick="setPriorityFilter('ANY_PRIO')" style="cursor:pointer;margin-top:10px" title="Klik untuk filter ke Semua Priority">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px">
          <span style="font-size:12px;font-weight:700;color:#d8b4fe">⭐ Total Action Progress (All Priority Sites)</span>
          <span style="font-size:12px;color:#60a5fa;font-weight:700">${{allProg}} / ${{allTotal}} Site Terisi (${{allProgPct}}%)</span>
        </div>
        <div style="height:18px;background:rgba(255,255,255,0.06);border-radius:6px;overflow:hidden;display:flex">
          <div style="width:${{allProgPct}}%;background:linear-gradient(90deg, #3b82f6, #2563eb);display:flex;align-items:center;justify-content:center;color:#fff;font-size:10px;font-weight:700" title="Sudah Ada Action: ${{allProg}} (${{allProgPct}}%)">${{allProgPct}}% Terisi</div>
          <div style="width:${{allPendingPct}}%;background:rgba(239,68,68,0.25);display:flex;align-items:center;justify-content:center;color:#fca5a5;font-size:10px;font-weight:600" title="Pending: ${{allPending}} (${{allPendingPct}}%)">${{allPending}} Pending</div>
        </div>
      </div>

      <div style="margin-top:14px;padding-top:12px;border-top:1px dashed rgba(255,255,255,0.1)">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px">
          <span style="font-size:11px;color:var(--text-muted);font-weight:700;text-transform:uppercase;letter-spacing:0.5px">Rincian Part Priority Network (Klik untuk filter):</span>
          <span style="font-size:10.5px;color:#d8b4fe">Filter Part Network</span>
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:6px">
          <div class="prio-part-chip" onclick="setNetPartFilter('PART_SQA')" style="background:rgba(239,68,68,0.08);border:1px solid rgba(239,68,68,0.25);border-radius:6px;padding:6px 9px;cursor:pointer;display:flex;justify-content:space-between;font-size:11.5px" title="Klik untuk filter ke Part SQA">
            <span style="color:#fca5a5">🎯 Part SQA:</span>
            <span style="font-weight:700;color:#f87171">${{netSqa}} / ${{nTotal}}</span>
          </div>
          <div class="prio-part-chip" onclick="setNetPartFilter('PART_RHI')" style="background:rgba(167,139,250,0.08);border:1px solid rgba(167,139,250,0.25);border-radius:6px;padding:6px 9px;cursor:pointer;display:flex;justify-content:space-between;font-size:11.5px" title="Klik untuk filter ke Part RHI">
            <span style="color:#c4b5fd">📡 Part RHI:</span>
            <span style="font-weight:700;color:#a78bfa">${{netRhi}} / ${{nTotal}}</span>
          </div>
          <div class="prio-part-chip" onclick="setNetPartFilter('PART_THI')" style="background:rgba(56,189,248,0.08);border:1px solid rgba(56,189,248,0.25);border-radius:6px;padding:6px 9px;cursor:pointer;display:flex;justify-content:space-between;font-size:11.5px" title="Klik untuk filter ke Part THI">
            <span style="color:#7dd3fc">📶 Part THI:</span>
            <span style="font-weight:700;color:#38bdf8">${{netThi}} / ${{nTotal}}</span>
          </div>
          <div class="prio-part-chip" onclick="setNetPartFilter('PART_AVAIL')" style="background:rgba(234,179,8,0.08);border:1px solid rgba(234,179,8,0.25);border-radius:6px;padding:6px 9px;cursor:pointer;display:flex;justify-content:space-between;font-size:11.5px" title="Klik untuk filter ke Part Availability">
            <span style="color:#fde047">⚡ Part Avail:</span>
            <span style="font-weight:700;color:#eab308">${{netAvail}} / ${{nTotal}}</span>
          </div>
          <div class="prio-part-chip" onclick="setNetPartFilter('PART_PL')" style="background:rgba(59,130,246,0.08);border:1px solid rgba(59,130,246,0.25);border-radius:6px;padding:6px 9px;cursor:pointer;display:flex;justify-content:space-between;font-size:11.5px" title="Klik untuk filter ke Part PL / MBA">
            <span style="color:#93c5fd">🔌 Part PL / MBA:</span>
            <span style="font-weight:700;color:#60a5fa">${{netPl}} / ${{nTotal}}</span>
          </div>
          <div class="prio-part-chip" onclick="setNetPartFilter('PART_COST')" style="background:rgba(20,184,166,0.08);border:1px solid rgba(20,184,166,0.25);border-radius:6px;padding:6px 9px;cursor:pointer;display:flex;justify-content:space-between;font-size:11.5px" title="Klik untuk filter ke Part Cost">
            <span style="color:#5eead4">💰 Part Cost:</span>
            <span style="font-weight:700;color:#14b8a6">${{netCost}} / ${{nTotal}}</span>
          </div>
        </div>

        <div style="margin-top:10px;padding-top:8px;border-top:1px dashed rgba(255,255,255,0.08);display:flex;justify-content:space-between;font-size:11.5px;color:var(--text-muted)">
          <span>💼 Priority Sales Action Breakdown:</span>
          <span style="color:#fcd34d">NOP: <strong>${{sNop}}</strong> | Sales: <strong>${{sSales}}</strong></span>
        </div>
      </div>
    `;
    containerProg.innerHTML = progHtml;
  }}

  function renderProgCharts() {{
    const total = filteredProgRows.length;
    const badge = document.getElementById('prog-chart-badge');
    if (badge) {{
      badge.innerText = `${{total.toLocaleString()}} Sites Filtered`;
    }}

    const planContainer = document.getElementById('prog-plan-bars');
    const domainContainer = document.getElementById('prog-domain-bars');
    const branchContainer = document.getElementById('prog-branch-bars');
    const chipsContainer = document.getElementById('prog-kpi-chips');
    if (!domainContainer || !branchContainer || !chipsContainer) return;

    if (total === 0) {{
      if (planContainer) planContainer.innerHTML = '<div style="color:var(--text-muted);font-size:12px;padding:15px;text-align:center">Tidak ada data untuk ditampilkan</div>';
      domainContainer.innerHTML = '<div style="color:var(--text-muted);font-size:12px;padding:15px;text-align:center">Tidak ada data untuk ditampilkan</div>';
      branchContainer.innerHTML = '<div style="color:var(--text-muted);font-size:12px;padding:15px;text-align:center">Tidak ada data untuk ditampilkan</div>';
      chipsContainer.innerHTML = '';
      return;
    }}

    // 1. Rencana Aksi (Action to be taken)
    let planSqa = 0, planSales = 0, planNop = 0, planRhi = 0;
    let planThi = 0, planAvail = 0, planPl = 0, planNos = 0;

    // 2. Realisasi Progres (Update action taken & propose action)
    let doneSqa = 0, doneSales = 0, doneNop = 0, doneRhi = 0;
    let doneThi = 0, doneAvail = 0, donePl = 0, doneNos = 0;

    let plClearCnt = 0;
    let anyActionCnt = 0;

    filteredProgRows.forEach(r => {{
      // Plan counts
      const hasPlanSqa = isFilledVal(r[20]);   // Col AH: Action to be taken SQA
      const hasPlanSales = isFilledVal(r[22]); // Col AJ: Action to be taken Sales
      const hasPlanNop = isFilledVal(r[24]);   // Col AL: Action to be taken NOP
      const hasPlanRhi = isFilledVal(r[27]);   // Col AO: Root Cause RHI
      const hasPlanThi = isFilledVal(r[30]);   // Col AR: Root Cause THI
      const hasPlanAvail = isFilledVal(r[34]); // Col AV: update average status availability
      const hasPlanPl = isFilledVal(r[36]);    // Col AX: Status PL week terakhir
      const hasPlanNos = isFilledVal(r[38]);   // Col AZ: Breakdown cost terbesar

      if (hasPlanSqa) planSqa++;
      if (hasPlanSales) planSales++;
      if (hasPlanNop) planNop++;
      if (hasPlanRhi) planRhi++;
      if (hasPlanThi) planThi++;
      if (hasPlanAvail) planAvail++;
      if (hasPlanPl) planPl++;
      if (hasPlanNos) planNos++;

      // Realisasi / Done counts (User requested 8 items)
      const hasDoneSqa = isFilledVal(r[21]);   // Col AI: Update action taken SQA
      const hasDoneSales = isFilledVal(r[23]); // Col AK: Update action taken Sales
      const hasDoneNop = isFilledVal(r[25]);   // Col AM: Update action taken NOP
      const hasDoneRhi = isFilledVal(r[28]);   // Col AP: propose action RHI
      const hasDoneThi = isFilledVal(r[31]);   // Col AS: propose action THI
      const hasDoneAvail = isFilledVal(r[35]); // Col AW: plan action availability
      const hasDonePl = isFilledVal(r[37]);    // Col AY: propose action PL
      const hasDoneNos = isFilledVal(r[39]);   // Col BB: plan improvement cost terbesar

      if (hasDoneSqa) doneSqa++;
      if (hasDoneSales) doneSales++;
      if (hasDoneNop) doneNop++;
      if (hasDoneRhi) doneRhi++;
      if (hasDoneThi) doneThi++;
      if (hasDoneAvail) doneAvail++;
      if (hasDonePl) donePl++;
      if (hasDoneNos) doneNos++;

      const isPlClear = (String(r[36]).trim() === 'CLEAR');
      if (isPlClear) plClearCnt++;

      if (hasDoneSqa || hasDoneSales || hasDoneNop || hasDoneRhi || hasDoneThi || hasDoneAvail || hasDonePl || hasDoneNos ||
          hasPlanSqa || hasPlanSales || hasPlanNop || hasPlanRhi || hasPlanThi || hasPlanAvail || hasPlanPl || hasPlanNos) {{
        anyActionCnt++;
      }}
    }});

    const anyPct = (anyActionCnt / total * 100).toFixed(1);
    const plClearPct = (plClearCnt / total * 100).toFixed(1);

    // KPI Chips with clean labels (without parentheses)
    chipsContainer.innerHTML = `
      <div class="prog-chip" style="background:rgba(16,185,129,0.12);border-color:rgba(16,185,129,0.3)">
        <span style="color:#34d399">🎯 Action Plan Terisi:</span>
        <strong style="color:#34d399">${{anyActionCnt}}</strong>
        <span style="font-size:11px;color:var(--text-muted)">(${{anyPct}}%)</span>
      </div>
      <div class="prog-chip" style="background:rgba(239,68,68,0.12);border-color:rgba(239,68,68,0.25)">
        <span style="color:#f87171">SQA:</span>
        <strong style="color:#f87171">${{doneSqa}}</strong>
        <span style="font-size:10.5px;color:var(--text-dim)">/${{planSqa}} plan</span>
      </div>
      <div class="prog-chip" style="background:rgba(245,158,11,0.12);border-color:rgba(245,158,11,0.25)">
        <span style="color:#fbbf24">Sales:</span>
        <strong style="color:#fbbf24">${{doneSales}}</strong>
        <span style="font-size:10.5px;color:var(--text-dim)">/${{planSales}} plan</span>
      </div>
      <div class="prog-chip" style="background:rgba(16,185,129,0.12);border-color:rgba(16,185,129,0.25)">
        <span style="color:#34d399">NOP:</span>
        <strong style="color:#34d399">${{doneNop}}</strong>
        <span style="font-size:10.5px;color:var(--text-dim)">/${{planNop}} plan</span>
      </div>
      <div class="prog-chip" style="background:rgba(139,92,246,0.12);border-color:rgba(139,92,246,0.25)">
        <span style="color:#c084fc">RHI:</span>
        <strong style="color:#c084fc">${{doneRhi}}</strong>
      </div>
      <div class="prog-chip" style="background:rgba(236,72,153,0.12);border-color:rgba(236,72,153,0.25)">
        <span style="color:#f472b6">THI:</span>
        <strong style="color:#f472b6">${{doneThi}}</strong>
      </div>
      <div class="prog-chip" style="background:rgba(234,179,8,0.12);border-color:rgba(234,179,8,0.25)">
        <span style="color:#fde047">Avail:</span>
        <strong style="color:#fde047">${{doneAvail}}</strong>
      </div>
      <div class="prog-chip" style="background:rgba(59,130,246,0.12);border-color:rgba(59,130,246,0.25)">
        <span style="color:#60a5fa">PL:</span>
        <strong style="color:#60a5fa">${{donePl}}</strong>
      </div>
      <div class="prog-chip" style="background:rgba(6,182,212,0.12);border-color:rgba(6,182,212,0.25)">
        <span style="color:#38bdf8">NOS:</span>
        <strong style="color:#38bdf8">${{doneNos}}</strong>
      </div>
      <div class="prog-chip" style="background:rgba(59,130,246,0.12);border-color:rgba(59,130,246,0.3)">
        <span style="color:#60a5fa">🔌 Status PL CLEAR:</span>
        <strong style="color:#60a5fa">${{plClearCnt}}</strong>
        <span style="font-size:11px;color:var(--text-muted)">(${{plClearPct}}%)</span>
      </div>
      <div class="prog-chip" style="cursor:pointer;background:rgba(239,68,68,0.1);border-color:rgba(239,68,68,0.25)" onclick="setProgOAFilter('Not OA')" title="Klik untuk filter site Not OA">
        <span style="color:#f87171">⚠️ Not OA:</span>
        <strong style="color:#f87171">${{filteredProgRows.filter(r => r[8] !== 'OA').length}}</strong>
      </div>
    `;

    // 1. Box Rencana Aksi (Action to be taken)
    const planData = [
      {{ name: 'Action to be taken SQA', count: planSqa, color: '#ef4444', gradient: 'linear-gradient(90deg, #f87171, #dc2626)', dot: '#ef4444' }},
      {{ name: 'Action to be taken Sales', count: planSales, color: '#f59e0b', gradient: 'linear-gradient(90deg, #fcd34d, #d97706)', dot: '#f59e0b' }},
      {{ name: 'Action to be taken NOP', count: planNop, color: '#10b981', gradient: 'linear-gradient(90deg, #34d399, #059669)', dot: '#10b981' }},
      {{ name: 'RHI Root Cause', count: planRhi, color: '#8b5cf6', gradient: 'linear-gradient(90deg, #a78bfa, #7c3aed)', dot: '#8b5cf6' }},
      {{ name: 'THI Root Cause', count: planThi, color: '#ec4899', gradient: 'linear-gradient(90deg, #f472b6, #db2777)', dot: '#ec4899' }},
      {{ name: 'Status Availability', count: planAvail, color: '#eab308', gradient: 'linear-gradient(90deg, #fde047, #ca8a04)', dot: '#eab308' }},
      {{ name: 'Status PL week terakhir', count: planPl, color: '#3b82f6', gradient: 'linear-gradient(90deg, #60a5fa, #2563eb)', dot: '#3b82f6' }},
      {{ name: 'Breakdown cost terbesar', count: planNos, color: '#06b6d4', gradient: 'linear-gradient(90deg, #38bdf8, #0891b2)', dot: '#06b6d4' }}
    ];

    if (planContainer) {{
      let planHtml = '';
      planData.forEach(d => {{
        const pct = total > 0 ? (d.count / total * 100).toFixed(1) : '0.0';
        const isZero = (d.count === 0);
        planHtml += `
          <div class="hbar-item">
            <div class="hbar-header">
              <span class="hbar-label">
                <span style="width:8px;height:8px;border-radius:50%;background:${{d.dot}};display:inline-block"></span>
                ${{d.name}}
              </span>
              <span class="hbar-value" style="color:${{isZero ? 'var(--text-dim)' : d.color}};font-weight:700">
                ${{d.count.toLocaleString()}} <span style="font-size:11px;color:var(--text-dim);font-weight:normal">(${{pct}}%)</span>
              </span>
            </div>
            <div class="hbar-track">
              <div class="hbar-fill" style="width:${{Math.max(parseFloat(pct), isZero ? 0 : 2)}}%;background:${{d.gradient}};opacity:${{isZero ? 0.25 : 1}}"></div>
            </div>
          </div>
        `;
      }});
      planContainer.innerHTML = planHtml;
    }}

    // 2. Box Realisasi Progres (Update action taken & propose action) - CLEAN WITHOUT PARENTHESES
    const domainData = [
      {{ name: 'Update action taken SQA', count: doneSqa, color: '#ef4444', gradient: 'linear-gradient(90deg, #f87171, #dc2626)', dot: '#ef4444' }},
      {{ name: 'Update action taken Sales', count: doneSales, color: '#f59e0b', gradient: 'linear-gradient(90deg, #fcd34d, #d97706)', dot: '#f59e0b' }},
      {{ name: 'Update action taken NOP', count: doneNop, color: '#10b981', gradient: 'linear-gradient(90deg, #34d399, #059669)', dot: '#10b981' }},
      {{ name: 'propose action RHI', count: doneRhi, color: '#8b5cf6', gradient: 'linear-gradient(90deg, #a78bfa, #7c3aed)', dot: '#8b5cf6' }},
      {{ name: 'propose action THI', count: doneThi, color: '#ec4899', gradient: 'linear-gradient(90deg, #f472b6, #db2777)', dot: '#ec4899' }},
      {{ name: 'plan action availability', count: doneAvail, color: '#eab308', gradient: 'linear-gradient(90deg, #fde047, #ca8a04)', dot: '#eab308' }},
      {{ name: 'propose action PL', count: donePl, color: '#3b82f6', gradient: 'linear-gradient(90deg, #60a5fa, #2563eb)', dot: '#3b82f6' }},
      {{ name: 'plan improvement cost terbesar', count: doneNos, color: '#06b6d4', gradient: 'linear-gradient(90deg, #38bdf8, #0891b2)', dot: '#06b6d4' }}
    ];

    let domainHtml = '';
    domainData.forEach(d => {{
      const pct = total > 0 ? (d.count / total * 100).toFixed(1) : '0.0';
      const isZero = (d.count === 0);
      domainHtml += `
        <div class="hbar-item">
          <div class="hbar-header">
            <span class="hbar-label">
              <span style="width:8px;height:8px;border-radius:50%;background:${{d.dot}};display:inline-block"></span>
              ${{d.name}}
            </span>
            <span class="hbar-value" style="color:${{isZero ? 'var(--text-dim)' : d.color}};font-weight:700">
              ${{d.count.toLocaleString()}} <span style="font-size:11px;color:var(--text-dim);font-weight:normal">(${{pct}}%)</span>
            </span>
          </div>
          <div class="hbar-track">
            <div class="hbar-fill" style="width:${{Math.max(parseFloat(pct), isZero ? 0 : 2)}}%;background:${{d.gradient}};opacity:${{isZero ? 0.25 : 1}}"></div>
          </div>
        </div>
      `;
    }});
    domainContainer.innerHTML = domainHtml;

    // 3. Branch Breakdown
    const branchConfigs = [
      {{ name: 'AMBON', label: 'Ambon', color: '#f59e0b', gradient: 'linear-gradient(90deg, #f59e0b, #d97706)' }},
      {{ name: 'JAYAPURA', label: 'Jayapura', color: '#38bdf8', gradient: 'linear-gradient(90deg, #38bdf8, #0284c7)' }},
      {{ name: 'SORONG', label: 'Sorong', color: '#10b981', gradient: 'linear-gradient(90deg, #10b981, #059669)' }},
      {{ name: 'TIMIKA', label: 'Timika', color: '#a855f7', gradient: 'linear-gradient(90deg, #a855f7, #7e22ce)' }}
    ];

    let branchHtml = '';
    branchConfigs.forEach(b => {{
      const bRows = filteredProgRows.filter(r => r[3] === b.name);
      const bTotal = bRows.length;
      let bWithAction = 0;
      let bSqa = 0, bSales = 0, bNop = 0;
      bRows.forEach(r => {{
        const hasSqa = isFilledVal(r[21]);
        const hasSales = isFilledVal(r[23]);
        const hasNop = isFilledVal(r[25]);
        const hasOther = isFilledVal(r[28]) || isFilledVal(r[31]) || isFilledVal(r[35]) || isFilledVal(r[37]) || isFilledVal(r[39]);
        if (hasSqa) bSqa++;
        if (hasSales) bSales++;
        if (hasNop) bNop++;
        if (hasSqa || hasSales || hasNop || hasOther) {{
          bWithAction++;
        }}
      }});
      const bPct = bTotal > 0 ? (bWithAction / bTotal * 100).toFixed(1) : '0.0';

      branchHtml += `
        <div class="branch-bar-item" onclick="setProgBranchFilter('${{b.name}}')" title="Klik untuk filter tabel ke Branch ${{b.name}}">
          <div class="branch-bar-name">
            <span style="width:8px;height:8px;border-radius:50%;background:${{b.color}};display:inline-block"></span>
            ${{b.label}}
          </div>
          <div class="branch-bar-track-wrap">
            <div class="branch-bar-meta">
              <span><strong>${{bWithAction}}</strong> dari ${{bTotal}} site terisi action (SQA: ${{bSqa}} | Sales: ${{bSales}} | NOP: ${{bNop}})</span>
              <span>${{bTotal - bWithAction}} belum terisi</span>
            </div>
            <div class="branch-bar-track">
              <div class="branch-bar-filled" style="width:${{bPct}}%;background:${{b.gradient}}"></div>
            </div>
          </div>
          <div class="branch-bar-pct" style="color:${{bPct > 50 ? '#34d399' : (bPct > 20 ? '#fbbf24' : '#f87171')}}">
            ${{bPct}}%
          </div>
        </div>
      `;
    }});
    branchContainer.innerHTML = branchHtml;
  }}


  function setProgBranchFilter(branchName) {{
    const select = document.getElementById('prog-filter-nop');
    if (select) {{
      select.value = (select.value === branchName) ? 'ALL' : branchName;
      applyProgFilters();
    }}
  }}

  function setProgOAFilter(oaVal) {{
    const select = document.getElementById('prog-filter-oa');
    if (select) {{
      select.value = (select.value === oaVal) ? 'ALL' : oaVal;
      applyProgFilters();
    }}
  }}

  function exportProgressToCSV() {{
    const headers = [
      'Site ID', 'Site Name', 'Branch', 'Kabupaten', 'Status PnL',
      'Defisit / Margin (Rp)', 'Cost (Rp)',
      'Site Class', 'Status OA', 'Site Loss Tracker (Kolom AG)',
      'Priority Sales', 'Priority Network',
      'Action to be taken SQA', 'Update action taken sqa',
      'Action to be taken Sales', 'Update action taken sales',
      'Action to be taken NOP', 'Update action taken NOP',
      'update RHI Week terakhir', 'Root Cause RHI', 'propose action RHI',
      'update THI Week terakhir', 'Root Cause THI', 'propose action THI',
      'Typer Transport', 'utilisasi average weekly BDBH',
      'update average status availability', 'plan action availability',
      'Status PL week terakhir', 'propose action (MBA)',
      'Breakdown cost terbesar', 'Nilai Cost', 'plan improvement cost terbesar',
      'Cek / Kategori Site'
    ];

    let csv = [];
    csv.push(headers.map(h => `"${{h}}"`).join(','));

    filteredProgRows.forEach(r => {{
      const rowVals = [
        r[0], r[1], r[3], r[4], r[13],
        r[11], r[10],
        r[7], r[8], r[42],
        r[17], r[18],
        r[20], r[21], r[22], r[23], r[24], r[25],
        r[26], r[27], r[28],
        r[29], r[30], r[31],
        r[32], r[33], r[34], r[35],
        r[36], r[37],
        r[38], r[44], r[39],
        r[43]
      ];
      csv.push(rowVals.map(val => `"${{String(val || '').replace(/"/g, '""')}}"`).join(','));
    }});

    const blob = new Blob([csv.join('\\\\n')], {{ type: 'text/csv;charset=utf-8;' }});
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `Progress_Action_Tracker_${{new Date().toISOString().slice(0,10)}}.csv`;
    link.click();
  }}

  // --- DATABASE AUG 26 TABLE LOGIC ---
  let filteredAugRows = [...augData];
  let curPage = 1;
  let pageSize = 50;

  function onAugSearch() {{
    applyAugFilters();
  }}

  function onAugSortChange(val) {{
    if (val === 'NONE') {{
      augSortCol = null;
      augSortDir = 'DESC';
    }} else if (val === 'DEFISIT_DESC') {{
      augSortCol = 'defisit';
      augSortDir = 'DESC';
    }} else if (val === 'DEFISIT_ASC') {{
      augSortCol = 'defisit';
      augSortDir = 'ASC';
    }} else if (val === 'COST_DESC') {{
      augSortCol = 'cost';
      augSortDir = 'DESC';
    }} else if (val === 'REV_DESC') {{
      augSortCol = 'rev';
      augSortDir = 'DESC';
    }} else if (val === 'NILAI_COST_DESC') {{
      augSortCol = 'nilai_cost';
      augSortDir = 'DESC';
    }} else if (val === 'SITE_ID_ASC') {{
      augSortCol = 'site_id';
      augSortDir = 'ASC';
    }}
    applyAugFilters();
  }}

  function sortAugTable(colKey) {{
    if (augSortCol === colKey) {{
      augSortDir = (augSortDir === 'DESC') ? 'ASC' : 'DESC';
    }} else {{
      augSortCol = colKey;
      augSortDir = 'DESC';
    }}
    const sel = document.getElementById('aug-sort-select');
    if (sel) {{
      if (colKey === 'defisit') sel.value = augSortDir === 'DESC' ? 'DEFISIT_DESC' : 'DEFISIT_ASC';
      else if (colKey === 'cost') sel.value = 'COST_DESC';
      else if (colKey === 'rev') sel.value = 'REV_DESC';
      else if (colKey === 'nilai_cost') sel.value = 'NILAI_COST_DESC';
      else if (colKey === 'site_id') sel.value = 'SITE_ID_ASC';
      else sel.value = 'NONE';
    }}
    applyAugFilters();
  }}

  function applyAugFilters() {{
    const q = (document.getElementById('aug-search-input').value || '').toLowerCase().trim();
    const nopVal = document.getElementById('aug-filter-nop').value;
    const pnlVal = document.getElementById('aug-filter-pnl').value;
    const classVal = document.getElementById('aug-filter-class').value;
    const oaVal = document.getElementById('aug-filter-oa').value;
    const agVal = document.getElementById('aug-filter-ag') ? document.getElementById('aug-filter-ag').value : 'ALL';
    const gapVal = document.getElementById('aug-filter-gap').value;
    const salesPrioVal = document.getElementById('aug-filter-priority-sales') ? document.getElementById('aug-filter-priority-sales').value : 'ALL';
    const netPrioVal = document.getElementById('aug-filter-priority-net') ? document.getElementById('aug-filter-priority-net').value : 'ALL';
    const cekVal = document.getElementById('aug-filter-cek') ? document.getElementById('aug-filter-cek').value : 'ALL';

    filteredAugRows = augData.filter(row => {{
      if (q) {{
        const text = (row[0] + ' ' + row[1] + ' ' + row[2] + ' ' + row[3] + ' ' + row[4] + ' ' + row[17] + ' ' + row[18] + ' ' + row[20] + ' ' + row[21] + ' ' + (row[43] || '')).toLowerCase();
        if (!text.includes(q)) return false;
      }}
      if (nopVal !== 'ALL' && row[2] !== nopVal) return false;
      if (pnlVal !== 'ALL' && row[13] !== pnlVal) return false;
      if (classVal !== 'ALL' && row[7] !== classVal) return false;
      if (oaVal !== 'ALL' && row[8] !== oaVal) return false;
      if (cekVal !== 'ALL' && row[43] !== cekVal) return false;
      if (gapVal !== 'ALL') {{
        if (gapVal === 'Profit' && row[14] !== 'Profit') return false;
        if (gapVal !== 'Profit' && row[14] !== gapVal) return false;
      }}

      // Priority Filters
      const isSalesPrio = (row[17] === 'P1' || row[17] === 'P2' || row[17] === 'P3');
      const isNetPrio = (row[18] === 'P1' || row[18] === 'P2' || row[18] === 'P3');

      if (salesPrioVal !== 'ALL') {{
        if (salesPrioVal === 'PRIORITY_ONLY' && !isSalesPrio) return false;
        if (salesPrioVal !== 'PRIORITY_ONLY' && row[17] !== salesPrioVal) return false;
      }}
      if (netPrioVal !== 'ALL') {{
        if (netPrioVal === 'PRIORITY_ONLY' && !isNetPrio) return false;
        if (netPrioVal !== 'PRIORITY_ONLY' && row[18] !== netPrioVal) return false;
      }}
      if (agVal !== 'ALL') {{
        if (agVal === 'Program H2' && row[42] !== 'Program H2') return false;
        if (agVal === '0' && row[42] !== '0') return false;
        if (agVal === '#N/A' && row[42] !== '#N/A') return false;
      }}

      return true;
    }});

    // Apply Sorting (Besar ke Kecil / Shorting)
    if (augSortCol && augSortCol !== 'NONE') {{
      filteredAugRows = sortRows(filteredAugRows, augSortCol, augSortDir);
    }}

    curPage = 1;
    renderAugTable();
  }}

  function changePageSize(val) {{
    pageSize = parseInt(val) || 50;
    curPage = 1;
    renderAugTable();
  }}

  function goToPage(p) {{
    curPage = p;
    renderAugTable();
  }}

  function prevPage() {{
    if (curPage > 1) {{
      curPage--;
      renderAugTable();
    }}
  }}

  function nextPage() {{
    const maxPage = Math.ceil(filteredAugRows.length / pageSize) || 1;
    if (curPage < maxPage) {{
      curPage++;
      renderAugTable();
    }}
  }}

  function goToLastPage() {{
    const maxPage = Math.ceil(filteredAugRows.length / pageSize) || 1;
    curPage = maxPage;
    renderAugTable();
  }}

  function renderAugTable() {{
    const tbody = document.getElementById('aug-table-body');
    const total = filteredAugRows.length;
    const totalPages = Math.ceil(total / pageSize) || 1;

    if (curPage > totalPages) curPage = totalPages;
    if (curPage < 1) curPage = 1;

    const startIdx = (curPage - 1) * pageSize;
    const endIdx = Math.min(startIdx + pageSize, total);

    document.getElementById('current-page-num').innerText = curPage;
    document.getElementById('total-pages-num').innerText = totalPages;
    document.getElementById('btn-prev').disabled = (curPage === 1);
    document.getElementById('btn-first').disabled = (curPage === 1);
    document.getElementById('btn-next').disabled = (curPage === totalPages);
    document.getElementById('btn-last').disabled = (curPage === totalPages);

    document.getElementById('aug-pagination-info').innerText = total === 0 
      ? 'Tidak ada data ditemukan' 
      : `Menampilkan ${{startIdx + 1}} - ${{endIdx}} dari ${{total.toLocaleString()}} site`;

    if (total === 0) {{
      tbody.innerHTML = `<tr><td colspan="45" style="text-align:center;padding:40px;color:var(--text-muted)">Tidak ada data site yang sesuai dengan kriteria filter/pencarian</td></tr>`;
      return;
    }}

    let rowsHtml = '';
    for (let i = startIdx; i < endIdx; i++) {{
      const r = filteredAugRows[i];
      const isLoss = (r[13] === 'Loss');
      const pnlBadge = isLoss 
        ? `<span class="badge badge-danger">Loss</span>` 
        : `<span class="badge badge-success">Profit</span>`;

      let gapBadge = r[14];
      if (r[14] === '>10 Mio') gapBadge = `<span class="badge badge-danger">>10 Mio</span>`;
      else if (r[14] === '5-10 Mio') gapBadge = `<span class="badge badge-warning">5-10 Mio</span>`;
      else if (r[14] === '3-5 Mio' || r[14] === '<3 Mio') gapBadge = `<span class="badge badge-info">${{r[14]}}</span>`;

      const ssgjBadge = (r[15] === 'Inline') 
        ? `<span class="badge badge-info">Inline</span>` 
        : `<span class="badge" style="background:rgba(255,255,255,0.06);color:var(--text-muted)">Not Inline</span>`;

      const intervBadge = (r[41] === 'Intervention')
        ? `<span class="badge badge-warning">Intervention</span>`
        : `<span class="badge" style="background:rgba(255,255,255,0.06);color:var(--text-muted)">Non</span>`;

      const sClass = r[7] || '-';
      let classBadge = `<span class="badge badge-silver">${{sClass}}</span>`;
      if (sClass === 'DIAMOND') classBadge = `<span class="badge badge-diamond">DIAMOND</span>`;
      else if (sClass === 'PLATINUM') classBadge = `<span class="badge badge-platinum">PLATINUM</span>`;
      else if (sClass === 'GOLD') classBadge = `<span class="badge badge-gold">GOLD</span>`;
      else if (sClass === 'BRONZE') classBadge = `<span class="badge badge-bronze">BRONZE</span>`;

      const sOA = r[8] || '-';
      const oaBadge = (sOA === 'OA') 
        ? `<span class="badge badge-success">OA</span>` 
        : `<span class="badge badge-danger">Not OA</span>`;

      const plWeekBadge = (r[36] === 'CLEAR')
        ? `<span class="badge badge-success">CLEAR</span>`
        : (r[33] ? `<span class="badge badge-warning">${{r[33]}}</span>` : '-');

      rowsHtml += `<tr>
        <td style="color:var(--text-dim);font-family:'JetBrains Mono'">${{i + 1}}</td>
        <td><strong style="color:${{isLoss ? '#f87171' : '#38bdf8'}};font-family:'JetBrains Mono'">${{r[0]}}</strong></td>
        <td title="${{r[1]}}" style="max-width:210px;overflow:hidden;text-overflow:ellipsis">${{r[1]}}</td>
        <td>${{r[2]}}</td>
        <td>${{r[3]}}</td>
        <td>${{r[4]}}</td>
        <td>${{classBadge}}</td>
        <td>${{oaBadge}}</td>
        <td>${{r[5]}}</td>
        <td>${{r[6]}}</td>
        <td style="text-align:right">${{Math.round(r[9]).toLocaleString()}}</td>
        <td style="text-align:right">${{Math.round(r[10]).toLocaleString()}}</td>
        <td style="text-align:right;color:${{r[11] < 0 ? '#f87171' : '#34d399'}};font-weight:700">${{Math.round(r[11]).toLocaleString()}}</td>
        <td style="text-align:right">${{r[12]}}%</td>
        <td>${{pnlBadge}}</td>
        <td>${{gapBadge}}</td>
        <td>${{ssgjBadge}}</td>
        <td style="color:var(--text-muted)">${{r[16] || '-'}}</td>
        <td>${{getPriorityBadge(r[17], 'sales')}}</td>
        <td>${{getPriorityBadge(r[18], 'net')}}</td>
        <td>${{intervBadge}}</td>
        <td><span class="badge" style="background:rgba(255,255,255,0.06);color:#94a3b8">${{r[40] || '-'}}</span></td>
        <td>${{r[42] === 'Program H2' ? '<span class="badge" style="background:rgba(245,158,11,0.2);color:#fbbf24;border:1px solid rgba(245,158,11,0.4)">Program H2</span>' : (r[42] || '-')}}</td>
        <td style="font-family:'JetBrains Mono'">${{r[19] || '-'}}</td>
        <!-- 18 ACTION / TRACKING COLUMNS (AF to AW) -->
        <td class="new-col-td" style="color:#fca5a5">${{r[20] || '-'}}</td>
        <td class="new-col-td" style="color:#fca5a5">${{r[21] || '-'}}</td>
        <td class="new-col-td" style="color:#fcd34d">${{r[22] || '-'}}</td>
        <td class="new-col-td" style="color:#fcd34d">${{r[23] || '-'}}</td>
        <td class="new-col-td" style="color:#6ee7b7">${{r[24] || '-'}}</td>
        <td class="new-col-td" style="color:#6ee7b7">${{r[25] || '-'}}</td>
        <td class="new-col-td" style="font-family:'JetBrains Mono'">${{r[26] ? (r[26] === 'Pass' ? '<span class="badge badge-success">Pass</span>' : (r[26] === 'Fail' ? '<span class="badge badge-danger">Fail</span>' : r[26])) : '-'}}</td>
        <td class="new-col-td" style="color:#c4b5fd">${{r[27] || '-'}}</td>
        <td class="new-col-td" style="color:#c4b5fd">${{r[28] || '-'}}</td>
        <td class="new-col-td" style="font-family:'JetBrains Mono'">${{r[29] ? (r[29] === 'Pass' ? '<span class="badge badge-success">Pass</span>' : (r[29] === 'Fail' ? '<span class="badge badge-danger">Fail</span>' : r[29])) : '-'}}</td>
        <td class="new-col-td" style="color:#7dd3fc">${{r[30] || '-'}}</td>
        <td class="new-col-td" style="color:#7dd3fc">${{r[31] || '-'}}</td>
        <td class="new-col-td">${{r[32] ? '<span class="badge badge-purple">' + r[32] + '</span>' : '-'}}</td>
        <td class="new-col-td" style="font-family:'JetBrains Mono'">${{r[33] || '-'}}</td>
        <td class="new-col-td">${{r[34] ? (r[34] === 'Achieved' ? '<span class="badge badge-success">Achieved</span>' : r[34]) : '-'}}</td>
        <td class="new-col-td" style="color:#fbbf24">${{r[35] || '-'}}</td>
        <td class="new-col-td">${{plWeekBadge}}</td>
        <td class="new-col-td" style="color:#93c5fd">${{r[37] || '-'}}</td>
        <!-- NOS (AX, AY, BB, BC) -->
        <td class="new-col-td" style="color:#38bdf8">${{r[38] || '-'}}</td>
        <td class="new-col-td" style="text-align:right;color:#38bdf8;font-family:'JetBrains Mono'">${{r[44] ? ('Rp ' + Number(r[44]).toLocaleString()) : '-'}}</td>
        <td class="new-col-td" style="color:#38bdf8">${{r[39] || '-'}}</td>
        <!-- Cek / Kategori Site -->
        <td class="new-col-td">${{r[43] ? ('<span class="badge" style="background:rgba(6,182,212,0.15);color:#22d3ee;border:1px solid rgba(6,182,212,0.3)">' + r[43] + '</span>') : '-'}}</td>
      </tr>`;
    }}

    tbody.innerHTML = rowsHtml;
    updateSortIndicators('aug', augSortCol, augSortDir);
  }}

  // Export Filtered Aug 26 To CSV
  function exportFilteredAugToCSV() {{
    const headers = [
      'Site ID', 'Site Name', 'NOP', 'Branch', 'Kabupaten', 'Kelas Site', 'Status OA', 'Type Site', 'BTS Equip',
      'Revenue Aug-26 (Rp)', 'Direct Cost (Rp)', 'Margin (Rp)', '% Margin', 'Status PnL',
      'GAP Dist', 'Inline SSGJ', 'Program Group', 'Priority Sales', 'Priority Network', 'Intervention Flag', 'Status Aug', 'Payload',
      'Action to be taken SQA', 'Update Action SQA',
      'Action to be taken Sales', 'Update Action Sales',
      'Action to be taken NOP', 'Update Action NOP',
      'update RHI Week Terakhir', 'Root Cause RHI', 'propose action RHI',
      'update THI Week Terakhir', 'Root Cause THI', 'propose action THI',
      'Typer Transport', 'utilisasi avg weekly BDBH',
      'update avg status availability', 'plan action availability',
      'Status PL week terakhir', 'propose action (MBA)',
      'Breakdown cost terbesar', 'Nilai Cost', 'plan improvement cost terbesar', 'Cek / Kategori Site'
    ];

    let csv = [];
    csv.push(headers.map(h => `"${{h}}"`).join(','));

    filteredAugRows.forEach(r => {{
      const cols = [
        r[0], r[1], r[2], r[3], r[4], r[7], r[8], r[5], r[6],
        Math.round(r[9]), Math.round(r[10]), Math.round(r[11]), r[12] + '%', r[13],
        r[14], r[15], r[16], r[17], r[18], r[41], r[40], r[19],
        r[20], r[21], r[22], r[23], r[24], r[25],
        r[26], r[27], r[28], r[29], r[30], r[31],
        r[32], r[33], r[34], r[35],
        r[36], r[37],
        r[38], r[44], r[39], r[43]
      ];
      csv.push(cols.map(val => `"${{String(val || '').replace(/"/g, '""')}}"`).join(','));
    }});

    const blob = new Blob([csv.join('\\n')], {{ type: 'text/csv;charset=utf-8;' }});
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `Database_Site_Aug26_${{new Date().toISOString().slice(0,10)}}.csv`;
    link.click();
  }}

  function exportTableToCSV(tableId, filename) {{
    const table = document.getElementById(tableId);
    let csv = [];
    const rows = table.querySelectorAll('tr');

    for (let i = 0; i < rows.length; i++) {{
      let row = [], cols = rows[i].querySelectorAll('td, th');
      for (let j = 0; j < cols.length; j++) {{
        let text = cols[j].innerText.replace(/"/g, '""').trim();
        row.push('"' + text + '"');
      }}
      csv.push(row.join(','));
    }}

    const csvFile = new Blob([csv.join('\\n')], {{ type: 'text/csv' }});
    const downloadLink = document.createElement('a');
    downloadLink.download = filename;
    downloadLink.href = window.URL.createObjectURL(csvFile);
    downloadLink.style.display = 'none';
    document.body.appendChild(downloadLink);
    downloadLink.click();
    document.body.removeChild(downloadLink);
  }}

  // --- REAL-TIME AUTO-REFRESH & INSTANT LOCAL ENGINE (0s DELAY) ---
  let refreshCountdown = 30;
  let isRefreshing = false;
  let isLocalEngine = false;
  const currentVersionHash = "{build_timestamp_hash}";

  function canAutoReload() {{
    const lastReload = parseInt(sessionStorage.getItem('last_auto_reload') || '0', 10);
    return (Date.now() - lastReload) > 10000;
  }}

  function triggerReloadWithBust() {{
    sessionStorage.setItem('last_auto_reload', Date.now());
    const url = new URL(window.location.href);
    url.searchParams.set('t', Date.now());
    window.location.href = url.toString();
  }}

  function detectLocalEngine() {{
    fetch('http://127.0.0.1:8092/api/version', {{ cache: 'no-store', signal: AbortSignal.timeout(600) }})
      .then(res => res.json())
      .then(data => {{
        if (data && data.version) {{
          isLocalEngine = true;
          const statusElem = document.getElementById('sync-status-label');
          if (statusElem) {{
            statusElem.innerHTML = `⚡ <strong style="color:#10b981">Local Engine (0s Delay)</strong>`;
          }}
          const localBtn = document.getElementById('btn-local-mode');
          if (localBtn && window.location.hostname !== 'localhost' && window.location.hostname !== '127.0.0.1') {{
            localBtn.style.display = 'inline-flex';
          }}
          // Only auto-reload locally if user is browsing localhost directly
          const isLocalPage = (window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1');
          if (isLocalPage && data.version !== currentVersionHash && canAutoReload()) {{
            showToast('⚡ Data Excel baru terdeteksi! Memperbarui...', true);
            setTimeout(triggerReloadWithBust, 300);
          }}
        }}
      }})
      .catch(() => {{
        isLocalEngine = false;
      }});
  }}

  function manualTriggerRefresh() {{
    if (isRefreshing) return;
    isRefreshing = true;

    const icon = document.getElementById('refresh-icon');
    if (icon) icon.classList.add('spin');

    const isLocalPage = (window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1');

    if (isLocalPage) {{
      showToast('⚡ Memperbarui data langsung dari Excel (0s delay)...', false);
      fetch('http://127.0.0.1:8092/api/trigger-sync')
        .then(() => {{
          setTimeout(triggerReloadWithBust, 400);
        }})
        .catch(() => {{
          triggerReloadWithBust();
        }});
      return;
    }}

    // Cloud GitHub Pages mode
    showToast('🔄 Mengecek data versi terbaru...', false);
    fetch('version.json?t=' + Date.now(), {{ cache: 'no-store' }})
      .then(res => res.json())
      .then(data => {{
        if (data && data.version && data.version !== currentVersionHash) {{
          showToast('✨ Memuat data versi terbaru...', true);
          setTimeout(triggerReloadWithBust, 300);
        }} else {{
          showToast('Data di web ini sudah paling update. Bila baru saja save di Excel, tunggu sebentar (cloud build ±30s) atau buka Mode 0s Delay.', true);
          setTimeout(() => {{
            if (icon) icon.classList.remove('spin');
            isRefreshing = false;
          }}, 1200);
        }}
      }})
      .catch(() => {{
        triggerReloadWithBust();
      }});
  }}

  function initAutoRefresh() {{
    // Check local engine every 5 seconds
    setInterval(detectLocalEngine, 5000);

    // Countdown and cloud poll every 30s
    setInterval(() => {{
      refreshCountdown--;
      const timerElem = document.getElementById('refresh-timer');
      if (timerElem && !isLocalEngine) {{
        timerElem.innerText = `${{refreshCountdown}}s`;
      }}
      if (refreshCountdown <= 0) {{
        refreshCountdown = 30;
        checkAutoUpdates();
      }}
    }}, 1000);

    // Initial check
    detectLocalEngine();
  }}

  function checkAutoUpdates() {{
    if (window.location.protocol.startsWith('http')) {{
      fetch('version.json?t=' + Date.now(), {{ cache: 'no-store' }})
        .then(res => {{
          if (!res.ok) throw new Error('Not found');
          return res.json();
        }})
        .then(data => {{
          if (data && data.version && data.version !== currentVersionHash && canAutoReload()) {{
            showToast('✨ Data Excel terbaru terdeteksi! Memperbarui...', true);
            setTimeout(triggerReloadWithBust, 800);
          }}
        }})
        .catch(() => {{}});
    }}
  }}

  function showToast(message, isSuccess = true) {{
    let toast = document.getElementById('live-toast');
    if (!toast) {{
      toast = document.createElement('div');
      toast.id = 'live-toast';
      toast.className = 'toast-notification';
      document.body.appendChild(toast);
    }}
    toast.innerHTML = isSuccess 
      ? `<svg width="18" height="18" fill="none" stroke="#10b981" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 13l4 4L19 7"/></svg><span>${{message}}</span>`
      : `<svg width="18" height="18" fill="none" stroke="#38bdf8" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"/></svg><span>${{message}}</span>`;
    toast.classList.add('show');
    setTimeout(() => {{
      toast.classList.remove('show');
    }}, 3500);
  }}

  function updateThemeButton() {{
    const btn = document.getElementById('btn-theme-toggle');
    if (!btn) return;
    const light = isLightTheme();
    btn.innerHTML = light ? '<span>🌙 Versi Gelap</span>' : '<span>☀️ Versi Putih</span>';
    btn.title = light ? 'Kembali ke Versi Gelap' : 'Ganti ke Versi Putih';
  }}

  function toggleTheme() {{
    const next = isLightTheme() ? 'dark' : 'light';
    document.documentElement.setAttribute('data-theme', next);
    try {{ localStorage.setItem('sqa_puma_theme', next); }} catch (e) {{}}
    try {{
      const u = new URL(window.location.href);
      u.searchParams.set('theme', next);
      window.history.replaceState({{}}, '', u.toString());
    }} catch (e) {{}}
    updateThemeButton();
    renderTrendChart();
    if (typeof activeBranch !== 'undefined') updateSideCard(activeBranch);
    if (typeof renderProgCharts === 'function') renderProgCharts();
    if (typeof renderPriorityCharts === 'function') renderPriorityCharts();
  }}

  updateThemeButton();

  window.addEventListener('load', () => {{
    updateThemeButton();
    initAutoRefresh();
    try {{
      const pParam = (new URLSearchParams(window.location.search).get('period') || '').toLowerCase();
      if (pParam === 'aug' || pParam === 'aug26' || pParam === 'agustus') {{
        switchPeriod('AUG26', true);
      }} else {{
        switchPeriod('SEP26', true);
      }}
    }} catch (e) {{
      switchPeriod('SEP26', true);
    }}
  }});
  window.addEventListener('resize', renderTrendChart);
</script>


  <!-- Floating Action Button for Mobile -->
  <button id="mobile-back-to-top" class="mobile-fab" onclick="window.scrollTo({{top:0, behavior:'smooth'}})">
    <svg width="16" height="16" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 10l7-7m0 0l7 7m-7-7v18"/></svg>
    <span>Ke Atas</span>
  </button>

</body>
</html>
"""

output_path = r"C:\Users\Zulfahmi Asriadi\.gemini\antigravity-ide\scratch\pivot_master_dashboard.html"
with open(output_path, "w", encoding="utf-8") as f:
    f.write(html_template)


scratch_dir = r"C:\Users\Zulfahmi Asriadi\.gemini\antigravity-ide\scratch"
version_info = {
    "version": build_timestamp_hash,
    "last_updated": current_time_str,
    "rows": len(sep_rows),
    "sep_loss": sep_meta['branchStats']['ALL']['loss_num'],
    "aug_loss": aug_meta['branchStats']['ALL']['loss_num'],
    "branches": ["AMBON", "JAYAPURA", "SORONG", "TIMIKA"]
}
with open(os.path.join(scratch_dir, "version.json"), "w", encoding="utf-8") as vf:
    json.dump(version_info, vf, indent=2)
repo_path = r"C:\Users\Zulfahmi Asriadi\sqa-puma-site-loss"
if os.path.exists(repo_path):
    with open(os.path.join(repo_path, "version.json"), "w", encoding="utf-8") as vf:
        json.dump(version_info, vf, indent=2)

print(f"Generated complete dashboard HTML with Progress Tab in Branch Focus: {len(html_template)} bytes")
