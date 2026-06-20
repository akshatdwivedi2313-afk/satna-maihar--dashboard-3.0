import os, re, json, math, sys, datetime
from urllib.parse import urlencode
import requests
import pandas as pd
from bs4 import BeautifulSoup

BASE = 'https://jgsa.nregsmp.org'
BLOCKS = ['AMARPATAN','MAIHAR','MAJHGAWAN','NAGOD','RAMNAGAR','RAMPUR BAGHELAN','SATNA','UNCHAHARA']

# Business rule: Gap Filling in Plantation works are valid only up to FY 2021-2022.
# Any Gap Filling in Plantation work from FY 2022-2023 onward is excluded from dashboard data.
GAP_FILLING_MAX_FY_START = 2021

def fy_start_year(fin_year):
    m = re.search(r'(20\d{2})\s*-\s*(20\d{2})', str(fin_year or ''))
    return int(m.group(1)) if m else None

def is_gap_filling_after_allowed_fy(work):
    wt = norm(work.get('workType') or '')
    if 'GAP FILLING' not in wt or 'PLANTATION' not in wt:
        return False
    y = fy_start_year(work.get('finYear'))
    # If the FY is missing/unparseable, keep it rather than dropping uncertain legacy data.
    return y is not None and y > GAP_FILLING_MAX_FY_START

def filter_gap_filling_after_allowed_fy(works):
    kept, excluded = [], []
    for w in works:
        if is_gap_filling_after_allowed_fy(w):
            excluded.append(w)
        else:
            kept.append(w)
    return kept, excluded

DISTRICT_GROUPS = {
    'Satna': {'MAJHGAWAN','NAGOD','RAMPUR BAGHELAN','SATNA','UNCHAHARA'},
    'Maihar': {'AMARPATAN','RAMNAGAR','MAIHAR'}
}
def district_for_block(block):
    b = norm(block)
    for d, arr in DISTRICT_GROUPS.items():
        if b in arr:
            return d
    return 'Satna'
DATE = os.environ.get('JGSA_DATE') or datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=5, minutes=30))).strftime('%Y-%m-%d')
PREV_DATE = os.environ.get('JGSA_PREV_DATE') or '2026-06-01'
DISTRICT = 'SATNA'
OUT = os.environ.get('JGSA_OUT', 'jgsa_live_data.js')
ENG = os.environ.get('ENGNAME_FILE', 'engname.xlsx')
SESSION = requests.Session()
SESSION.headers.update({'User-Agent':'Mozilla/5.0 JGSA-Satna-Dashboard/3.0'})

def norm(s):
    return re.sub(r'\s+', ' ', str(s or '').strip()).upper()

def num(x):
    if x is None: return 0.0
    s = str(x)
    # remove commas, rupee, percent and Hindi/English words, keep signs/dots
    s = s.replace(',', '')
    m = re.findall(r'-?\d+(?:\.\d+)?', s)
    if not m: return 0.0
    try: return float(m[-1])
    except: return 0.0



def extract_fin_year_from_text(text):
    t = str(text or '')
    # JGSA Work Monitor shows FIN. YEAR like 2025-2026 / 2024-2025.
    # Restrict to normal FY ranges so long work-code IDs do not become fake years.
    m = re.search(r'(20(?:0[0-9]|1[0-9]|2[0-7]))\s*[-–]\s*(20(?:0[1-9]|1[0-9]|2[0-8]))', t)
    if m:
        return f"{m.group(1)}-{m.group(2)}"
    # Fallback only when an explicit campaign/work name year is visible, not from work-code numbers.
    m = re.search(r'(?:JGSA|JGS|जल\s*गंगा|अभियान)\D*(20(?:0[0-9]|1[0-9]|2[0-7]))', t, re.I)
    if m:
        y = int(m.group(1))
        return f"{y}-{y+1}"
    return ''

def get_html(url):
    r = SESSION.get(url, timeout=40)
    r.raise_for_status()
    return r.text

def read_tables_bs4(html):
    # Manual parser keeps columns like FIN. YEAR even when pandas drops/merges scrollable table cells.
    # IMPORTANT: use direct th/td children only. Official rankings cells contain nested markup;
    # recursive find_all() splits one category cell into its inner numbers (Started/Completed/Phys),
    # causing wrong values like Amrit Sarovar=456. Direct children preserve the category cell text.
    soup = BeautifulSoup(html, 'html.parser')
    tables = []
    for tbl in soup.find_all('table'):
        headers = []
        for tr in tbl.find_all('tr'):
            ths = tr.find_all('th', recursive=False)
            if ths:
                headers = [re.sub(r'\s+', ' ', th.get_text(' ', strip=True)).strip() for th in ths]
                break
        body_rows = []
        for tr in tbl.find_all('tr'):
            cells = tr.find_all('td', recursive=False)
            if not cells:
                continue
            row = [re.sub(r'\s+', ' ', td.get_text(' ', strip=True)).strip() for td in cells]
            if headers:
                if len(row) < len(headers):
                    row += [''] * (len(headers) - len(row))
                elif len(row) > len(headers):
                    row = row[:len(headers)]
            body_rows.append(row)
        if body_rows:
            if not headers:
                headers = [f'col_{i}' for i in range(max(len(r) for r in body_rows))]
            maxcols = max(len(r) for r in body_rows)
            hdr = headers + [f'col_{i}' for i in range(len(headers), maxcols)]
            body_rows = [r + ['']*(maxcols-len(r)) for r in body_rows]
            tables.append(pd.DataFrame(body_rows, columns=hdr[:maxcols]))
    return tables

def read_tables(html):
    tables = read_tables_bs4(html)
    if tables:
        return tables
    try:
        return pd.read_html(html, displayed_only=False)
    except Exception:
        try:
            return pd.read_html(html)
        except Exception:
            return []

def clean_df(df):
    df = df.copy()
    if hasattr(df.columns, 'to_flat_index'):
        df.columns = [' '.join([str(x) for x in tup if str(x) != 'nan']).strip() if isinstance(tup, tuple) else str(tup).strip() for tup in df.columns.to_flat_index()]
    else:
        df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how='all')
    for c in df.columns:
        df[c] = df[c].astype(str).replace({'nan':''})
    return df

def choose_table(tables, min_rows=1):
    if not tables: return pd.DataFrame()
    tables = [clean_df(t) for t in tables]
    tables = [t for t in tables if len(t) >= min_rows]
    if not tables: return pd.DataFrame()
    return max(tables, key=lambda d: len(d) * max(1, len(d.columns)))

def find_col(cols, keywords):
    norm_cols = [(c, norm(c)) for c in cols]
    # Exact / strong matches first, important for FIN. YEAR vs random YEAR text.
    for k in keywords:
        nk = norm(k)
        for c, nc in norm_cols:
            if nc == nk:
                return c
    for k in keywords:
        nk = norm(k)
        for c, nc in norm_cols:
            if nk in nc:
                return c
    return None


def choose_work_table(tables):
    """Pick the actual Work Monitor grid, preferring tables containing both Work Code and FIN. YEAR.
    The page may contain helper/filter tables, so the largest table is not always safest."""
    if not tables:
        return pd.DataFrame()
    cleaned=[clean_df(t) for t in tables if len(t) > 0]
    def score(df):
        cols=' | '.join(str(c) for c in df.columns).upper()
        sc=len(df)*2 + len(df.columns)
        if 'WORK CODE' in cols or 'कार्य कोड' in cols: sc += 100000
        if 'FIN. YEAR' in cols or 'FIN YEAR' in cols or 'FINANCIAL YEAR' in cols or 'वित्तीय' in cols: sc += 100000
        if 'PANCHAYAT' in cols or 'GRAM PANCHAYAT' in cols or 'GP' in cols or 'ग्राम' in cols: sc += 20000
        if 'WAGE SANCTIONED' in cols: sc += 5000
        if 'MATERIAL SANCTIONED' in cols: sc += 5000
        return sc
    return max(cleaned, key=score) if cleaned else pd.DataFrame()

def parse_work_rows(df, block):
    rows=[]
    if df.empty: return rows
    cols=list(df.columns)
    col_block=find_col(cols,['block','janpad','जनपद'])
    col_gp=find_col(cols,['panchayat','gram panchayat','ग्राम पंचायत','gp'])
    col_code=find_col(cols,['work code','कार्य कोड','code'])
    col_name=find_col(cols,['work name','कार्य नाम','name'])
    col_type=find_col(cols,['work type','category','श्रेणी','type'])
    col_status=find_col(cols,['status','स्थिति'])
    col_year=find_col(cols,['FIN. YEAR','fin. year','fin year','financial year','वित्तीय वर्ष','year','वर्ष'])
    col_wage_sanc=find_col(cols,['wage sanctioned','wage sanction','wage sanc','मजदूरी स्वीकृत'])
    col_mat_sanc=find_col(cols,['material sanctioned','material sanction','material sanc','सामग्री स्वीकृत'])
    col_sanc=find_col(cols,['sanction','स्वीकृत','sanctioned','estimated'])
    col_wage_book=find_col(cols,['wage booked','wage booking','wage exp','मजदूरी व्यय'])
    col_mat_book=find_col(cols,['material booked','material booking','material exp','सामग्री व्यय'])
    col_book=find_col(cols,['booked','expenditure','व्यय','खर्च','exp'])
    col_pct=find_col(cols,['% booked','booked %','exp %','%', 'percent','प्रतिशत'])
    for _,r in df.iterrows():
        text = ' '.join(str(v) for v in r.values)
        if not text.strip(): continue
        # Skip obvious header-like rows
        if 'work code' in text.lower() and 'panchayat' in text.lower(): continue
        status = str(r.get(col_status,'')) if col_status else ''
        w = {
            'block': str(r.get(col_block, block)).strip() if col_block else block,
            'panchayat': str(r.get(col_gp,'')).strip() if col_gp else '',
            'workCode': str(r.get(col_code,'')).strip() if col_code else '',
            'workName': str(r.get(col_name,'')).strip() if col_name else text[:140],
            'workType': str(r.get(col_type,'Uncategorised')).strip() if col_type else 'Uncategorised',
            'status': status,
            'finYear': (str(r.get(col_year,'')).strip() if col_year else '') or extract_fin_year_from_text(text),
            'sanctionAmount': ((num(r.get(col_wage_sanc,0)) if col_wage_sanc else 0) + (num(r.get(col_mat_sanc,0)) if col_mat_sanc else 0)) or (num(r.get(col_sanc,0)) if col_sanc else 0),
            'bookedAmount': ((num(r.get(col_wage_book,0)) if col_wage_book else 0) + (num(r.get(col_mat_book,0)) if col_mat_book else 0)) or (num(r.get(col_book,0)) if col_book else 0),
            'expPercent': num(r.get(col_pct,0)) if col_pct else 0,
            'needsVerification': bool(re.search(r'Needs\s*Verification|Verification\s*Needed|सत्यापन', text, re.I)),
            'rawText': re.sub(r'\s+', ' ', text).strip()[:600]
        }
        # If booked % not explicit, calculate where possible
        if not w['expPercent'] and w['sanctionAmount']:
            w['expPercent'] = round((w['bookedAmount']/w['sanctionAmount'])*100,2)
        rows.append(w)
    return rows

def load_eng_map(path):
    mapping={}
    if not os.path.exists(path): return mapping
    try:
        df=pd.read_excel(path, header=None)
    except Exception as e:
        print('engname read failed', e, file=sys.stderr); return mapping
    # detect header row containing जनपद and ग्राम पंचायत
    header_idx=0
    for i in range(min(20,len(df))):
        row=' '.join(str(x) for x in df.iloc[i].tolist())
        if ('जनपद' in row or 'JANPAD' in row.upper()) and ('ग्राम' in row or 'PANCHAYAT' in row.upper()):
            header_idx=i; break
    data=df.iloc[header_idx+1:].copy()
    # Known structure: क्रमांक, जनपद, क्लस्टर, ग्राम पंचायत, उपयंत्री
    for _,r in data.iterrows():
        vals=[str(x).strip() for x in r.tolist()]
        if len(vals)<5: continue
        block, gp, eng = vals[1], vals[3], vals[4]
        if not block or block.lower()=='nan' or not gp or gp.lower()=='nan' or not eng or eng.lower()=='nan': continue
        mapping[(norm(block), norm(gp))] = eng.strip()
    return mapping

def status_flags(status, text=''):
    # Statuses are exclusive on Work Monitor: Completed, Physically Completed, or Ongoing.
    # Do not let "Physically Completed" count as both Completed and Physical.
    s=norm(str(status))
    physical = any(k in s for k in ['PHYSICAL','PHYSICALLY','PHYCS','भौतिक'])
    complete = (not physical) and any(k in s for k in ['COMPLETE','COMPLETED','पूर्ण'])
    ongoing = any(k in s for k in ['ONGOING','प्रगतिरत','IN PROGRESS','चालू']) or (not complete and not physical)
    return complete, physical, ongoing

def grade(score):
    if score >= 8: return 'A'
    if score >= 6: return 'B'
    if score >= 4: return 'C'
    return 'D'

def grade_text(g):
    return {'A':'अच्छा Performance','B':'Progressing','C':'Progress Needed','D':'Critical / Poor Performance'}.get(g,'')

def calc_category_score(items):
    started=len(items)
    if not started:
        return {'score':0,'partA':0,'partB':0,'avgExpPct':0,'works':0,'completedPhy':0,'completed':0,'physical':0,'ongoing':0,'sanction':0,'booked':0,'ongoingSanction':0,'ongoingBooked':0}
    comp=phy=ongo=0
    total_sanc=total_book=ongo_sanc=ongo_book=0
    pct_sum=0
    for w in items:
        c,p,o=status_flags(w.get('status',''), w.get('rawText',''))
        comp += int(c)
        phy += int(p)
        ongo += int(o)
        total_sanc += w.get('sanctionAmount',0) or 0
        total_book += w.get('bookedAmount',0) or 0
        if o:
            ongo_sanc += w.get('sanctionAmount',0) or 0
            ongo_book += w.get('bookedAmount',0) or 0
        pct_sum += w.get('expPercent',0) or 0
    comp_phy = comp + phy
    partA = min(5, (comp_phy/started)*5) if started else 0
    if not started:
        partB = 0
    elif ongo:
        partB = min(5, (ongo_book/ongo_sanc)*5) if ongo_sanc else 0
    elif comp_phy >= started:
        partB = 5
    else:
        partB = 0
    return {'score':round(partA+partB,2), 'partA':round(partA,2), 'partB':round(partB,2), 'avgExpPct':round(pct_sum/started,2), 'works':started, 'completedPhy':comp_phy, 'completed':comp, 'physical':phy, 'ongoing':ongo, 'sanction':round(total_sanc,2), 'booked':round(total_book,2), 'ongoingSanction':round(ongo_sanc,2), 'ongoingBooked':round(ongo_book,2)}

def calc_engineers(works):
    groups={}
    for w in works:
        eng=w.get('engineer') or 'Unmapped'
        groups.setdefault(eng,[]).append(w)
    out=[]
    for eng,items in groups.items():
        bycat={}
        for w in items: bycat.setdefault(w.get('workType') or 'Uncategorised',[]).append(w)
        total_sanc=sum((w.get('sanctionAmount',0) or 0) for w in items)
        cats=[]; weighted=0
        for cat,ci in bycat.items():
            cs=calc_category_score(ci)
            weight=(cs['sanction']/total_sanc) if total_sanc else (cs['works']/len(items))
            weighted += weight*cs['score']
            cs.update({'category':cat, 'weight':round(weight*100,2), 'grade':grade(cs['score'])})
            cats.append(cs)
        cats=sorted(cats, key=lambda x: x['score'], reverse=True)
        comp=phy=ongo=needs=0; pct_sum=0; booked=0
        for w in items:
            c,p,o=status_flags(w.get('status',''), w.get('rawText',''))
            comp+=int(c); phy+=int(p); ongo+=int(o)
            needs+=int(w.get('needsVerification',False))
            pct_sum += w.get('expPercent',0) or 0
            booked += w.get('bookedAmount',0) or 0
        sc=round(weighted,2)
        g=grade(sc)
        blocks=sorted(set(w.get('block','') for w in items if w.get('block')))
        out.append({'engineer':eng, 'janpad':', '.join(blocks), 'works':len(items), 'completed':comp, 'physicalCompleted':phy, 'ongoing':ongo, 'needsVerification':needs, 'score':sc, 'grade':g, 'gradeText':grade_text(g), 'avgBookedPct':round(pct_sum/len(items),2) if items else 0, 'sanction':round(total_sanc,2), 'booked':round(booked,2), 'categories':cats})
    out=sorted(out, key=lambda x:(-x['score'], x['needsVerification'], -x['works']))
    for i,x in enumerate(out,1): x['rank']=i
    return out

def calc_blocks(works):
    groups={}
    for w in works: groups.setdefault(w.get('block','Unknown'),[]).append(w)
    arr=[]
    for b,items in groups.items():
        comp=phy=ongo=needs=0; sanc=book=0
        bycat={}
        for w in items:
            c,p,o=status_flags(w.get('status',''), w.get('rawText',''))
            comp+=int(c); phy+=int(p); ongo+=int(o); needs+=int(w.get('needsVerification',False))
            sanc += w.get('sanctionAmount',0) or 0; book += w.get('bookedAmount',0) or 0
            bycat.setdefault(w.get('workType') or 'Uncategorised',[]).append(w)
        cats=[]; weighted=0
        for cat,ci in bycat.items():
            cs=calc_category_score(ci); weight=(cs['sanction']/sanc) if sanc else (cs['works']/len(items)); weighted+=weight*cs['score']; cs.update({'category':cat,'weight':round(weight*100,2)}); cats.append(cs)
        sc=round(weighted,2); g=grade(sc)
        arr.append({'block':b, 'works':len(items), 'completed':comp, 'physicalCompleted':phy, 'ongoing':ongo, 'needsVerification':needs, 'sanction':round(sanc,2), 'booked':round(book,2), 'avgBookedPct':round(book/sanc*100,2) if sanc else 0, 'score':sc, 'grade':g, 'gradeText':grade_text(g), 'categories':sorted(cats,key=lambda x:x['score'], reverse=True)})
    arr=sorted(arr,key=lambda x:-x['score'])
    for i,x in enumerate(arr,1): x['rank']=i
    return arr

def fetch_work_monitor():
    all_works=[]; urls={}
    for block in BLOCKS:
        params={'district':DISTRICT,'block':block,'panchayat':'','work_type':'','status':'','exp_pct':'','q':'','date':DATE}
        url=BASE+'/work-monitor.php?'+urlencode(params)
        urls[block]=url
        print('fetch work monitor', block)
        try:
            html=get_html(url)
            df=choose_work_table(read_tables(html))
            works=parse_work_rows(df, block)
            # Add link and force block name
            for w in works:
                w['block']=block
                w['district']=district_for_block(block)
                w['sourceUrl']=url
            print(' rows', len(works))
            all_works.extend(works)
        except Exception as e:
            print('failed block', block, e, file=sys.stderr)
    return all_works, urls

OFFICIAL_CATEGORY_COLUMNS = [
    ('Farm Pond', ['FARM POND']),
    ('Amrit Sarovar', ['AMRIT SAROVAR','AMRIT SAROWAR']),
    ('Dug Well Recharge', ['DUG WELL RECHARGE','DUG WELL']),
    ('Irrigation Infrastructure', ['IRRIGATION INFRASTRUCTURE','IRRIGATION']),
    ('Water Conservation & Recharge', ['WATER CONSERVATION','WATER CONSERVATION & RECHARGE']),
    ('Watershed Related Works', ['WATERSHED RELATED']),
    ('Repair & Maintenance (Water Structures)', ['REPAIR & MAINTENANCE','WATER STRUCTURES']),
    ('Gap Filling in Plantation', ['GAP FILLING']),
    ('Work Not Permissible in VB-GRAM-G', ['WORK NOT PERMISSIBLE','VB-GRAM'])
]

BLOCK_ALIASES = ['MAJHGAWAN','NAGOD','AMARPATAN','UNCHAHARA','RAMNAGAR','RAMPUR BAGHELAN','RAMPUR','SATNA','MAIHAR']

def extract_score_from_cell(value):
    """Official rankings cells are sometimes long explanatory text.
    Prefer 'Final score X / 10', otherwise a simple numeric cell."""
    s = re.sub(r'\s+', ' ', str(value or '')).strip()
    if not s or s.lower() == 'nan':
        return ''
    m = re.search(r'Final\s+score\s+(-?\d+(?:\.\d+)?)\s*/\s*10', s, re.I)
    if m:
        return round(float(m.group(1)), 2)
    m = re.search(r'(-?\d+(?:\.\d+)?)\s*/\s*10', s, re.I)
    if m:
        return round(float(m.group(1)), 2)
    nums = re.findall(r'-?\d+(?:\.\d+)?', s.replace(',', ''))
    if len(nums) == 1:
        return round(float(nums[0]), 2)
    # Many official category cells start with weight% then category score,
    # e.g. "36.0% 5.43 Category ... Final score 5.43 / 10".
    if len(nums) >= 2 and '%' in s[:20]:
        try:
            return round(float(nums[1]), 2)
        except Exception:
            pass
    return s

def pick_value_from_row(row, keywords):
    # First match by column name.
    for col, val in row.items():
        nc = norm(col)
        if all(k in nc for k in [norm(x) for x in keywords]):
            return val
    # Then match if any keyword is in column name.
    for col, val in row.items():
        nc = norm(col)
        if any(norm(x) in nc for x in keywords):
            return val
    return ''

def normalize_official_row(row, fallback_rank):
    vals = {str(k): str(v) for k, v in row.items()}
    rowtxt = re.sub(r'\s+', ' ', ' '.join(vals.values())).strip()
    block = ''
    for v in vals.values():
        nv = norm(v)
        for b in BLOCK_ALIASES:
            if b in nv:
                block = 'RAMPUR BAGHELAN' if b == 'RAMPUR' else b
                break
        if block:
            break
    if not block:
        return None

    rank_raw = pick_value_from_row(vals, ['rank']) or pick_value_from_row(vals, ['#'])
    rank = int(num(rank_raw)) if num(rank_raw) else fallback_rank

    total_raw = pick_value_from_row(vals, ['total']) or pick_value_from_row(vals, ['score'])
    total = num(total_raw)
    if not total:
        # Fallback: first decimal looking like score /10 after block text.
        candidates = [float(x) for x in re.findall(r'\b\d+(?:\.\d+)?\b', rowtxt) if 0 <= float(x) <= 10]
        total = candidates[0] if candidates else 0

    traj = pick_value_from_row(vals, ['trajectory']) or pick_value_from_row(vals, ['grade']) or ''
    traj = str(traj).strip()
    if len(traj) > 8:
        m = re.search(r'\b([ABCD])\b', traj.upper())
        traj = m.group(1) if m else traj[:8]

    out = {
        'Rank': rank,
        'Block': block,
        'Total': round(float(total), 2),
        'Trajectory': traj or grade(float(total)),
    }
    for label, keys in OFFICIAL_CATEGORY_COLUMNS:
        val = ''
        # Best match by column header.
        for col, cell in vals.items():
            nc = norm(col)
            if any(norm(k) in nc for k in keys):
                val = cell
                break
        # Fallback by text around category label is intentionally conservative.
        out[label] = extract_score_from_cell(val)
    out['Source'] = 'Official rankings.php'
    return out



def fetch_official_overview(date=None):
    """Fetch top overview KPI cards from the official JGSA overview page.

    Important guard: the overview page has nested/parent containers. If a parent
    container is parsed, every KPI can accidentally become the first number
    (5953) and money can become 0. This function therefore reads the number
    nearest to each KPI label and later validates values before summary override.
    """
    use_date = date or DATE
    url = BASE + '/?' + urlencode({'status':'all','district':DISTRICT,'block':'','worktype_id':'0','date':use_date})
    out = {}
    try:
        html = get_html(url)
        soup = BeautifulSoup(html, 'html.parser')

        def clean_text(x):
            return re.sub(r'\s+', ' ', str(x or '').strip())

        def parse_money_token(token, unit):
            try:
                val = float(str(token).replace(',', ''))
            except Exception:
                return 0
            u = (unit or '').lower()
            if u in ['cr', 'करोड़']:
                return round(val * 10000000, 2)
            if u in ['lakh', 'लाख']:
                return round(val * 100000, 2)
            return val

        def norm_label_text(s):
            return re.sub(r'[^A-Z0-9]+', ' ', str(s or '').upper()).strip()

        def extract_nearest_value(text, labels, money=False):
            """Return the number/money closest to the label in a compact card text."""
            if not text:
                return 0
            nt = norm_label_text(text)
            for lbl in labels:
                nl = norm_label_text(lbl)
                pos = nt.find(nl)
                if pos < 0:
                    continue
                # Mapping normalized position to raw position is approximate; use whole compact text.
                if money:
                    money_matches = list(re.finditer(r'₹\s*(-?\d[\d,]*(?:\.\d+)?)\s*(Cr|CR|करोड़|लाख|Lakh|LAKH)?', text))
                    if not money_matches:
                        money_matches = list(re.finditer(r'(-?\d[\d,]*(?:\.\d+)?)\s*(Cr|CR|करोड़|लाख|Lakh|LAKH)', text))
                    if money_matches:
                        # On official cards, the main amount appears first inside the card.
                        m = money_matches[0]
                        return parse_money_token(m.group(1), m.group(2) if m.lastindex and m.lastindex >= 2 else '')
                    return 0
                nums = list(re.finditer(r'-?\d[\d,]*(?:\.\d+)?', text))
                if not nums:
                    return 0
                # For official cards, the main KPI number usually appears before label.
                # Pick the first large visible KPI number in the compact element.
                vals=[]
                for m in nums:
                    try:
                        v=float(m.group(0).replace(',',''))
                        vals.append((m.start(), int(v) if v.is_integer() else v))
                    except Exception:
                        pass
                # Prefer a number before the label in raw text if possible.
                raw_label_pos = min([i for i in [text.upper().find(lbl.upper()) for lbl in labels] if i >= 0] or [len(text)])
                before=[v for s,v in vals if s < raw_label_pos]
                if before:
                    return before[-1]
                return vals[0][1] if vals else 0
            return 0

        def best_card_value(labels, money=False):
            labels_norm = [norm_label_text(x) for x in labels]
            candidates=[]
            for el in soup.find_all(['div','section','article','li','span']):
                txt = clean_text(el.get_text(' ', strip=True))
                if not txt:
                    continue
                nt = norm_label_text(txt)
                if not any(lbl in nt for lbl in labels_norm):
                    continue
                # Reject very large parent containers with many KPI labels.
                label_hits = sum(1 for k in ['TOTAL TARGET WORKS','TOTAL COMPLETED','ABHIYAN PROGRESS','TOTAL SANCTIONED','TOTAL BOOKED'] if norm_label_text(k) in nt)
                if len(txt) > 260 or label_hits > 2:
                    continue
                val = extract_nearest_value(txt, labels, money=money)
                if val:
                    candidates.append((len(txt), val, txt))
            if candidates:
                # For non-money KPI cards, nested small badges can contain numbers like
                # AFTER 19-MAR, Completed 856, Phys 2303 etc. The visible main KPI is
                # the largest compact numeric value in that card (e.g. 2963 for Abhiyan Progress).
                if not money:
                    numeric_vals = [v for _ln, v, _txt in candidates if isinstance(v, (int, float))]
                    if numeric_vals:
                        return max(numeric_vals)
                candidates.sort(key=lambda x: x[0])
                return candidates[0][1]

            # Conservative fallback on full page: find a number just before/after the label.
            full = clean_text(soup.get_text(' ', strip=True))
            for lbl in labels:
                m = re.search(r'((?:₹\s*)?-?\d[\d,]*(?:\.\d+)?\s*(?:Cr|CR|करोड़|लाख|Lakh|LAKH)?)\s*.{0,30}?' + re.escape(lbl), full, re.I)
                if m:
                    if money:
                        mm = re.search(r'₹?\s*(-?\d[\d,]*(?:\.\d+)?)\s*(Cr|CR|करोड़|लाख|Lakh|LAKH)?', m.group(1), re.I)
                        if mm: return parse_money_token(mm.group(1), mm.group(2) or '')
                    else:
                        return num(m.group(1))
                m = re.search(re.escape(lbl) + r'.{0,30}?((?:₹\s*)?-?\d[\d,]*(?:\.\d+)?\s*(?:Cr|CR|करोड़|लाख|Lakh|LAKH)?)', full, re.I)
                if m:
                    if money:
                        mm = re.search(r'₹?\s*(-?\d[\d,]*(?:\.\d+)?)\s*(Cr|CR|करोड़|लाख|Lakh|LAKH)?', m.group(1), re.I)
                        if mm: return parse_money_token(mm.group(1), mm.group(2) or '')
                    else:
                        return num(m.group(1))
            return 0

        out['totalWorks'] = best_card_value(['TOTAL TARGET WORKS','Total Target Works','कुल लक्ष्य'])
        out['totalCompleted'] = best_card_value(['TOTAL COMPLETED','Total Completed'])
        out['officialTotalCompleted'] = out.get('totalCompleted', 0)
        out['abhiyanProgress'] = best_card_value(['ABHIYAN PROGRESS','Abhiyan Progress'])
        sanction = best_card_value(['TOTAL SANCTIONED','Total Sanctioned','TOTAL SANCTION'], money=True)
        booked = best_card_value(['TOTAL BOOKED','Total Booked'], money=True)
        if sanction:
            out['sanction'] = sanction
        if booked:
            out['booked'] = booked
        if sanction and booked:
            out['bookedPct'] = round((booked / sanction) * 100, 2)
        out = {k:v for k,v in out.items() if v not in [0, 0.0, '', None]}
        print('official overview', out)
    except Exception as e:
        print('official overview failed', e, file=sys.stderr)
    return out, url



def parse_official_ranking_text_rows(html):
    """Robust official rankings.php parser.

    Uses each table row's full text and extracts the official "Final score X / 10"
    inside each category cell. This avoids both failure modes we saw:
    1) nested count columns shifting into category score columns, and
    2) stale old rows being preserved when the table layout changes.

    This is date-agnostic: whatever DATE is requested from rankings.php is parsed fresh.
    """
    soup = BeautifulSoup(html, 'html.parser')
    rows = []
    seen = set()

    category_patterns = [
        ('Farm Pond', r'Farm\s+Pond'),
        ('Amrit Sarovar', r'Amrit\s+Sarow?ar'),
        ('Dug Well Recharge', r'Dug\s+Well\s+Recharge'),
        ('Irrigation Infrastructure', r'Irrigation\s+infrastructure'),
        ('Water Conservation & Recharge', r'Water\s+conservation\s*&\s*recharge'),
        ('Watershed Related Works', r'Watershed\s+Related\s+Works'),
        ('Repair & Maintenance (Water Structures)', r'Repair\s*&\s*Maintenance\s*\(\s*Water\s+Structures\s*\)'),
        ('Gap Filling in Plantation', r'Gap\s+Filling\s+in\s+Plantation'),
        ('Work Not Permissible in VB-GRAM-G', r'Work\s+Not\s+Permissible\s+in\s+VB-GRAM-G'),
    ]

    for tr in soup.find_all('tr'):
        rowtxt = re.sub(r'\s+', ' ', tr.get_text(' ', strip=True)).strip()
        if not rowtxt or 'Final score' not in rowtxt:
            continue

        block = ''
        for b in BLOCK_ALIASES:
            if re.search(r'\b' + re.escape(b) + r'\b', rowtxt, re.I):
                block = 'RAMPUR BAGHELAN' if b == 'RAMPUR' else b
                break
        if not block or block in seen:
            continue

        # Require that this is a real ranking row, not a legend/formula block.
        if sum(1 for _, pat in category_patterns if re.search(r'Category\s+' + pat, rowtxt, re.I)) < 3:
            continue

        cells = [re.sub(r'\s+', ' ', td.get_text(' ', strip=True)).strip()
                 for td in tr.find_all('td', recursive=False)]

        # Rank: first small integer cell if present, otherwise row order.
        rank = len(rows) + 1
        if cells:
            r0 = num(cells[0])
            if r0 and 1 <= r0 <= 99:
                rank = int(r0)

        # Total score: prefer the numeric cell immediately after the block cell.
        total = 0.0
        for i, c in enumerate(cells):
            if re.search(r'\b' + re.escape(block) + r'\b', norm(c)):
                if i + 1 < len(cells):
                    t = num(cells[i + 1])
                    if 0 <= t <= 10:
                        total = t
                break
        if not total:
            # Fallback: first 0-10 decimal after block name in row text.
            m = re.search(re.escape(block) + r'\s*(?:↗|↘|→)?\s*(\d+(?:\.\d+)?)', rowtxt, re.I)
            if m:
                total = float(m.group(1))
        if not total or total > 10:
            continue

        # Trajectory grade: direct cell after total, or first A/B/C/D after total.
        traj = 'D'
        for i, c in enumerate(cells):
            if re.search(r'\b' + re.escape(block) + r'\b', norm(c)):
                if i + 2 < len(cells):
                    m = re.search(r'\b([ABCD])\b', str(cells[i + 2]).upper())
                    if m:
                        traj = m.group(1)
                break

        out = {
            'Rank': rank,
            'Block': block,
            'Total': round(float(total), 2),
            'Trajectory': traj,
        }

        for label, pat in category_patterns:
            # Match from "Category <name>" until its own final score.
            m = re.search(
                r'Category\s+' + pat + r'.{0,1200}?Final\s+score\s+(-?\d+(?:\.\d+)?)\s*/\s*10',
                rowtxt, re.I
            )
            if not m:
                # Some cells start with "weight% score Category <name> ..." and still have final score later.
                m = re.search(
                    pat + r'.{0,1200}?Final\s+score\s+(-?\d+(?:\.\d+)?)\s*/\s*10',
                    rowtxt, re.I
                )
            out[label] = round(float(m.group(1)), 2) if m else ''

        out['Source'] = 'Official rankings.php text-row parser'
        rows.append(out)
        seen.add(block)

    if rows:
        rows = sorted(rows, key=lambda r: int(r.get('Rank') or 999))
    return rows

def parse_official_ranking_bs4(html):
    """Parse rankings.php official table using ONLY top-level row cells.
    This avoids nested count/detail values being shifted into score columns.
    """
    soup = BeautifulSoup(html, 'html.parser')
    best_headers = []
    best_rows = []
    for tbl in soup.find_all('table'):
        headers = []
        body = []
        for tr in tbl.find_all('tr'):
            # IMPORTANT: recursive=False prevents nested mini-tables/spans from becoming extra columns.
            ths = tr.find_all('th', recursive=False)
            tds = tr.find_all('td', recursive=False)
            if ths and len(ths) >= 6:
                headers = [re.sub(r'\s+', ' ', c.get_text(' ', strip=True)).strip() for c in ths]
                continue
            if tds and len(tds) >= 6:
                vals = [re.sub(r'\s+', ' ', c.get_text(' ', strip=True)).strip() for c in tds]
                body.append(vals)
        blob = norm(' '.join(headers) + ' ' + ' '.join(' '.join(r) for r in body[:10]))
        if not any(b in blob for b in BLOCK_ALIASES):
            continue
        if not any(norm(k) in blob for _, keys in OFFICIAL_CATEGORY_COLUMNS for k in keys):
            continue
        if len(body) > len(best_rows):
            best_headers, best_rows = headers, body
    rows=[]
    fallback_rank=1
    for vals in best_rows:
        if best_headers and len(best_headers) >= 6:
            hdr = best_headers + [f'col_{i}' for i in range(len(best_headers), len(vals))]
            row = {hdr[i]: vals[i] if i < len(vals) else '' for i in range(min(len(hdr), len(vals)))}
            nr = normalize_official_row(row, fallback_rank)
        else:
            nr = None
        if nr:
            rows.append(nr)
            fallback_rank += 1
    return rows

def official_rows_valid(rows):
    if not rows or len(rows) < 5:
        return False
    for r in rows:
        if not r.get('Block') or not r.get('Total'):
            return False
        for label, _ in OFFICIAL_CATEGORY_COLUMNS:
            v = r.get(label, '')
            if v == '':
                continue
            try:
                if float(v) < 0 or float(v) > 10:
                    return False
            except Exception:
                return False
    return True

LAST_GOOD_OFFICIAL_BLOCK_RANKING_ROWS = [
  {
    "Rank": 1,
    "Block": "MAJHGAWAN",
    "Total": 5.86,
    "Trajectory": "D",
    "Farm Pond": 5.47,
    "Amrit Sarovar": 1.92,
    "Dug Well Recharge": 6.63,
    "Irrigation Infrastructure": 10.0,
    "Water Conservation & Recharge": 7.89,
    "Watershed Related Works": 7.76,
    "Repair & Maintenance (Water Structures)": 10.0,
    "Gap Filling in Plantation": 1.78,
    "Work Not Permissible in VB-GRAM-G": 2.17,
    "Source": "Official rankings.php text parser"
  },
  {
    "Rank": 2,
    "Block": "SATNA",
    "Total": 5.46,
    "Trajectory": "D",
    "Farm Pond": 4.94,
    "Amrit Sarovar": 1.51,
    "Dug Well Recharge": 6.33,
    "Irrigation Infrastructure": 7.05,
    "Water Conservation & Recharge": 7.39,
    "Watershed Related Works": 7.35,
    "Repair & Maintenance (Water Structures)": 5.64,
    "Gap Filling in Plantation": 0.3,
    "Work Not Permissible in VB-GRAM-G": 2.89,
    "Source": "Official rankings.php text parser"
  },
  {
    "Rank": 3,
    "Block": "NAGOD",
    "Total": 5.35,
    "Trajectory": "D",
    "Farm Pond": 4.35,
    "Amrit Sarovar": 1.0,
    "Dug Well Recharge": 6.51,
    "Irrigation Infrastructure": 10.0,
    "Water Conservation & Recharge": 6.82,
    "Watershed Related Works": 10.0,
    "Repair & Maintenance (Water Structures)": 10.0,
    "Gap Filling in Plantation": 1.86,
    "Work Not Permissible in VB-GRAM-G": 0.86,
    "Source": "Official rankings.php text parser"
  },
  {
    "Rank": 4,
    "Block": "RAMPUR BAGHELAN",
    "Total": 5.03,
    "Trajectory": "D",
    "Farm Pond": 4.24,
    "Amrit Sarovar": 0.47,
    "Dug Well Recharge": 6.92,
    "Irrigation Infrastructure": 6.03,
    "Water Conservation & Recharge": 8.04,
    "Watershed Related Works": 6.42,
    "Repair & Maintenance (Water Structures)": 7.44,
    "Gap Filling in Plantation": 1.13,
    "Work Not Permissible in VB-GRAM-G": 1.81,
    "Source": "Official rankings.php text parser"
  },
  {
    "Rank": 5,
    "Block": "MAIHAR",
    "Total": 4.94,
    "Trajectory": "D",
    "Farm Pond": 4.56,
    "Amrit Sarovar": 0.0,
    "Dug Well Recharge": 7.17,
    "Irrigation Infrastructure": 8.33,
    "Water Conservation & Recharge": 7.59,
    "Watershed Related Works": 6.94,
    "Repair & Maintenance (Water Structures)": 6.48,
    "Gap Filling in Plantation": 1.55,
    "Work Not Permissible in VB-GRAM-G": 2.17,
    "Source": "Official rankings.php text parser"
  },
  {
    "Rank": 6,
    "Block": "AMARPATAN",
    "Total": 4.87,
    "Trajectory": "D",
    "Farm Pond": 2.85,
    "Amrit Sarovar": "",
    "Dug Well Recharge": 4.43,
    "Irrigation Infrastructure": "",
    "Water Conservation & Recharge": 10.0,
    "Watershed Related Works": "",
    "Repair & Maintenance (Water Structures)": 10.0,
    "Gap Filling in Plantation": 0.9,
    "Work Not Permissible in VB-GRAM-G": 1.88,
    "Source": "Official rankings.php text parser"
  },
  {
    "Rank": 7,
    "Block": "RAMNAGAR",
    "Total": 4.83,
    "Trajectory": "D",
    "Farm Pond": 4.47,
    "Amrit Sarovar": "",
    "Dug Well Recharge": 10.0,
    "Irrigation Infrastructure": 10.0,
    "Water Conservation & Recharge": 10.0,
    "Watershed Related Works": 10.0,
    "Repair & Maintenance (Water Structures)": 10.0,
    "Gap Filling in Plantation": 10.0,
    "Work Not Permissible in VB-GRAM-G": 1.48,
    "Source": "Official rankings.php text parser"
  },
  {
    "Rank": 8,
    "Block": "UNCHAHARA",
    "Total": 4.59,
    "Trajectory": "D",
    "Farm Pond": 2.92,
    "Amrit Sarovar": "",
    "Dug Well Recharge": 6.14,
    "Irrigation Infrastructure": "",
    "Water Conservation & Recharge": 7.07,
    "Watershed Related Works": 8.71,
    "Repair & Maintenance (Water Structures)": 7.39,
    "Gap Filling in Plantation": 1.51,
    "Work Not Permissible in VB-GRAM-G": 1.73,
    "Source": "Official rankings.php text parser"
  }
]

def load_existing_official_rows(expected_date=None, allow_any_date=True):
    """Return last valid official rows from current jgsa_live_data.js.

    Priority:
    1) same-date rows when available,
    2) any-date rows from the existing deployed data file,
    3) bundled last-good official rows.

    This prevents the dashboard from being overwritten with an empty Official
    Block Ranking when the portal ranking page is temporarily unavailable.
    """
    same_date_rows = []
    any_date_rows = []
    try:
        if os.path.exists(OUT):
            txt = open(OUT, encoding='utf-8').read()
            m = re.search(r'window\.JGSA_LIVE_DATA\s*=\s*(\{.*\})\s*;?\s*$', txt, re.S)
            if m:
                old = json.loads(m.group(1))
                rows = old.get('officialBlockRankingRows') or old.get('officialBlockRanking') or []
                if official_rows_valid(rows):
                    any_date_rows = rows
                    if expected_date and str(old.get('date') or '') == str(expected_date):
                        same_date_rows = rows
    except Exception as e:
        print('existing official rows fallback failed', e, file=sys.stderr)

    if same_date_rows:
        return same_date_rows
    if allow_any_date and any_date_rows:
        print('official ranking fallback: using previous valid deployed official rows')
        return any_date_rows
    if official_rows_valid(LAST_GOOD_OFFICIAL_BLOCK_RANKING_ROWS):
        print('official ranking fallback: using bundled last-good official rows')
        return LAST_GOOD_OFFICIAL_BLOCK_RANKING_ROWS
    return []


def fetch_official_ranking(date=None):
    """Fetch official JGSA block ranking from rankings.php and normalize it.

    Critical rule: fresh valid official rows must always replace old rows.
    Fallback to existing rows is allowed only when the fresh parser returns no
    valid table for the same date. This prevents stale 08-06 values from being
    preserved on 09-06 while still protecting the dashboard from a transient
    portal/parse failure.
    """
    use_date = date or DATE
    url = BASE + '/rankings.php?' + urlencode({'level':'block','date':use_date,'district':DISTRICT})
    rows = []
    try:
        html = get_html(url)

        # 1) Preferred parser: full row text + "Final score X / 10".
        # This is the most stable representation of the official scorecard.
        text_rows = parse_official_ranking_text_rows(html)
        if official_rows_valid(text_rows):
            rows = text_rows
            print('official ranking text rows', len(rows), 'date', use_date)
        else:
            print('official ranking text parser invalid/empty rows', len(text_rows), 'date', use_date)

        # 2) Backup parser: BeautifulSoup with direct cells only.
        if not official_rows_valid(rows):
            bs4_rows = parse_official_ranking_bs4(html)
            if official_rows_valid(bs4_rows):
                rows = bs4_rows
                print('official ranking bs4 rows', len(rows), 'date', use_date)
            else:
                print('official ranking bs4 invalid/empty rows', len(bs4_rows), 'date', use_date)

        # 3) Last parser: pandas/manual tables. Used only if both text and BS4 failed.
        if not official_rows_valid(rows):
            tables = read_tables(html)
            candidates = []
            for df in tables:
                df = clean_df(df)
                text_blob = ' '.join([str(c) for c in df.columns]) + ' ' + (' '.join(df.astype(str).values.flatten()[:300]))
                if re.search(r'MAJHGAWAN|NAGOD|AMARPATAN|UNCHAHARA|RAMNAGAR|RAMPUR|SATNA|MAIHAR', text_blob, re.I):
                    candidates.append(df)
            if candidates:
                def cscore(d):
                    blob = norm(' '.join(map(str,d.columns))+' '+(' '.join(d.astype(str).values.flatten()[:200])))
                    score = len(d)*100 + len(d.columns)
                    for b in BLOCK_ALIASES:
                        if b in blob: score += 200
                    for label, keys in OFFICIAL_CATEGORY_COLUMNS:
                        if any(norm(k) in blob for k in keys): score += 50
                    return score
                df = max(candidates, key=cscore)
                df.columns = [re.sub(r'\s+',' ',str(c)).strip() for c in df.columns]
                fallback_rank = 1
                p_rows = []
                for _, r in df.iterrows():
                    rowtxt = ' '.join(str(v) for v in r.values)
                    if not re.search(r'MAJHGAWAN|NAGOD|AMARPATAN|UNCHAHARA|RAMNAGAR|RAMPUR|SATNA|MAIHAR', rowtxt, re.I):
                        continue
                    nr = normalize_official_row({str(k):str(v) for k,v in r.items()}, fallback_rank)
                    if nr:
                        p_rows.append(nr)
                        fallback_rank += 1
                if official_rows_valid(p_rows):
                    rows = p_rows
                    print('official ranking pandas rows', len(rows), 'date', use_date)
                else:
                    print('official ranking pandas invalid/empty rows', len(p_rows), 'date', use_date)

        if official_rows_valid(rows):
            rows = sorted(rows, key=lambda r: (int(r.get('Rank') or 999), -float(r.get('Total') or 0)))
        else:
            rows = []
    except Exception as e:
        print('official ranking failed', e, file=sys.stderr)
    return rows, url

def validate_before_write(data):
    total = len(data.get('works', []))
    if not official_rows_valid(data.get('officialBlockRankingRows') or data.get('officialBlockRanking') or []):
        raise RuntimeError('Official Block Ranking rows are empty/invalid; refusing to overwrite dashboard with blank ranking table.')
    if total < 4000:
        raise RuntimeError(f'Fetched only {total} works after Gap Filling FY filter; refusing to overwrite dashboard data. Check Work Monitor parsing/portal availability.')
    return True

def main():
    engmap=load_eng_map(ENG)
    works, work_urls=fetch_work_monitor()
    works_before_gap_filter = len(works)
    works, excluded_gap_filling_after_fy = filter_gap_filling_after_allowed_fy(works)
    print('gap filling filter: excluded FY 2022-2023 onward', len(excluded_gap_filling_after_fy), 'kept', len(works), 'from', works_before_gap_filter)
    # map engineer exact names, including अति/अति0 suffixes
    unmapped=0
    for w in works:
        key=(norm(w.get('block')), norm(w.get('panchayat')))
        eng=engmap.get(key)
        if not eng:
            unmapped+=1; eng='Unmapped'
        w['engineer']=eng
    engineerRanking=calc_engineers(works)
    internalBlock=calc_blocks(works)
    officialRows, rankingUrl=fetch_official_ranking(DATE)
    # Do NOT silently use old official ranking rows for today's/current DATE.
    # If fresh official ranking cannot be parsed, fail the Action so yesterday's
    # already-deployed data remains visible instead of showing stale scores with
    # a new timestamp.
    if not official_rows_valid(officialRows):
        raise RuntimeError(f'Fresh Official Block Ranking could not be parsed for {DATE}; refusing to overwrite with stale rows.')
    previousOfficialRows, previousRankingUrl=fetch_official_ranking(PREV_DATE)
    if not official_rows_valid(previousOfficialRows):
        previousOfficialRows = load_existing_official_rows(PREV_DATE, allow_any_date=True)
    officialOverview, overviewUrl=fetch_official_overview(DATE)
    total=len(works); needs=sum(1 for w in works if w.get('needsVerification'))
    comp=phy=ongo=0; sanc=book=0
    for w in works:
        c,p,o=status_flags(w.get('status',''), w.get('rawText',''))
        comp+=int(c); phy+=int(p); ongo+=int(o); sanc+=w.get('sanctionAmount',0) or 0; book+=w.get('bookedAmount',0) or 0
    summary={'totalWorks':total,
             'completed':comp,
             'completedOnly':comp,
             'physicalCompleted':phy,
             'totalCompleted':comp+phy,
             'officialTotalCompleted':comp+phy,
             'ongoing':ongo,
             'needsVerification':needs,
             'sanction':round(sanc,2),
             'booked':round(book,2),
             'bookedPct':round(book/sanc*100,2) if sanc else 0,
             'engineers':len(engineerRanking),
             'unmappedWorks':unmapped}
    # Official overview card values override only when they pass sanity checks.
    # This prevents a bad parser run from turning Completed/Progress into 5953 and money into ₹0.
    if officialOverview:
        ow_total = officialOverview.get('totalWorks')
        if ow_total and 5000 <= ow_total <= 7000:
            summary['totalWorks'] = ow_total
        ow_completed = officialOverview.get('totalCompleted')
        if ow_completed and 0 < ow_completed < summary['totalWorks']:
            summary['totalCompleted'] = ow_completed
            summary['officialTotalCompleted'] = ow_completed
        ow_progress = officialOverview.get('abhiyanProgress')
        # Guard against parsing the "AFTER 19-MAR" label as the progress value.
        # Valid Abhiyan Progress should be a large count, at least around physical completed works.
        if ow_progress and max(1000, phy) <= ow_progress < summary['totalWorks']:
            summary['abhiyanProgress'] = ow_progress
        ow_sanction = officialOverview.get('sanction')
        if ow_sanction and ow_sanction > 100000000:
            summary['sanction'] = ow_sanction
        ow_booked = officialOverview.get('booked')
        if ow_booked and ow_booked > 10000000:
            summary['booked'] = ow_booked
        if summary.get('sanction') and summary.get('booked'):
            summary['bookedPct'] = round((summary['booked']/summary['sanction'])*100,2)
    if not summary.get('abhiyanProgress'):
        summary['abhiyanProgress']=summary.get('physicalCompleted', phy)
    data={'generatedAt':datetime.datetime.utcnow().isoformat()+'Z','date':DATE,'district':DISTRICT,
          'sourceUrls':{'main':overviewUrl, 'officialBlockRanking':rankingUrl, 'weeklyCurrentOfficialBlockRanking':rankingUrl, 'weeklyPreviousOfficialBlockRanking':previousRankingUrl, 'workMonitorByBlock':work_urls},
          'summary':summary,
          'works':works,
          'engineerRanking':engineerRanking,
          'blockRankingInternal':internalBlock,
          'officialBlockRankingRows':officialRows,
          'officialBlockRanking':officialRows,
          'officialRankingRows':officialRows,
          'officialBlockTopCards':[{'Rank':r.get('Rank'), 'Block':r.get('Block'), 'Total':r.get('Total'), 'Completed':'', 'Started':''} for r in officialRows[:3]],
          'weeklyPreviousDate':PREV_DATE,
          'weeklyCurrentDate':DATE,
          'weeklyPreviousOfficialBlockRows':previousOfficialRows,
          'weeklyPreviousOfficialBlockRanking':previousOfficialRows,
          'weeklyPreviousRankingRows':previousOfficialRows,
          'gradeLegend':{'A':'अच्छा Performance','B':'Progressing','C':'Progress Needed','D':'Critical / Poor Performance'},
          'notes':['Work data is fetched block-wise to avoid the 2000 row All-Janpad limit.','Gap Filling in Plantation works after FY 2021-2022 are excluded as per dashboard rule.','Engineer mapping comes only from engname.xlsx. JGSA work values come from live JGSA pages.']}
    validate_before_write(data)
    js='window.JGSA_LIVE_DATA = '+json.dumps(data, ensure_ascii=False, indent=2)+';\n'
    with open(OUT,'w',encoding='utf-8') as f: f.write(js)
    print('wrote', OUT, 'works', total, 'needs', needs, 'engineers', len(engineerRanking), 'unmapped', unmapped)

if __name__=='__main__': main()
