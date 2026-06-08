import os, re, json, math, sys, datetime
from urllib.parse import urlencode
import requests
import pandas as pd
from bs4 import BeautifulSoup

BASE = 'https://jgsa.nregsmp.org'
BLOCKS = ['AMARPATAN','MAIHAR','MAJHGAWAN','NAGOD','RAMNAGAR','RAMPUR BAGHELAN','SATNA','UNCHAHARA']
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
    if not started: return {'score':0,'partA':0,'partB':0,'avgExpPct':0,'works':0,'completedPhy':0,'sanction':0,'booked':0}
    comp_phy=0; sanc=0; book=0; pct_sum=0
    for w in items:
        c,p,o=status_flags(w.get('status',''), w.get('rawText',''))
        if c or p: comp_phy += 1
        sanc += w.get('sanctionAmount',0) or 0
        book += w.get('bookedAmount',0) or 0
        pct_sum += w.get('expPercent',0) or 0
    partA = min(5, (comp_phy/started)*5) if started else 0
    partB = min(5, (book/sanc)*5) if sanc else min(5, (pct_sum/started)/100*5)
    return {'score':round(partA+partB,2), 'partA':round(partA,2), 'partB':round(partB,2), 'avgExpPct':round(pct_sum/started,2), 'works':started, 'completedPhy':comp_phy, 'sanction':round(sanc,2), 'booked':round(book,2)}

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

    Keeps official overview values separate from Work Monitor row-level data:
    - totalCompleted = Completed + Physically Completed from official overview
    - abhiyanProgress = official Abhiyan Progress card
    """
    use_date = date or DATE
    url = BASE + '/?' + urlencode({'status':'all','district':DISTRICT,'block':'','worktype_id':'0','date':use_date})
    out = {}
    try:
        html = get_html(url)
        soup = BeautifulSoup(html, 'html.parser')

        def clean_text(x):
            return re.sub(r'\s+', ' ', str(x or '').strip())

        def parse_money_or_number(text):
            t = clean_text(text)
            nums = re.findall(r'₹?\s*(-?\d[\d,]*(?:\.\d+)?)\s*(Cr|CR|करोड़|लाख|Lakh|LAKH)?', t)
            if not nums:
                return 0
            n, unit = nums[0]
            val = float(n.replace(',', ''))
            unit = (unit or '').lower()
            if unit in ['cr', 'करोड़']:
                return round(val * 10000000, 2)
            if unit in ['lakh', 'लाख']:
                return round(val * 100000, 2)
            return int(val) if float(val).is_integer() else val

        def first_card_value(labels, money=False):
            labels_u = [norm(x) for x in labels]
            best = None
            # Prefer compact card-like elements so the whole page text is not parsed.
            for el in soup.find_all(['div','section','article','li']):
                txt = clean_text(el.get_text(' ', strip=True))
                if not txt or len(txt) > 450:
                    continue
                nt = norm(txt)
                if any(lbl in nt for lbl in labels_u):
                    best = txt
                    break
            if not best:
                full = clean_text(soup.get_text(' ', strip=True))
                for lbl in labels:
                    m = re.search(re.escape(lbl) + r'.{0,120}', full, re.I)
                    if m:
                        best = m.group(0)
                        break
            if not best:
                return 0
            if money:
                return parse_money_or_number(best)
            m = re.search(r'-?\d[\d,]*(?:\.\d+)?', best)
            if not m:
                return 0
            val = float(m.group(0).replace(',', ''))
            return int(val) if val.is_integer() else val

        out['totalWorks'] = first_card_value(['TOTAL TARGET WORKS','Total Target Works','कुल लक्ष्य'])
        out['totalCompleted'] = first_card_value(['TOTAL COMPLETED','Total Completed'])
        out['officialTotalCompleted'] = out.get('totalCompleted', 0)
        out['abhiyanProgress'] = first_card_value(['ABHIYAN PROGRESS','Abhiyan Progress'])
        sanction = first_card_value(['TOTAL SANCTIONED','Total Sanctioned','TOTAL SANCTION'], money=True)
        booked = first_card_value(['TOTAL BOOKED','Total Booked'], money=True)
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


def fetch_official_ranking(date=None):
    """Fetch official JGSA block ranking from rankings.php and normalize it for the dashboard.
    This source must remain separate from Work Monitor internal calculations.
    """
    use_date = date or DATE
    url=BASE+'/rankings.php?'+urlencode({'level':'block','date':use_date,'district':DISTRICT})
    rows=[]
    try:
        html=get_html(url)
        tables=read_tables(html)
        candidates=[]
        for df in tables:
            df=clean_df(df)
            text_blob=' '.join([str(c) for c in df.columns])+' '+(' '.join(df.astype(str).values.flatten()[:300]))
            if re.search(r'MAJHGAWAN|NAGOD|AMARPATAN|UNCHAHARA|RAMNAGAR|RAMPUR|SATNA|MAIHAR', text_blob, re.I):
                candidates.append(df)
        if candidates:
            # Prefer table with all janpads and category headers.
            def cscore(d):
                blob = norm(' '.join(map(str,d.columns))+' '+(' '.join(d.astype(str).values.flatten()[:200])))
                score = len(d)*100 + len(d.columns)
                for b in BLOCK_ALIASES:
                    if b in blob: score += 200
                for label, keys in OFFICIAL_CATEGORY_COLUMNS:
                    if any(norm(k) in blob for k in keys): score += 50
                return score
            df=max(candidates, key=cscore)
            df.columns=[re.sub(r'\s+',' ',str(c)).strip() for c in df.columns]
            fallback_rank=1
            for _,r in df.iterrows():
                rowtxt=' '.join(str(v) for v in r.values)
                if not re.search(r'MAJHGAWAN|NAGOD|AMARPATAN|UNCHAHARA|RAMNAGAR|RAMPUR|SATNA|MAIHAR', rowtxt, re.I):
                    continue
                nr = normalize_official_row({str(k):str(v) for k,v in r.items()}, fallback_rank)
                if nr:
                    rows.append(nr)
                    fallback_rank += 1
        rows = sorted(rows, key=lambda r: (int(r.get('Rank') or 999), -float(r.get('Total') or 0)))
    except Exception as e:
        print('official ranking failed', e, file=sys.stderr)
    return rows, url

def validate_before_write(data):
    total = len(data.get('works', []))
    if total < 5000:
        raise RuntimeError(f'Fetched only {total} works; refusing to overwrite dashboard data. Check Work Monitor parsing/portal availability.')
    return True

def main():
    engmap=load_eng_map(ENG)
    works, work_urls=fetch_work_monitor()
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
    previousOfficialRows, previousRankingUrl=fetch_official_ranking(PREV_DATE)
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
    # Official overview card values override only top KPI fields, not row-level Work Monitor counts.
    if officialOverview:
        for k in ['totalWorks','totalCompleted','officialTotalCompleted','abhiyanProgress','sanction','booked','bookedPct']:
            if officialOverview.get(k):
                summary[k]=officialOverview[k]
    if not summary.get('abhiyanProgress'):
        summary['abhiyanProgress']=summary.get('totalCompleted', comp+phy)
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
          'notes':['Work data is fetched block-wise to avoid the 2000 row All-Janpad limit.','Engineer mapping comes only from engname.xlsx. JGSA work values come from live JGSA pages.']}
    validate_before_write(data)
    js='window.JGSA_LIVE_DATA = '+json.dumps(data, ensure_ascii=False, indent=2)+';\n'
    with open(OUT,'w',encoding='utf-8') as f: f.write(js)
    print('wrote', OUT, 'works', total, 'needs', needs, 'engineers', len(engineerRanking), 'unmapped', unmapped)

if __name__=='__main__': main()
