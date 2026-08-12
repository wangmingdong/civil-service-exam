# -*- coding: utf-8 -*-
# 全量校验上岸分口径 v4：健壮解析（整表读入内存 + 跨省行表头合并 + 代码列12位兜底）
# 校验口径逻辑与 v2 一致（已验证0偏差），仅重写了 Excel 抽取层。
import json, re, openpyxl
GWY = r"E:/workspace/civil-service-exam/excel/名睿整理—2026年陕西省考体检名单汇总表.xlsx"
SY  = r"E:/workspace/civil-service-exam/excel/26上事业单位体检名单.xlsx"
DATA = r"E:/workspace/civil-service-exam/web/data.js"

def clean_num(s):
    if s is None: return None
    if isinstance(s,(int,float)): return float(s)
    t = str(s).replace('\u2002','').replace('\u00a0','').replace(' ','').replace('\n','').replace('\t','')
    if t in ('','/','缺考','—','-','无'): return None
    m = re.search(r'-?\d+(?:\.\d+)?', t)
    return float(m.group(0)) if m else None

def norm(s):
    return re.sub(r'\s+','', str(s)) if s is not None else ''

def find_header_rows(ws):
    """返回 (hr_list_idx, composite_header_list)。composite 把下一行同列文本拼上，兼容跨省行表头。"""
    rows = list(ws.iter_rows(values_only=True))
    n = min(len(rows), 40)
    for ri in range(n):
        cur = " ".join(norm(c) for c in rows[ri])
        nxt = " ".join(norm(c) for c in rows[ri+1]) if ri+1 < n else ""
        has_name = ("姓名" in cur or "准考证" in cur)
        has_tijian = ("体检" in cur or "体检" in nxt)
        if has_name and has_tijian:
            # composite: 当前行 + 下一行 按列拼接
            comp = []
            nxt_row = rows[ri+1] if ri+1 < n else []
            for ci in range(max(len(rows[ri]), len(nxt_row))):
                a = norm(rows[ri][ci]) if ci < len(rows[ri]) else ''
                b = norm(nxt_row[ci]) if ci < len(nxt_row) else ''
                comp.append(a + b)
            return ri, comp, rows
    return None, None, rows

def map_cols(header):
    """返回 code,name,written,comp_cands,tijian。
    comp_cands 是综合分候选列列表 [(优先级,列号)]，真正选哪一列由 choose_comp 按数据值(应≤100)决定，
    以排除"笔试总成绩"(300量纲)被跨行表头误合并进"综合成绩"字样的陷阱。"""
    code=name=written=tijian=None
    comp_cands=[]
    for ci,c in enumerate(header):
        s = norm(c)
        if not s: continue
        if code is None:
            if ('代码' in s or '报考职位' in s):
                code = ci
            elif s in ('职位','岗位'):
                code = ci
        if name is None and '姓名' in s:
            name = ci
        if written is None and '笔试' in s and ('总' in s or '成绩' in s):
            written = ci
        if '综合成绩' in s:
            pri = 0 if ('笔试' not in s) else 3   # 含"笔试"的综合字样优先降低(多是笔试总成绩误合并)
            comp_cands.append((pri,ci))
        elif '考试总成绩' in s:
            comp_cands.append((1,ci))
        elif (s=='总成绩' or s.startswith('总成绩')) and '笔试' not in s[:4]:
            comp_cands.append((2,ci))
        if tijian is None and '体检' in s and ('是否' in s or '进入' in s or '资格' in s):
            tijian = ci
    comp_cands.sort(key=lambda x:(x[0],x[1]))
    return code,name,written,comp_cands,tijian

def choose_comp(comp_cands, rows, hr):
    """从候选里挑真正综合分列：真实综合分(折算后)必≤100，笔试总成绩为300量纲。
    若最佳候选的≤100占比<0.5，视为该 sheet 无可用综合分列，返回 None(跳过onBoard校验)。"""
    if not comp_cands:
        return None
    best=None; bestscore=-1; bestfrac=-1
    for pri,ci in comp_cands:
        vals=[clean_num(r[ci]) for r in rows[hr+1:hr+50] if ci<len(r) and r[ci] is not None]
        vals=[v for v in vals if v is not None]
        if not vals:
            score=-1; frac=0.0
        else:
            le100=sum(1 for v in vals if v<=100)
            frac=le100/len(vals)
            score=frac*10 - pri
        if score>bestscore:
            bestscore=score; best=ci; bestfrac=frac
    if bestfrac<0.5:    # 该列数值基本都是300量纲，不是综合分
        return None
    return best

def code_col_fallback(rows, hr, name_c, tij_c):
    """若名字没匹配到代码列，扫描各列找 12 位数字最多的列。"""
    best=None; bestcnt=-1
    width = max((len(r) for r in rows[hr+1:hr+60]), default=0)
    for ci in range(width):
        cnt=0
        for r in rows[hr+1:hr+60]:
            if ci < len(r) and r[ci] is not None and re.match(r'^\d{12}$', norm(r[ci])):
                cnt+=1
        if cnt>bestcnt:
            bestcnt=cnt; best=ci
    return best if bestcnt>0 else None

def is_yes(v):
    s = norm(v)
    if not s: return False
    return s in ('是','进入体检','√','是√') or ('是' in s and '不' not in s and s!='')

PAT = re.compile(r'(?<!\d)\d{9,13}(?!\d)')  # 9~13 位孤立数字串（公务员9位/事业编11~12位），排除13位准考证号等长数字干扰
def best_code_col(rows, hr, header):
    """扫描数据区，找含孤立数字串最多的列作为代码列；排除表头含'准考证'的列（准考证号）。"""
    best=None; bestcnt=-1
    data = rows[hr+1:hr+200]
    width = max((len(r) for r in data), default=0)
    for ci in range(width):
        hc = header[ci] if ci < len(header) else ''
        hh = hc.replace(' ','')
        if hh in ('准考证号','准考证') or hh.startswith('准考证号'): continue   # 跳过准考证号列（不误伤含"准考证"字样的岗位列）
        cnt = sum(1 for r in data if ci < len(r) and r[ci] is not None and PAT.search(norm(r[ci])))
        if cnt>bestcnt:
            bestcnt=cnt; best=ci
    return best if bestcnt>=2 else None

def parse(path, label, skip_sheets=(), code_len=12):
    wb = openpyxl.load_workbook(path, data_only=True)
    out = {}; skipped=[]
    for sh in wb.sheetnames:
        if sh in skip_sheets: continue
        ws = wb[sh]
        hr, header, rows = find_header_rows(ws)
        if hr is None:
            skipped.append((sh,'无表头')); continue
        cols = map_cols(header)
        name_c,w_c,comp_cands,tij_c = cols[1],cols[2],cols[3],cols[4]
        bcc = best_code_col(rows, hr, header)
        code_c = bcc if bcc is not None else cols[0]
        comp_c = choose_comp(comp_cands, rows, hr)
        if code_c is None or tij_c is None:
            skipped.append((sh,'列缺失 code=%s tij=%s'%(code_c,tij_c))); continue
        n=0
        for r in rows[hr+1:]:
            if code_c>=len(r) or r[code_c] is None: continue
            m = PAT.search(norm(r[code_c]))
            if not m: continue
            code = m.group(0)

            nm = norm(r[name_c]) if name_c is not None and name_c<len(r) and r[name_c] else ''
            w = clean_num(r[w_c]) if w_c is not None and w_c<len(r) else None
            c = clean_num(r[comp_c]) if comp_c is not None and comp_c<len(r) else None
            yes = is_yes(r[tij_c]) if tij_c<len(r) else False
            out.setdefault(code, {'yes':[], 'all':[]})
            rec = {'name':nm,'w':w,'c':c}
            out[code]['all'].append(rec)
            if yes: out[code]['yes'].append(rec)
            n+=1
        if n==0: skipped.append((sh,'无数据'))
    wb.close()
    print("%s 解析: codes=%d  跳过sheets=%d" % (label, len(out), len(skipped)))
    for s in skipped: print("   跳过", s)
    return out

gwy_map = parse(GWY,'GWY', code_len=9)
sy_map  = parse(SY,'SY', skip_sheets=('Sheet9',), code_len=12)

txt = open(DATA, encoding='utf-8').read()
i = txt.index('[', txt.index('window.JOBS=')); depth=0
for j in range(i,len(txt)):
    if txt[j]=='[': depth+=1
    elif txt[j]==']':
        depth-=1
        if depth==0: break
J = json.loads(txt[i:j+1])

MISSING=[]
def verify(typ, year, emap):
    recs = [x for x in J if x.get('type')==typ and x.get('year')==year]
    checked=0; no_entrants=0; extra=0; cnt=0; wm=0; cm=0; missing=0
    ex=[]
    for x in recs:
        code=str(x.get('code')); entrants=x.get('entrants') or []
        em=emap.get(code)
        if entrants:
            if not em:
                no_entrants+=1
                MISSING.append((typ,year,code,x.get('city'),x.get('county')))
                if len(ex)<5: ex.append((code,'data有entrants但Excel无code'))
                continue
            yes_names={re.sub(r'\s+','',p['name']) for p in em['yes']}
            for e in entrants:
                if re.sub(r'\s+','',str(e.get('n'))) not in yes_names:
                    extra+=1
                    if len(ex)<10: ex.append((code,'entrants混入非进体检人:%s'%(e.get('n'))))
                    break
            if len(entrants)!=len(em['yes']):
                cnt+=1
                if len(ex)<10 and not any(code==e[0] for e in ex): ex.append((code,'entrants数%d vs Excel进体检数%d'%(len(entrants),len(em['yes']))))
            ws_v=[p['w'] for p in em['yes'] if p['w'] is not None]
            if ws_v:
                exp_w=min(ws_v); got=min((e['w'] for e in entrants if e.get('w') is not None),default=None)
                if got is None or abs(got-exp_w)>0.01:
                    wm+=1
                    if len(ex)<10 and not any(code==e[0] for e in ex): ex.append((code,'上岸笔试分 data=%s vs Excel=%s'%(got,exp_w)))
            cs_v=[p['c'] for p in em['yes'] if p['c'] is not None]
            if cs_v:
                exp_c=min(cs_v); gotc=x.get('onBoard')
                if gotc is None or abs(gotc-exp_c)>0.01:
                    cm+=1
                    if len(ex)<10 and not any(code==e[0] for e in ex): ex.append((code,'上岸综合分 data=%s vs Excel=%s'%(gotc,exp_c)))
            checked+=1
        else:
            if em and em['yes']:
                missing+=1
                if len(ex)<5 and not any(code==e[0] for e in ex): ex.append((code,'Excel有进体检但data无entrants'))
    print("\n==== %s %d 校验 ====" % (typ,year))
    print("总记录:%d  参与校验(有entrants且Excel有code):%d" % (len(recs), checked))
    print("data有entrants但Excel无code:%d | Excel有进体检但data无entrants:%d" % (no_entrants,missing))
    print("★entrants混入非进体检人(严重BUG):%d" % extra)
    print("人数不一致:%d | 上岸笔试分(min w)不一致:%d | 上岸综合分(onBoard)不一致:%d" % (cnt,wm,cm))
    if ex:
        print("--- 示例 ---")
        for e in ex: print("   ",e)
    return dict(checked=checked,no=no_entrants,extra=extra,cnt=cnt,wm=wm,cm=cm,missing=missing)

g=verify('公务员',2026,gwy_map)
s=verify('事业编',2026,sy_map)

sy_recs=[x for x in J if x.get('type')=='事业编' and x.get('year')==2026 and x.get('entrants')]
sy_in_excel=sum(1 for x in sy_recs if str(x['code']) in sy_map)
print("\n[覆盖] 事业编2026 有entrants记录=%d  其中Excel能查到=%d (%.1f%%)" % (len(sy_recs), sy_in_excel, 100*sy_in_excel/len(sy_recs)))
gw_recs=[x for x in J if x.get('type')=='公务员' and x.get('year')==2026 and x.get('entrants')]
gw_in_excel=sum(1 for x in gw_recs if str(x['code']) in gwy_map)
print("[覆盖] 公务员2026 有entrants记录=%d  其中Excel能查到=%d (%.1f%%)" % (len(gw_recs), gw_in_excel, 100*gw_in_excel/len(gw_recs)))

print("\n[无法核验的code明细] 共 %d 条:" % len(MISSING))
for m in MISSING:
    print("  ", m)
json.dump(MISSING, open(r"E:/workspace/civil-service-exam/_missing_codes.json","w",encoding="utf-8"), ensure_ascii=False)
