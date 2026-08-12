# -*- coding: utf-8 -*-
# 最终回填：用健壮解析从体检名单 Excel 取进体检人员(w,i,c)，
# 1) 补 data.js 中 entrant 空姓名(n)；2) 补 entrant.c 与 onBoard(上岸综合分) 缺口。
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

def norm(s): return re.sub(r'\s+','',str(s)) if s is not None else ''

PAT = re.compile(r'(?<!\d)\d{9,13}(?!\d)')
def find_header_rows(ws):
    rows = list(ws.iter_rows(values_only=True)); n=min(len(rows),40)
    for ri in range(n):
        cur=' '.join(norm(c) for c in rows[ri])
        nxt=' '.join(norm(c) for c in rows[ri+1]) if ri+1<n else ''
        if ('姓名' in cur or '准考证' in cur) and ('体检' in cur or '体检' in nxt):
            comp=[]; nxt_row=rows[ri+1] if ri+1<n else []
            for ci in range(max(len(rows[ri]),len(nxt_row))):
                a=norm(rows[ri][ci]) if ci<len(rows[ri]) else ''
                b=norm(nxt_row[ci]) if ci<len(nxt_row) else ''
                comp.append(a+b)
            return ri,comp,rows
    return None,None,rows

def map_cols(header):
    name=written=comp=tijian=None
    for ci,c in enumerate(header):
        s=norm(c)
        if not s: continue
        if name is None and '姓名' in s: name=ci
        if written is None and '笔试' in s and ('总' in s or '成绩' in s): written=ci
        if comp is None and (('综合成绩' in s) or ('考试总成绩' in s) or (s.startswith('总成绩') and '笔试' not in s[:4])): comp=ci
        if tijian is None and '体检' in s and ('是否' in s or '进入' in s or '资格' in s): tijian=ci
    return name,written,comp,tijian

def best_code_col(rows, hr, header):
    best=None;bestcnt=-1;data=rows[hr+1:hr+200]
    width=max((len(r) for r in data),default=0)
    for ci in range(width):
        hc=header[ci] if ci<len(header) else ''
        hh=hc.replace(' ','')
        if hh in ('准考证号','准考证') or hh.startswith('准考证号'): continue
        cnt=sum(1 for r in data if ci<len(r) and r[ci] is not None and PAT.search(norm(r[ci])))
        if cnt>bestcnt: bestcnt=cnt;best=ci
    return best if bestcnt>=2 else None

def is_yes(v):
    s=norm(v)
    if not s: return False
    return s in ('是','进入体检','√','是√') or ('是' in s and '不' not in s and s!='')

def parse(path, skip=()):
    wb=openpyxl.load_workbook(path, data_only=True)
    out={}  # code -> {name: {'w':,'i':,'c':}}
    for sh in wb.sheetnames:
        if sh in skip: continue
        ws=wb[sh]; hr,header,rows=find_header_rows(ws)
        if hr is None: continue
        name_c,w_c,comp_c,tij_c=map_cols(header)
        code_c=best_code_col(rows,hr,header)
        if code_c is None or tij_c is None: continue
        for r in rows[hr+1:]:
            if code_c>=len(r) or r[code_c] is None: continue
            m=PAT.search(norm(r[code_c]))
            if not m: continue
            code=m.group(0)
            nm=norm(r[name_c]) if name_c is not None and name_c<len(r) and r[name_c] else ''
            if not nm: continue
            w=clean_num(r[w_c]) if w_c is not None and w_c<len(r) else None
            i=clean_num(r[comp_c-1]) if comp_c is not None and comp_c-1<len(r) else None  # 面试成绩在综合前
            c=clean_num(r[comp_c]) if comp_c is not None and comp_c<len(r) else None
            yes=is_yes(r[tij_c]) if tij_c<len(r) else False
            if not yes: continue
            out.setdefault(code,{})[nm]={'w':w,'i':i,'c':c}
    wb.close()
    return out

emap={}
emap.update(parse(GWY))
emap.update(parse(SY, skip=('Sheet9',)))
print("Excel 进体检人员 code->name 映射: codes=%d" % len(emap))

# ---- 载入 data.js ----
txt=open(DATA,encoding='utf-8').read()
si=txt.index('window.JOBS='); i=txt.index('[',si); depth=0
for j in range(i,len(txt)):
    if txt[j]=='[':depth+=1
    elif txt[j]==']':
        depth-=1
        if depth==0: break
J=json.loads(txt[i:j+1])

filled_name=0; filled_c=0; filled_onboard=0; touched=0
for x in J:
    if x.get('year')!=2026: continue
    entrants=x.get('entrants') or []
    if not entrants: continue
    code=str(x.get('code'))
    pe=emap.get(code)
    if not pe: continue
    changed=False
    for e in entrants:
        en=norm(e.get('n'))
        ew=e.get('w'); ec=e.get('c')
        # 补空姓名：按 w(及c) 匹配唯一 Excel 进体检人
        if not en:
            cands=[nm for nm,v in pe.items() if v['w'] is not None and ew is not None and abs(v['w']-ew)<0.01
                   and (ec is None or (v['c'] is not None and abs(v['c']-ec)<0.01))]
            if len(cands)==1:
                e['n']=cands[0]; filled_name+=1; changed=True
        # 补 entrant.c：按姓名匹配
        if e.get('c') is None:
            nm=norm(e.get('n'))
            if nm in pe and pe[nm]['c'] is not None and 40<=pe[nm]['c']<=100:
                e['c']=pe[nm]['c']; filled_c+=1; changed=True
    # 补 onBoard
    if changed or x.get('onBoard') is None:
        cs=[e['c'] for e in entrants if e.get('c') is not None]
        if cs:
            new_on=round(min(cs),2)
            if x.get('onBoard') is None or abs((x.get('onBoard') or 0)-new_on)>0.01:
                x['onBoard']=new_on; filled_onboard+=1; touched+=1
                changed=True
print("补空姓名人次:%d  补entrants.c人次:%d  补onBoard记录:%d  涉及记录:%d" % (filled_name,filled_c,filled_onboard,touched))

out=txt[:i] + json.dumps(J, ensure_ascii=False) + txt[j+1:]
open(DATA,'w',encoding='utf-8').write(out)
import shutil
shutil.copy(DATA, r"E:/workspace/civil-service-exam/web/data.js.bak")
print("已写回 data.js，大小=%d，备份 web/data.js.bak" % len(out))
