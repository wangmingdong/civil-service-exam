# -*- coding: utf-8 -*-
# 回填事业编/公务员 2026 上岸综合分：从体检名单 Excel 按姓名匹配填补 entrants[].c，并重算 onBoard=min(c)
import json, re, openpyxl, os
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

def find_header(ws):
    rows = list(ws.iter_rows(min_row=1, max_row=min(ws.max_row,30), values_only=True))
    for ri in range(len(rows)):
        cur = " ".join(norm(c) for c in rows[ri])
        nxt = " ".join(norm(c) for c in rows[ri+1]) if ri+1<len(rows) else ""
        if ("姓名" in cur or "准考证" in cur) and ("体检" in cur or "体检" in nxt):
            return ri
    return None

def map_cols(header):
    code=name=written=comp=tijian=interview=None
    for ci,c in enumerate(header):
        s = norm(c)
        if not s: continue
        if code is None and ('岗位代码' in s or '职位代码' in s or '报考职位' in s or '职位' in s or '岗位' in s):
            code = ci
        if name is None and '姓名' in s: name = ci
        if written is None and '笔试' in s and ('总' in s or '成绩' in s): written = ci
        if interview is None and '面试' in s and '成绩' in s: interview = ci
        if comp is None and (('综合成绩' in s) or ('考试总成绩' in s) or (s.startswith('总成绩') and '笔试' not in s[:4])):
            comp = ci
        if tijian is None and '体检' in s and ('是否' in s or '进入' in s or '资格' in s): tijian = ci
    return code,name,written,comp,interview,tijian

def is_yes(v):
    s=norm(v)
    if not s: return False
    return s in ('是','进入体检','√','是√') or ('是' in s and '不' not in s and s!='')

def parse(path, skip=()):
    wb=openpyxl.load_workbook(path, read_only=True, data_only=True)
    out={}  # code -> {name: {w,i,c}}
    for sh in wb.sheetnames:
        if sh in skip: continue
        ws=wb[sh]; hr=find_header(ws)
        if hr is None: continue
        hrow=[c for c in next(ws.iter_rows(min_row=hr+1,max_row=hr+1,values_only=True))]
        code_c,name_c,w_c,comp_c,i_c,tij_c=map_cols(hrow)
        if code_c is None or tij_c is None: continue
        for r in ws.iter_rows(min_row=hr+2, values_only=True):
            if code_c>=len(r) or r[code_c] is None: continue
            m=re.match(r'(\d+)', norm(r[code_c]))
            if not m: continue
            code=m.group(1)
            nm=norm(r[name_c]) if name_c is not None and name_c<len(r) and r[name_c] else ''
            if not nm: continue
            w=clean_num(r[w_c]) if w_c is not None and w_c<len(r) else None
            i=clean_num(r[i_c]) if i_c is not None and i_c<len(r) else None
            c=clean_num(r[comp_c]) if comp_c is not None and comp_c<len(r) else None
            yes=is_yes(r[tij_c]) if tij_c<len(r) else False
            if not yes: continue
            out.setdefault(code,{})[nm]={'w':w,'i':i,'c':c}
    wb.close()
    return out

emap = {}
emap.update(parse(GWY))
emap.update(parse(SY, skip=('Sheet9',)))
print("Excel 进体检人员 map: codes=%d" % len(emap))

# ---- 加载 data.js ----
txt=open(DATA,encoding='utf-8').read()
si=txt.index('window.JOBS='); i=txt.index('[',si); depth=0
for j in range(i,len(txt)):
    if txt[j]=='[':depth+=1
    elif txt[j]==']':
        depth-=1
        if depth==0: break
J=json.loads(txt[i:j+1])

filled_c=0; filled_onboard=0; touched=0
for x in J:
    if x.get('year')!=2026: continue
    entrants=x.get('entrants') or []
    if not entrants: continue
    code=str(x.get('code'))
    pe=emap.get(code)
    changed=False
    for e in entrants:
        en=norm(e.get('n'))
        if e.get('c') is None and pe and en in pe:
            cval=pe[en]['c']
            if cval is not None and 40<=cval<=100:
                e['c']=cval; filled_c+=1; changed=True
    if changed:
        touched+=1
        cs=[e['c'] for e in entrants if e.get('c') is not None]
        if cs:
            new_on=round(min(cs),2)
            if x.get('onBoard') is None:
                x['onBoard']=new_on; filled_onboard+=1; changed=True
    # 若 onBoard 仍为空但 entrants 有 c，也补
    if x.get('onBoard') is None:
        cs=[e['c'] for e in entrants if e.get('c') is not None]
        if cs:
            x['onBoard']=round(min(cs),2); filled_onboard+=1; touched+=1

print("回填 entrants.c 人次:%d  补 onBoard 记录:%d  涉及记录数:%d" % (filled_c, filled_onboard, touched))

# ---- 写回 data.js（保持原有序列化风格：紧凑无空格）----
out=txt[:i] + json.dumps(J, ensure_ascii=False) + txt[j+1:]
open(DATA,'w',encoding='utf-8').write(out)
print("已写回 data.js，大小=%d" % len(out))
# 备份
import shutil
shutil.copy(DATA, r"E:/workspace/civil-service-exam/web/data.js.bak")
print("已备份至 web/data.js.bak")
