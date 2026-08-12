# -*- coding: utf-8 -*-
# 合并 2026 体检名单（含最终分数）到现有 data.js
import openpyxl, re, json, os

HERE = os.path.dirname(os.path.abspath(__file__))
DATA_JS = os.path.join(HERE, "web", "data.js")
GWY_FILE = os.path.join(HERE, "名睿整理—2026年陕西省考体检名单汇总表.xlsx")
SY_FILE = os.path.join(HERE, "26上事业单位体检名单.xlsx")
SY26_FILE = os.path.join(HERE, "26事业编岗位(1).xlsx")  # 官方26岗位表(含分数线)

PASS_OK = {"是", "进入体检", "取得体检资格"}
CODE_RE = re.compile(r"^\s*(\d{6,})")

def clean_num(v):
    if v is None:
        return None
    s = str(v).strip().replace("\u2002", "").replace("\xa0", "").replace(" ", "")
    if s in ("", "/", "\\", "缺考", "-"):
        return None
    try:
        return float(s)
    except Exception:
        return None

def find_header(rows):
    best = None
    bestscore = -1
    for ri, row in enumerate(rows):
        score = 0
        for c in row:
            cs = "" + str(c) if c is not None else ""
            for kw in ["代码", "职位", "笔试", "面试", "综合成绩", "总成绩", "是否", "准考证", "姓名", "单位"]:
                if kw in cs:
                    score += 1
                    break
        if score > bestscore:
            bestscore = score
            best = ri
    return best if bestscore >= 4 else None

def col(header, kws, exclude=None):
    for i, c in enumerate(header):
        cs = "" + str(c) if c is not None else ""
        cs = re.sub(r"\s+", "", cs)  # 归一化：去掉表头内空格（如“姓  名”→“姓名”）
        if exclude and exclude in cs:
            continue
        for kw in kws:
            if kw in cs:
                return i
    return None

def detect_code_col(ws, hr):
    header = None
    counts = {}
    seen = 0
    for ri, r in enumerate(ws.iter_rows(min_row=1, max_row=min(ws.max_row, 120), values_only=True)):
        if ri == hr:
            header = r
        if ri < hr + 1:
            continue
        if seen > 80:
            break
        seen += 1
        for ci, v in enumerate(r):
            if v is None:
                continue
            if CODE_RE.match(str(v)):
                counts[ci] = counts.get(ci, 0) + 1
    if not counts:
        return None
    best = None
    bestc = -1
    for ci, c in counts.items():
        hcs = "" + str(header[ci]) if header and ci < len(header) else ""
        if "准考证" in hcs:
            continue
        if c > bestc:
            bestc = c
            best = ci
    return best if bestc >= 1 else None

CITY_LIST = ["西安", "宝鸡", "咸阳", "渭南", "延安", "榆林", "汉中", "安康", "铜川", "商洛", "杨凌"]

def norm_city(sheet, kind):
    s = sheet.strip()
    if kind == "公务员":
        s = s.replace("市", "")
        if s.startswith("省"):
            return "省直"
        if s.startswith("杨凌"):
            return "杨凌"
        return s
    # 事业编
    for c in CITY_LIST:
        if s.startswith(c):
            return c
    if s.startswith("陕西"):
        return "省直"
    return "省直"

def load_existing_jobs():
    txt = open(DATA_JS, encoding="utf-8").read()
    m = re.search(r"window\.JOBS\s*=\s*(\[.*?\])\s*;\s*window\.META", txt, re.S)
    return json.loads(m.group(1)), txt

# ---------- 岗位类别归一化（供前端筛选）----------

def sy_cat(exam_cat):
    """事业编笔试类别 → 简洁标签：A类(综合管理)/B类/C类/D类(教师)/E类(医疗)/其他"""
    s = re.sub(r"\s+", "", str(exam_cat or ""))
    if "综合管理" in s:
        return "A类(综合管理)"
    if "自然科学" in s:
        return "C类(自然)"
    if "社会科学" in s:
        return "B类(社科)"
    if ("教师" in s or "中小学" in s) and "医疗" not in s:
        return "D类(教师)"
    if "医疗" in s or any(k in s for k in ["西医", "中医", "护理", "药剂", "医学技术", "公共卫生"]):
        return "E类(医疗)"
    if s == "":
        return "未分类"
    return "其他"

def gwy_cat(j):
    """公务员岗位性质 → 常规/公安/法院/检察院/司法(监狱戒毒)（按关键词优先级）"""
    blob = " ".join(str(j.get(k) or "") for k in ["position", "desc", "dept", "org"])
    if "公安" in blob:
        return "公安"
    if "监狱" in blob or "戒毒" in blob:
        return "司法(监狱戒毒)"
    if "法院" in blob:
        return "法院"
    if "检察" in blob:
        return "检察院"
    if "司法" in blob:
        return "司法行政"
    return "常规"

def calc_cat(j):
    return sy_cat(j.get("examCat")) if j.get("type") == "事业编" else gwy_cat(j)

def iter_sheet(wb, sheet, kind):
    """yield (code, name, written, interview, composite, org, position) per pass row"""
    ws = wb[sheet]
    if ws.max_row > 20000:
        return  # 误入的目录/索引 sheet，跳过
    rows = list(ws.iter_rows(min_row=1, max_row=min(ws.max_row, 14), values_only=True))
    hr = find_header(rows)
    if hr is None:
        return
    header = rows[hr]
    c_code = detect_code_col(ws, hr)
    c_name = col(header, ["姓名"], exclude="准考证")
    c_wr = col(header, ["笔试"], exclude="准考证")
    c_iv = col(header, ["面试"], exclude="准考证")
    c_fin = col(header, ["综合成绩"], exclude="笔试")
    if c_fin is None:
        c_fin = col(header, ["总成绩", "考试总成绩", "最终综合成绩"], exclude="笔试")
    c_pass = col(header, ["是否"], exclude="准考证")
    c_pos = col(header, ["岗位", "职位"], exclude="代码")
    c_org = col(header, ["单位", "主管"], exclude="代码")
    city = norm_city(sheet, kind)
    for r in ws.iter_rows(min_row=hr + 2, values_only=True):
        p = r[c_pass] if (c_pass is not None and c_pass < len(r)) else None
        ps = str(p).strip() if p is not None else ""
        if ps not in PASS_OK:
            continue
        # code + position text
        code = None
        pos_text = None
        if c_code is not None and c_code < len(r) and r[c_code] is not None:
            raw = str(r[c_code]).strip()
            m = CODE_RE.match(raw)
            if m:
                code = m.group(1)
                pos_text = raw[m.end():].strip()
        if not pos_text and c_pos is not None and c_pos < len(r) and r[c_pos] is not None:
            pos_text = str(r[c_pos]).strip()
        org = ""
        if c_org is not None and c_org < len(r) and r[c_org] is not None:
            org = str(r[c_org]).strip()
        name = ""
        if c_name is not None and c_name < len(r) and r[c_name] is not None:
            name = str(r[c_name]).strip()
        if not code and not pos_text:
            continue  # 跳过无法定位岗位的脏行
        written = clean_num(r[c_wr]) if (c_wr is not None and c_wr < len(r)) else None
        interview = clean_num(r[c_iv]) if (c_iv is not None and c_iv < len(r)) else None
        composite = clean_num(r[c_fin]) if (c_fin is not None and c_fin < len(r)) else None
        yield (code, name, written, interview, composite, org, pos_text or "", city)

def build_records(fname, kind, log):
    wb = openpyxl.load_workbook(fname, read_only=True, data_only=True)
    groups = {}
    for sheet in wb.sheetnames:
        for (code, name, written, interview, composite, org, pos, city) in iter_sheet(wb, sheet, kind):
            key = (kind, city, code if code else ("_" + org + "|" + pos))
            g = groups.get(key)
            if g is None:
                g = {"type": kind, "year": 2026, "city": city, "county": None,
                     "dept": "", "org": org, "position": pos, "desc": "",
                     "code": code, "major": "", "edu": "", "degree": "",
                     "political": "", "other": "", "examCat": None,
                     "plan": 0, "competition": None,
                     "scores": {}, "onBoard": None, "entrants": []}
                groups[key] = g
            g["plan"] += 1
            if written is not None or interview is not None or composite is not None:
                g["entrants"].append({"n": name, "w": written, "i": interview, "c": composite})
    wb.close()
    recs = []
    for g in groups.values():
        ws = [e["w"] for e in g["entrants"] if e["w"] is not None]
        cs = [e["c"] for e in g["entrants"] if e["c"] is not None]
        g["scores"] = {2026: (min(ws) if ws else None)}
        g["onBoard"] = min(cs) if cs else None
        recs.append(g)
    log.append((kind, len(recs), sum(g["plan"] for g in recs)))
    return recs

# ---------- 2026 事业编：官方岗位表（Sheet2）----------

def _find_col(header, *kws):
    for i, c in enumerate(header):
        cs = re.sub(r"\s+", "", "" + str(c) if c is not None else "")
        for k in kws:
            if k in cs:
                return i
    return None

def build_sy26_official():
    """Sheet2 → 2026 事业编主记录 dict(code->rec)：完整岗位信息 + 官方 26最低分数线"""
    wb = openpyxl.load_workbook(SY26_FILE, read_only=True, data_only=True)
    ws = wb["Sheet2"]
    header = list(ws.iter_rows(min_row=1, max_row=1, values_only=True))[0]
    c_code = _find_col(header, "岗位代码")
    c_city = _find_col(header, "地市"); c_county = _find_col(header, "地县")
    c_dept = _find_col(header, "主管部门"); c_org = _find_col(header, "事业单位名称")
    c_pos = _find_col(header, "岗位名称")
    c_major = _find_col(header, "专业"); c_edu = _find_col(header, "学历")
    c_degree = _find_col(header, "学位"); c_pol = _find_col(header, "政治面貌")
    c_other = _find_col(header, "其他条件"); c_plan = _find_col(header, "招聘人数")
    c_min = _find_col(header, "26最低分数线"); c_exam = _find_col(header, "笔试类别")
    out = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        code = None
        if c_code is not None and c_code < len(r) and r[c_code] is not None:
            m = CODE_RE.match(str(r[c_code]))
            if m:
                code = m.group(1)
        if not code:
            continue
        g = lambda idx: (str(r[idx]).strip() if (idx is not None and idx < len(r) and r[idx] is not None) else "")
        minline = clean_num(r[c_min]) if (c_min is not None and c_min < len(r)) else None
        plan = clean_num(r[c_plan]) if (c_plan is not None and c_plan < len(r)) else 0
        out[code] = {
            "type": "事业编", "year": 2026,
            "city": norm_city(g(c_city), "事业编"),
            "county": g(c_county) or None,
            "dept": g(c_dept), "org": g(c_org), "position": g(c_pos), "desc": "",
            "code": code, "major": g(c_major), "edu": g(c_edu), "degree": g(c_degree),
            "political": g(c_pol), "other": g(c_other), "examCat": g(c_exam) or None,
            "plan": int(plan) if plan else 0, "competition": None,
            "scores": {2026: minline} if minline is not None else {},
            "onBoard": None, "entrants": [],
        }
    wb.close()
    return out

def build_sy26_sheet1():
    """Sheet1 → 2026 事业编补充岗位（无分数线）"""
    wb = openpyxl.load_workbook(SY26_FILE, read_only=True, data_only=True)
    ws = wb["Sheet1"]
    header = list(ws.iter_rows(min_row=1, max_row=1, values_only=True))[0]
    c_code = _find_col(header, "岗位代码")
    c_city = _find_col(header, "地市"); c_county = _find_col(header, "地县")
    c_dept = _find_col(header, "主管部门"); c_org = _find_col(header, "事业单位名称")
    c_pos = _find_col(header, "岗位简称", "岗位名称")
    c_major = _find_col(header, "专业"); c_edu = _find_col(header, "学历")
    c_degree = _find_col(header, "学位"); c_other = _find_col(header, "其他条件")
    c_plan = _find_col(header, "招聘人数"); c_exam = _find_col(header, "笔试类别")
    recs = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        code = None
        if c_code is not None and c_code < len(r) and r[c_code] is not None:
            m = CODE_RE.match(str(r[c_code]))
            if m:
                code = m.group(1)
        if not code:
            continue
        g = lambda idx: (str(r[idx]).strip() if (idx is not None and idx < len(r) and r[idx] is not None) else "")
        plan = clean_num(r[c_plan]) if (c_plan is not None and c_plan < len(r)) else 0
        recs.append({
            "type": "事业编", "year": 2026,
            "city": norm_city(g(c_city), "事业编"),
            "county": g(c_county) or None,
            "dept": g(c_dept), "org": g(c_org), "position": g(c_pos), "desc": "",
            "code": code, "major": g(c_major), "edu": g(c_edu), "degree": g(c_degree),
            "political": "", "other": g(c_other), "examCat": g(c_exam) or None,
            "plan": int(plan) if plan else 0, "competition": None,
            "scores": {}, "onBoard": None, "entrants": [],
        })
    wb.close()
    return recs

def build_sy26_supplement():
    """体检名单 → 2026 事业编补充：进面人员明细 entrants + 算出的上岸线 onBoard（按 code 匹配）"""
    wb = openpyxl.load_workbook(SY_FILE, read_only=True, data_only=True)
    supp = {}  # code -> {city,org,position,entrants,onBoard,plan}
    for sheet in wb.sheetnames:
        for (code, name, written, interview, composite, org, pos, city) in iter_sheet(wb, sheet, "事业编"):
            if not code:
                continue
            if code not in supp:
                supp[code] = {"city": city, "org": org, "position": pos,
                              "entrants": [], "onBoard": None, "plan": 0}
            s = supp[code]
            s["plan"] += 1
            if written is not None or interview is not None or composite is not None:
                s["entrants"].append({"n": name, "w": written, "i": interview, "c": composite})
    for s in supp.values():
        cs = [e["c"] for e in s["entrants"] if e["c"] is not None]
        s["onBoard"] = min(cs) if cs else None
    wb.close()
    return supp

def main():
    jobs, txt = load_existing_jobs()
    log = []
    # 公务员 2026（省考体检名单）—— 不变
    gwy = build_records(GWY_FILE, "公务员", log)
    # 事业编 2026：以官方岗位表(Sheet2)为主，附体检名单明细 + Sheet1 补充
    sy_main = build_sy26_official()       # dict code->rec（完整字段+官方线）
    sy_supp = build_sy26_supplement()     # dict code->明细+上岸线
    sy_s1 = build_sy26_sheet1()           # 补充岗位 list
    sy_recs = []
    n_entrants = 0
    for code, rec in sy_main.items():
        sup = sy_supp.get(code)
        if sup:
            rec["entrants"] = sup["entrants"]
            n_entrants += len(sup["entrants"])
            if sup["onBoard"] is not None:
                rec["onBoard"] = sup["onBoard"]
            if not rec["scores"].get(2026) and sup["entrants"]:  # 官方无线则补算进面线
                ws = [e["w"] for e in sup["entrants"] if e["w"] is not None]
                if ws:
                    rec["scores"][2026] = min(ws)
        sy_recs.append(rec)
    for code, sup in sy_supp.items():  # 体检名单独有（不在 Sheet2）
        if code not in sy_main:
            ws = [e["w"] for e in sup["entrants"] if e["w"] is not None]
            rec = {"type": "事业编", "year": 2026, "city": sup["city"], "county": None,
                   "dept": "", "org": sup["org"], "position": sup["position"], "desc": "",
                   "code": code, "major": "", "edu": "", "degree": "", "political": "",
                   "other": "", "examCat": None, "plan": sup["plan"], "competition": None,
                   "scores": {2026: min(ws)} if ws else {}, "onBoard": sup["onBoard"],
                   "entrants": sup["entrants"]}
            n_entrants += len(sup["entrants"])
            sy_recs.append(rec)
    for r in sy_s1:  # Sheet1 独有 code
        if r["code"] not in sy_main:
            sy_recs.append(r)
    log.append(("事业编", len(sy_recs), n_entrants))
    jobs.extend(gwy)
    jobs.extend(sy_recs)
    for i, j in enumerate(jobs):
        j["id"] = i + 1  # 1-based 唯一 id，供详情定位
        # 归一化学历字段：去除不规则空格（"本科 及 以上" → "本科及以上"）
        if j.get("edu"):
            j["edu"] = re.sub(r"\s+", "", str(j["edu"]))
        # 岗位类别归一化（事业编按笔试类别 / 公务员按岗位性质）
        j["cat"] = calc_cat(j)
    # 重写 data.js，保留 META
    m = re.search(r"(window\.META\s*=\s*\{.*?\};\s*)", txt, re.S)
    meta = m.group(1) if m else "window.META = {};\n"
    out = "window.JOBS = " + json.dumps(jobs, ensure_ascii=False, separators=(",", ":")) + ";\n" + meta
    with open(DATA_JS, "w", encoding="utf-8") as f:
        f.write(out)
    print("原 JOBS 条数(基底):", len(jobs) - len(gwy) - len(sy_recs))
    for kind, nrec, nent in log:
        print("  新增 %s: 岗位记录 %d 条, 进面人员 %d 人" % (kind, nrec, nent))
    print("合并后 JOBS 总条数:", len(jobs))
    sizes = os.path.getsize(DATA_JS)
    print("data.js 大小: %.2f MB" % (sizes / 1024 / 1024))

if __name__ == "__main__":
    main()
