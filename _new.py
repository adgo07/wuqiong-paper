# -*- coding: utf-8 -*-
import io, os, csv, collections, traceback
out = io.StringIO()
def p(*a): print(*a, file=out)
base = r"F:\布病"

# --- 1. 银川市-气象资料数据.xls ---
try:
    import pandas as pd
    path = os.path.join(base, "银川市-气象资料数据.xls")
    xl = pd.ExcelFile(path)
    p("### FILE: 银川市-气象资料数据.xls")
    p("SHEETS:", xl.sheet_names)
    for sh in xl.sheet_names:
        df = pd.read_excel(path, sheet_name=sh, header=None)
        p(f"  -- SHEET [{sh}] shape={df.shape}")
        for i in range(min(8, len(df))):
            p(f"    r{i:>3}: " + " | ".join(str(x)[:20] for x in df.iloc[i].tolist()))
        p("    ...")
        for i in range(max(0,len(df)-4), len(df)):
            p(f"    r{i:>3}: " + " | ".join(str(x)[:20] for x in df.iloc[i].tolist()))
        # try headers
        try:
            df2 = pd.read_excel(path, sheet_name=sh)
            p("    parsed columns:", list(df2.columns))
            p("    shape:", df2.shape, "dtypes:", dict(df2.dtypes.astype(str)))
            p("    nunique:", {c:int(df2[c].nunique()) for c in df2.columns})
            p("    head(3):"); p(df2.head(3).to_string(max_cols=20))
            p("    tail(3):"); p(df2.tail(3).to_string(max_cols=20))
        except Exception as e:
            p("    header parse err:", repr(e))
except Exception as e:
    p("META ERROR:", repr(e)); p(traceback.format_exc())

# --- 2. 两个 CSV ---
for f in ["报告卡2026-09-14+11_03_21.csv", "按照现住址报告卡2026-09-14+11_05_47.csv"]:
    p("\n" + "="*90)
    p("### FILE:", f)
    path = os.path.join(base, f)
    raw = open(path,"rb").read()
    p("bytes:", len(raw))
    txt=None
    for enc in ("utf-8-sig","utf-8","gbk","gb18030"):
        try: txt=raw.decode(enc); p("encoding:", enc); break
        except Exception: pass
    if txt is None:
        p("decode failed"); continue
    lines = txt.splitlines()
    p("lines:", len(lines))
    p("header:", lines[0] if lines else "")
    p("row1 :", lines[1] if len(lines)>1 else "")
    rows = list(csv.DictReader(io.StringIO(txt)))
    p("n rows:", len(rows))
    p("fields:", list(rows[0].keys()) if rows else None)
    import datetime
    def yr(s):
        s=(s or "").strip()
        for fmt in ("%Y/%m/%d","%Y-%m-%d","%Y.%m.%d"):
            try: return datetime.datetime.strptime(s,fmt).year
            except Exception: pass
        return None
    # find a date-ish column
    datecols = [c for c in (rows[0].keys() if rows else []) if "日期" in c or "时间" in c]
    p("date-like cols:", datecols)
    for dc in datecols:
        c = collections.Counter(yr(r.get(dc)) for r in rows)
        p(f"  year dist for [{dc}]:")
        for k in sorted(x for x in c if x): p(f"    {k}: {c[k]}")
        p("    None:", c.get(None,0))
    # region col
    for c in (rows[0].keys() if rows else []):
        if "国标" in c or "地区" in c or "县区" in c or "地址" in c:
            cc = collections.Counter(str(r.get(c))[:6] for r in rows)
            p(f"  [{c}] top:", cc.most_common(10))
    if "疾病名称" in (rows[0].keys() if rows else []):
        p("  疾病名称:", collections.Counter(r.get("疾病名称") for r in rows).most_common(10))

with open(os.path.join(base,"_new_out.txt"),"w",encoding="utf-8") as fh:
    fh.write(out.getvalue())
print("done")
