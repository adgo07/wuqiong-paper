# -*- coding: utf-8 -*-
import io, os, csv, collections, datetime, traceback
out = io.StringIO()
def p(*a): print(*a, file=out)
base = r"F:\布病"

# ============ A. 新气象日值表 ============
import pandas as pd
path = os.path.join(base, "银川市-气象资料数据.xls")
df = pd.read_excel(path, sheet_name="sheet1")
df.columns = [str(c).strip() for c in df.columns]
p("### A. 银川市-气象资料数据.xls")
p("shape:", df.shape)
df["d"] = pd.to_datetime(df["监测时间"], errors="coerce")
p("日期范围:", df["d"].min(), "→", df["d"].max())
p("行政区划码 unique:", df["行政区划码"].unique())
p("省 unique:", df["省"].unique())
p("市 unique:", df["市"].unique())
# 每日唯一性
dup = df["d"].duplicated().sum()
p("重复日期数:", dup)
# 每年天数
byyear = df.groupby(df["d"].dt.year)["d"].count()
p("\n每年记录天数:")
for k,v in byyear.items(): p(f"   {k}: {v}")
# 缺失率 by column by year
cols = [c for c in df.columns if c not in ("省","市","行政区划码","监测时间","d","霾日")]
p("\n各变量缺失率 by year (%):")
p("变量".ljust(18) + "".join(f"{y:>7}" for y in sorted(byyear.index)))
for c in cols:
    line = c.ljust(18)
    for y in sorted(byyear.index):
        sub = df[df["d"].dt.year==y]
        line += f"{sub[c].isna().mean()*100:>7.1f}"
    p(line)
# 霾日 取值
p("\n霾日 取值:", df["霾日"].value_counts(dropna=False).to_dict())
# 描述统计
p("\n描述统计:")
p(df[cols].describe().T.to_string())
# 极端值检查
p("\n异常检查:")
for c in ["平均温度(℃)","降水量(mm)","日平均风速(m/s)","平均相对湿度(%)","日照小时数(小时/日)"]:
    s = df[c].dropna()
    p(f"  {c}: min={s.min()} p1={s.quantile(.01):.2f} p50={s.median():.2f} p99={s.quantile(.99):.2f} max={s.max()}")

# 与旧表重叠期比较 (2019-2024, 银川市区)
p("\n### 重叠期一致性 (新月值 vs 旧4区域月值, 2019-2024)")
oldpath = os.path.join(base, "气象资料.xls")
o = pd.read_excel(oldpath, sheet_name="Sheet1")
o.columns = [str(c).strip() for c in o.columns]
o["ym"] = o["日期"].astype(str).str.strip()
df["ym"] = df["d"].dt.strftime("%Y-%m")
nm = df.groupby("ym").agg(temp=("平均温度(℃)","mean"), rh=("平均相对湿度(%)","mean"),
                          prcp=("降水量(mm)","sum"), wind=("日平均风速(m/s)","mean"),
                          sun=("日照小时数(小时/日)","sum")).round(2)
mapping = {"平均温度(℃)":"temp","月平均相对湿度（%）":"rh","月总降水量（mm）":"prcp",
           "月平均风速（m/s）":"wind","日照时数（h）":"sun"}
oldcols = {c.strip():c for c in o.columns}
for oldname, newname in mapping.items():
    if oldname in oldcols:
        sub = o[o["区域"]=="银川市区"][["ym", oldname]].dropna()
        j = sub.merge(nm.reset_index(), left_on="ym", right_on="ym", how="inner")
        if len(j):
            r = j[oldname].astype(float).corr(j[newname].astype(float))
            bias = (j[newname].astype(float) - j[oldname].astype(float)).mean()
            p(f"  {oldname:>20} vs 新[{newname}]: n={len(j)} r={r:.4f} 平均偏差={bias:+.3f}")
        else:
            p(f"  {oldname}: 无重叠匹配")
# 也看其他区域
p("\n  各区域对比 (温度):")
for reg in o["区域"].dropna().unique():
    try:
        sub = o.loc[o["区域"]==reg, ["ym","月平均气温(℃)"]].dropna()
        j = sub.merge(nm.reset_index(), left_on="ym", right_on="ym", how="inner")
        if len(j):
            r = j["月平均气温(℃)"].astype(float).corr(j["temp"].astype(float))
            bias = (j["temp"].astype(float)-j["月平均气温(℃)"].astype(float)).mean()
            p(f"    {reg}: r={r:.4f} 偏差={bias:+.3f} n={len(j)}")
    except Exception as e:
        p(f"    {reg}: ERR {e!r}")

# ============ B. 2025 病例 ============
p("\n" + "="*90)
p("### B. 2025 年个案（两份导出）")
for f, label in [("报告卡2026-09-14+11_03_21.csv","按报告单位（银川市医疗机构报告）"),
                 ("按照现住址报告卡2026-09-14+11_05_47.csv","按现住址（银川市居民）")]:
    raw = open(os.path.join(base,f),"rb").read().decode("gbk")
    rows = list(csv.DictReader(io.StringIO(raw)))
    p(f"\n--- {label} | {f} | n={len(rows)}")
    p("  卡片状态:", collections.Counter(r["卡片状态"] for r in rows))
    p("  病例分类:", collections.Counter(r["病例分类"] for r in rows))
    p("  病例分类2:", collections.Counter(r["病例分类2"] for r in rows))
    p("  病人属于:", collections.Counter(r["病人属于"] for r in rows))
    p("  审核状态:", collections.Counter(r["审核状态"] for r in rows))
    p("  删除/标注时间非空:", sum(1 for r in rows if r["（删除/标注）时间"].strip() not in ("",".","")))
    p("  未纳入统计原因非空:", collections.Counter(r["（删除/未纳入统计）原因"].strip() for r in rows))
    p("  性别:", collections.Counter(r["性别"] for r in rows))
    p("  疾病名称:", collections.Counter(r["疾病名称"] for r in rows))
    # 重复卡片ID
    ids = [r["卡片ID"].lstrip("'") for r in rows]
    p("  unique 卡片ID:", len(set(ids)), " / n =", len(ids))
    # 卡片编号重复（同一病例订正）
    nums = [r["卡片编号"] for r in rows]
    p("  unique 卡片编号:", len(set(nums)))
    dupnums = [k for k,v in collections.Counter(nums).items() if v>1]
    if dupnums: p("  重复卡片编号:", dupnums[:10])
    # 发病日期范围
    ds = []
    for r in rows:
        s = r["发病日期"].strip()
        for fmt in ("%Y-%m-%d","%Y/%m/%d"):
            try: ds.append(datetime.datetime.strptime(s,fmt)); break
            except Exception: pass
    p("  发病日期范围:", min(ds).date(), "→", max(ds).date())
    md = collections.Counter(d.month for d in ds)
    p("  月度分布:", [f"{m}:{md.get(m,0)}" for m in range(1,13)])
    # 国标码前6
    c6 = collections.Counter((r["现住地址国标"] or "").strip()[:6] for r in rows)
    p("  现住址国标前6:", c6.most_common(12))
    # 2025 银川市居民数（前6位=6401xx）
    yn = sum(1 for r in rows if (r["现住地址国标"] or "").strip().startswith("6401"))
    p("  >>> 现住址在银川市(6401xx)例数:", yn)
    # 乡镇级
    c9 = set((r["现住地址国标"] or "").strip()[:9] for r in rows if (r["现住地址国标"] or "").strip().startswith("6401") and len((r["现住地址国标"] or "").strip())>=9)
    p("  >>> 银川市乡镇级码个数:", len(c9))
    # 人群分类 top
    p("  人群分类 top10:", collections.Counter(r["人群分类"] for r in rows).most_common(10))

with open(os.path.join(base,"_new2_out.txt"),"w",encoding="utf-8") as fh:
    fh.write(out.getvalue())
print("done")
