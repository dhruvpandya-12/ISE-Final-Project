"""
STEP 2 - Analyze FDA medical device recalls to find where prevention effort pays off most.

Input : data/recalls.csv, data/enforcement.csv   (made by step1_extract.py)
Output: outputs/ folder with result tables (CSV), charts (PNG) and results.json
Run   : python step2_analysis.py     (takes under a minute)

Unit of analysis: a recall EVENT. One event is one quality problem (one root cause) that may
cover several recalled products. FDA gives each product its own recall number, so a single kit
recall can produce 200 product records; counting events avoids that distortion.

Study window: events initiated 2013-2025. FDA severity class (I, II, III) is available for
nearly every recall from 2013 on.
"""
import json
import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from scipy import stats
import statsmodels.formula.api as smf

OUT = "outputs"
os.makedirs(OUT, exist_ok=True)
START, END = 2013, 2025          # full years with severity class available
BASE = (2016, 2018)              # baseline: three full years before the COVID disruption
RECENT = (2023, 2025)            # the three most recent full years
results = {}

# ------------------------------------------------------------------ 1. Load and join the two files
rec = pd.read_csv("data/recalls.csv", dtype=str)
enf = pd.read_csv("data/enforcement.csv", dtype=str)
results["records_recall_file"] = len(rec)
results["records_enforcement_file"] = len(enf)

# The recall file has the root cause; the enforcement file has the FDA class.
# Both use the same recall number (e.g. Z-0129-2027), so we join on it.
df = rec.merge(enf[["recall_number", "classification"]], on="recall_number", how="left")
results["records_matched_to_class"] = int(df["classification"].notna().sum())
df["year"] = pd.to_numeric(df["date_initiated"].str[:4], errors="coerce")
df["root_cause"] = df["root_cause"].fillna("Not recorded").str.strip()

# ------------------------------------------------------------------ 2. Group root causes
# FDA uses about 40 detailed root-cause labels. Each is assigned to the quality-system area that
# should have prevented it, following ISO 13485:2016 (adopted by the FDA QMSR, Feb 2, 2026).
GROUPS = {
    "Design (hardware)": ["Device Design", "Component design/selection"],
    "Software": ["Software design", "Software Design Change", "Software change control",
                 "Software Manufacturing/Software Deployment", "Software design (manufacturing process)",
                 "Software in the Use Environment"],
    "Production process": ["Process control", "Process design", "Process change control", "Employee error",
                           "Equipment maintenance", "Environmental control", "Manufacturing material removal",
                           "Reprocessing Controls", "Release of Material/Component prior to receiving test results",
                           "Finished device change control"],
    "Supplier & material": ["Nonconforming Material/Component", "Material/Component Contamination",
                            "Component change control", "Vendor change control",
                            "Mixed-up of materials/components", "Counterfeit"],
    "Labeling & packaging": ["Packaging", "Packaging process control", "Packaging change control",
                             "Package design/selection", "Error in labeling", "Labeling design",
                             "Labeling Change Control", "Labeling mix-ups", "Labeling False and Misleading",
                             "Incorrect or no expiration date", "Storage"],
    "Use error": ["Use error"],
    "Regulatory": ["No Marketing Application", "PMA", "Radiation Control for Health and Safety Act"],
    "Not determined": ["Other", "Under Investigation by firm", "Unknown/Undetermined by firm",
                       "Pending", "Not recorded"],
}
ISO = {"Design (hardware)": "7.3 Design and development",
       "Software": "7.3 Design and development + IEC 62304 software life cycle",
       "Production process": "7.5 Production and service provision",
       "Supplier & material": "7.4 Purchasing",
       "Labeling & packaging": "7.5.1 Control of production (labeling, packaging)",
       "Use error": "7.3 Design and development + IEC 62366 usability",
       "Regulatory": "Regulatory submissions (not a manufacturing failure)",
       "Not determined": "8.5.2 Corrective action (root cause not found)"}
lookup = {cause: g for g, causes in GROUPS.items() for cause in causes}
df["category"] = df["root_cause"].map(lookup)
unmapped = sorted(df.loc[df["category"].isna(), "root_cause"].unique())
assert not unmapped, f"Root causes not assigned to a group: {unmapped}"
pd.DataFrame([(c, g) for g, cs in GROUPS.items() for c in cs], columns=["root_cause", "category"]) \
  .to_csv(f"{OUT}/table_root_cause_groups.csv", index=False)

# ------------------------------------------------------------------ 3. Build the event table
prod = df[(df["year"] >= START) & (df["year"] <= END) &
          df["classification"].isin(["Class I", "Class II", "Class III"])].copy()
prod["class1"] = (prod["classification"] == "Class I").astype(int)
prod["dc"] = pd.to_numeric(prod["device_class"], errors="coerce")   # FDA device class 1, 2 or 3 (other codes -> blank)
results["study_product_records"] = len(prod)

def mode_or_unknown(x):
    x = x.dropna()
    return x.mode().iat[0] if len(x) else "Unknown"

ev = prod.groupby("event_id").agg(
    year=("year", "min"),
    date=("date_initiated", "min"),
    firm=("recalling_firm", "first"),
    root_cause=("root_cause", mode_or_unknown),
    category=("category", mode_or_unknown),
    specialty=("medical_specialty", mode_or_unknown),
    products=("recall_number", "size"),
    class1=("class1", "max"),          # event is Class I if any of its products is Class I
    dc=("dc", "max"),                  # highest device class among the event's products
).reset_index()
ev["year"] = ev["year"].astype(int)
ev["devclass"] = ev["dc"].map({1: "Class 1", 2: "Class 2", 3: "Class 3"}).fillna("Unknown")
ev = ev.drop(columns="dc")
ev["period"] = pd.cut(ev["year"], [2012, 2018, 2022, 2025], labels=["2013-2018", "2019-2022", "2023-2025"]).astype(str)
SMALL = 200                            # specialties with fewer known-cause events are pooled for the regression
results["study_events"] = len(ev)
results["class1_events"] = int(ev["class1"].sum())
results["events_with_more_than_one_root_cause"] = int((prod.groupby("event_id")["root_cause"].nunique() > 1).sum())

# ------------------------------------------------------------------ 4. The problem: trend by year
by_year = ev.groupby("year").agg(events=("event_id", "size"), class1=("class1", "sum")).reset_index()
by_year["class1_share"] = by_year["class1"] / by_year["events"]
by_year.to_csv(f"{OUT}/table_by_year.csv", index=False)
y = by_year.set_index("year")
results["by_year"] = by_year.round(4).to_dict(orient="records")
results["class1_change_2020_to_2025_pct"] = round(float((y.loc[2025, "class1"] / y.loc[2020, "class1"] - 1) * 100), 1)
results["events_change_2020_to_2025_pct"] = round(float((y.loc[2025, "events"] / y.loc[2020, "events"] - 1) * 100), 1)

# Is the rise in the Class I share real? Two-proportion z-test, baseline vs recent period
def window(frame, w):
    return frame[(frame["year"] >= w[0]) & (frame["year"] <= w[1])]

b, r = window(ev, BASE), window(ev, RECENT)
x1, n1, x2, n2 = b["class1"].sum(), len(b), r["class1"].sum(), len(r)
p_pool = (x1 + x2) / (n1 + n2)
z = (x2 / n2 - x1 / n1) / np.sqrt(p_pool * (1 - p_pool) * (1 / n1 + 1 / n2))
results["share_test"] = {"baseline_share": round(float(x1 / n1), 4), "recent_share": round(float(x2 / n2), 4),
                         "baseline_n": int(n1), "recent_n": int(n2), "z": round(float(z), 2),
                         "p_value": float(2 * (1 - stats.norm.cdf(abs(z))))}

# ------------------------------------------------------------------ 5. Pareto of known root causes
results["share_not_determined"] = round(float((ev["category"] == "Not determined").mean()), 4)
results["share_not_determined_recent"] = round(float((r["category"] == "Not determined").mean()), 4)
known = ev[ev["category"] != "Not determined"].copy()
results["events_known_cause"] = len(known)

cat = (known.groupby("category").agg(events=("event_id", "size"), class1=("class1", "sum"))
       .sort_values("events", ascending=False))
cat["share"] = cat["events"] / cat["events"].sum()
cat["cum_share"] = cat["share"].cumsum()
cat["class1_rate"] = cat["class1"] / cat["events"]
cat["class1_share"] = cat["class1"] / cat["class1"].sum()

detail = (known.groupby(["root_cause", "category"]).agg(events=("event_id", "size"), class1=("class1", "sum"))
          .reset_index().sort_values("events", ascending=False))
detail["share"] = detail["events"] / detail["events"].sum()
detail["cum_share"] = detail["share"].cumsum()
detail["class1_rate"] = detail["class1"] / detail["events"]
detail.to_csv(f"{OUT}/table_pareto_detailed_causes.csv", index=False)

# ------------------------------------------------------------------ 6. Severity: chi-square test
ct = pd.crosstab(known["category"], known["class1"])
chi2, p, dof, expected = stats.chi2_contingency(ct)
results["chi_square"] = {"chi2": round(float(chi2), 1), "dof": int(dof), "p_value": float(p),
                         "cramers_v": round(float(np.sqrt(chi2 / (ct.values.sum() * (min(ct.shape) - 1)))), 3),
                         "min_expected_count": round(float(expected.min()), 1)}
overall_rate = known["class1"].mean()
results["class1_rate_known_causes"] = round(float(overall_rate), 4)
cat["relative_risk"] = cat["class1_rate"] / overall_rate
cat["rate_ci95"] = 1.96 * np.sqrt(cat["class1_rate"] * (1 - cat["class1_rate"]) / cat["events"])

# ------------------------------------------------------------------ 7. Trend by category
for name, w in (("base", BASE), ("recent", RECENT)):
    sub = window(known, w)
    yrs = w[1] - w[0] + 1
    cat[f"events_per_year_{name}"] = sub.groupby("category").size() / yrs
    cat[f"class1_per_year_{name}"] = sub.groupby("category")["class1"].sum() / yrs
    cat[f"class1_rate_{name}"] = sub.groupby("category")["class1"].mean()
cat = cat.fillna(0)
cat["class1_growth"] = np.where(cat["class1_per_year_base"] > 0,
                                cat["class1_per_year_recent"] / cat["class1_per_year_base"] - 1, np.nan)

# ------------------------------------------------------------------ 8. Risk priority number (FMEA style)
# Scored for the six manufacturing-quality categories (Regulatory is a paperwork issue, not a
# product failure, so it is reported but not ranked).
# O = occurrence: recall events per year, 2023-2025
# S = severity:   share of those events that were Class I, 2023-2025
# T = trend:      growth in Class I events per year, 2023-2025 vs 2016-2018
# Each factor is rescaled to 1-10 across the six categories: score = 1 + 9 * (x - min) / (max - min)
rank_set = cat.index != "Regulatory"

def scale(x):
    return 1 + 9 * (x - x.min()) / (x.max() - x.min())

cat.loc[rank_set, "O_score"] = scale(cat.loc[rank_set, "events_per_year_recent"])
cat.loc[rank_set, "S_score"] = scale(cat.loc[rank_set, "class1_rate_recent"])
cat.loc[rank_set, "T_score"] = scale(cat.loc[rank_set, "class1_growth"])
cat["RPN"] = cat["O_score"] * cat["S_score"] * cat["T_score"]
cat["priority_rank"] = cat["RPN"].rank(ascending=False)
cat["iso_13485_clause"] = [ISO[c] for c in cat.index]
cat = cat.reset_index()
cat.to_csv(f"{OUT}/table_categories.csv", index=False)
results["categories"] = cat.round(4).to_dict(orient="records")

# sensitivity: does the top-3 change if one factor is left out?
sens = {}
for drop in ("O_score", "S_score", "T_score"):
    keep = [c for c in ("O_score", "S_score", "T_score") if c != drop]
    sc = cat.loc[cat["category"] != "Regulatory"].assign(score=lambda d: d[keep[0]] * d[keep[1]])
    sens[f"without_{drop[0]}"] = sc.sort_values("score", ascending=False)["category"].head(3).tolist()
results["rpn_sensitivity_top3"] = sens

# ------------------------------------------------------------------ 9. p-chart: Class I share by quarter
ev["quarter"] = pd.PeriodIndex(pd.to_datetime(ev["date"], errors="coerce"), freq="Q")
q = ev.groupby("quarter").agg(n=("class1", "size"), d=("class1", "sum"))
# Center line and limits come from the baseline years 2013-2018 (Phase I), then later quarters
# are judged against them (Phase II), the standard way to set up a control chart.
phase1 = q[q.index.year <= 2018]
pbar = phase1["d"].sum() / phase1["n"].sum()
q["p"] = q["d"] / q["n"]
q["UCL"] = pbar + 3 * np.sqrt(pbar * (1 - pbar) / q["n"])
q["LCL"] = (pbar - 3 * np.sqrt(pbar * (1 - pbar) / q["n"])).clip(lower=0)
q["above_UCL"] = q["p"] > q["UCL"]
q.index = q.index.astype(str)
q.to_csv(f"{OUT}/table_pchart_quarterly.csv")
results["pchart"] = {"p_bar_2013_2018": round(float(pbar), 4), "quarters": len(q),
                     "quarters_above_UCL": q.index[q["above_UCL"]].tolist(),
                     "quarters_above_UCL_2023_2025": int(q.loc[q.index >= "2023Q1", "above_UCL"].sum())}

# ------------------------------------------------------------------ 10. Imaging / radiology focus
rad = known[known["specialty"] == "Radiology"]
rad_tbl = (rad.groupby("category").agg(events=("event_id", "size"), class1=("class1", "sum"))
           .sort_values("events", ascending=False))
rad_tbl["share"] = rad_tbl["events"] / rad_tbl["events"].sum()
rad_tbl["cum_share"] = rad_tbl["share"].cumsum()
rad_tbl["class1_rate"] = rad_tbl["class1"] / rad_tbl["events"]
rad_tbl = rad_tbl.reset_index()
rad_tbl.to_csv(f"{OUT}/table_radiology.csv", index=False)
results["radiology"] = {"events_known_cause": len(rad), "table": rad_tbl.round(4).to_dict(orient="records")}

# ------------------------------------------------------------------ 11. Impact scenario
ranked = cat[cat["category"] != "Regulatory"].sort_values("RPN", ascending=False)
top2 = ranked["category"].head(2).tolist()
c1_recent_total = cat["class1_per_year_recent"].sum()
c1_top2 = cat.loc[cat["category"].isin(top2), "class1_per_year_recent"].sum()
scen = []
for cut in (0.10, 0.25, 0.50):
    scen.append({"reduction": cut, "class1_avoided_per_year": round(float(c1_top2 * cut), 1),
                 "share_of_known_cause_class1": round(float(c1_top2 * cut / c1_recent_total), 4)})
results["impact"] = {"top_priority_categories": top2,
                     "class1_per_year_recent_known_causes": round(float(c1_recent_total), 1),
                     "class1_per_year_recent_top2": round(float(c1_top2), 1),
                     "top2_share_of_class1": round(float(c1_top2 / c1_recent_total), 4),
                     "scenarios": scen}

# ------------------------------------------------------------------ 12. 2026 year to date (context)
ytd = df[df["year"] == 2026]
results["ytd_2026"] = {"product_records": len(ytd), "events": int(ytd["event_id"].nunique()),
                       "class1_events_so_far": int(ytd.loc[ytd["classification"] == "Class I", "event_id"].nunique()),
                       "latest_initiation_date": str(ytd["date_initiated"].max()),
                       "note": "Many 2026 recalls are not yet classified, so counts will rise."}


# ------------------------------------------------------------------ 12b. Logistic regression
# Question: is a root cause still linked to Class I once device class, medical specialty and period
# are taken into account? (A design recall might look severe only because design problems are more
# common in high-risk devices. The regression separates those effects.)
# Model: log(p / (1 - p)) = b0 + b(root cause) + b(device class) + b(specialty) + b(period)
# where p = probability the event is Class I. exp(b) = odds ratio versus the reference level.
REF = {"category": "Production process", "devclass": "Class 2", "spec_model": "General Hospital",
       "period": "2013-2018"}
k = known.copy()                       # known-cause events; devclass and period were added in section 3
spec_n = k["specialty"].value_counts()
k["spec_model"] = k["specialty"].where(k["specialty"].map(spec_n) >= SMALL, "Other specialties")
ev["spec_model"] = ev["specialty"].where(ev["specialty"].map(spec_n).fillna(0) >= SMALL, "Other specialties")
TERMS = [("category", "Root cause"), ("devclass", "Device class"), ("spec_model", "Specialty"), ("period", "Period")]
formula = "class1 ~ " + " + ".join(f"C({v}, Treatment('{REF[v]}'))" for v, _ in TERMS)
logit = smf.logit(formula, data=k).fit(disp=0, maxiter=200)
assert logit.mle_retvals["converged"], "logistic regression did not converge"
robust = smf.logit(formula, data=k).fit(disp=0, maxiter=200, cov_type="cluster",
                                       cov_kwds={"groups": pd.factorize(k["firm"])[0]})
null = smf.logit("class1 ~ 1", data=k).fit(disp=0)
no_cause = smf.logit(formula.replace(f"C(category, Treatment('{REF['category']}')) + ", ""), data=k).fit(disp=0, maxiter=200)

rows = [{"variable": "Intercept", "level": "(baseline profile)", "coef": logit.params["Intercept"],
         "se": logit.bse["Intercept"], "events": len(k), "class1": int(k["class1"].sum())}]
for v, label in TERMS:
    levels = k[v].value_counts()
    for lev in [REF[v]] + sorted(l for l in levels.index if l != REF[v]):
        name = f"C({v}, Treatment('{REF[v]}'))[T.{lev}]"
        rows.append({"variable": label, "level": lev + (" (reference)" if lev == REF[v] else ""),
                     "coef": logit.params.get(name, 0.0), "se": logit.bse.get(name, np.nan),
                     "se_firm_clustered": robust.bse.get(name, np.nan),
                     "events": int(levels[lev]), "class1": int(k.loc[k[v] == lev, "class1"].sum())})
lt = pd.DataFrame(rows)
lt["z"] = lt["coef"] / lt["se"]
lt["p_value"] = 2 * (1 - stats.norm.cdf(lt["z"].abs()))
lt["p_value_firm_clustered"] = 2 * (1 - stats.norm.cdf((lt["coef"] / lt["se_firm_clustered"]).abs()))
lt["odds_ratio"] = np.exp(lt["coef"])
lt["or_low95"] = np.exp(lt["coef"] - 1.96 * lt["se"])
lt["or_high95"] = np.exp(lt["coef"] + 1.96 * lt["se"])
lt.loc[lt["variable"] == "Intercept", ["odds_ratio", "or_low95", "or_high95"]] = np.nan
lt.to_csv(f"{OUT}/table_logistic_regression.csv", index=False)

# Unadjusted odds ratios for root cause (no other variables), to show what changes after adjustment
base_odds = k.loc[k["category"] == REF["category"], "class1"].mean()
base_odds = base_odds / (1 - base_odds)
unadj = {}
for c_ in cat["category"]:
    r_ = k.loc[k["category"] == c_, "class1"].mean()
    unadj[c_] = round(float(r_ / (1 - r_) / base_odds), 3)

# Model fit: ROC AUC via the rank (Mann-Whitney) formula
k["p_hat"] = logit.predict(k)
ranks = stats.rankdata(k["p_hat"])
n1_, n0_ = int(k["class1"].sum()), int((1 - k["class1"]).sum())
auc = (ranks[k["class1"].values == 1].sum() - n1_ * (n1_ + 1) / 2) / (n1_ * n0_)
lr_all = 2 * (logit.llf - null.llf)
lr_cause = 2 * (logit.llf - no_cause.llf)
df_cause = int(logit.df_model - no_cause.df_model)

def prob(**kw):
    row = pd.DataFrame([{**REF, **kw}])
    return round(float(logit.predict(row).iat[0]), 4)

results["logistic_regression"] = {
    "events": len(k), "class1": n1_, "parameters": int(logit.df_model) + 1,
    "reference": REF, "small_specialty_threshold": SMALL,
    "pooled_specialties": sorted(spec_n[spec_n < SMALL].index.tolist()),
    "log_likelihood": round(float(logit.llf), 3), "log_likelihood_null": round(float(null.llf), 3),
    "lr_test_all": {"chi2": round(float(lr_all), 1), "df": int(logit.df_model),
                    "p_value": float(stats.chi2.sf(lr_all, logit.df_model))},
    "lr_test_root_cause": {"chi2": round(float(lr_cause), 1), "df": df_cause,
                           "p_value": float(stats.chi2.sf(lr_cause, df_cause))},
    "mcfadden_r2": round(float(1 - logit.llf / null.llf), 3), "auc": round(float(auc), 3),
    "unadjusted_or_root_cause": unadj,
    "example_probabilities": {
        "production_class2_general_hospital_2023_2025": prob(period="2023-2025"),
        "design_class2_general_hospital_2023_2025": prob(category="Design (hardware)", period="2023-2025"),
        "production_class2_general_hospital_2013_2018": prob(),
    },
    "table": lt.drop(columns=[]).round(4).replace({np.nan: None}).to_dict(orient="records"),
}

ev.to_csv(f"{OUT}/study_events.csv", index=False)   # saved here so it includes the regression columns

# ------------------------------------------------------------------ 13. Charts
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 10,
                     "axes.spines.top": False, "axes.spines.right": False})
GREY, BLUE, RED = "#9AA0A6", "#2A6FDB", "#C00000"

# Chart 1: the problem
fig, ax = plt.subplots(figsize=(9, 4.2))
ax.bar(by_year["year"], by_year["class1"], color=[BLUE if yy >= 2021 else GREY for yy in by_year["year"]])
for yy, v, s_ in zip(by_year["year"], by_year["class1"], by_year["class1_share"]):
    ax.text(yy, v + 2, f"{v}\n({s_:.0%})", ha="center", fontsize=8)
ax.set_xticks(by_year["year"]); ax.set_ylim(0, by_year["class1"].max() * 1.25)
ax.set_ylabel("Class I recall events")
ax.set_title("Class I recall events per year (share of all recall events in brackets)", loc="left", fontweight="bold")
fig.tight_layout(); fig.savefig(f"{OUT}/chart1_class1_by_year.png", dpi=200); plt.close(fig)

# Chart 2: Pareto
fig, ax = plt.subplots(figsize=(9, 4.5))
xi = np.arange(len(cat))
ax.bar(xi, cat["events"], color=GREY)
ax.set_xticks(xi); ax.set_xticklabels(cat["category"], rotation=25, ha="right")
ax.set_ylabel("Recall events, 2013-2025")
ax2 = ax.twinx(); ax2.plot(xi, cat["cum_share"] * 100, color="#333", marker="o", lw=1.5)
ax2.axhline(80, color=RED, ls="--", lw=1); ax2.set_ylim(0, 105); ax2.set_ylabel("Cumulative %")
ax2.spines["top"].set_visible(False)
for i, v in enumerate(cat["cum_share"]):
    ax2.text(i, v * 100 + 3, f"{v:.0%}", ha="center", fontsize=8)
ax.set_title("Pareto of recall root causes (events with a known cause)", loc="left", fontweight="bold")
fig.tight_layout(); fig.savefig(f"{OUT}/chart2_pareto.png", dpi=200); plt.close(fig)

# Chart 3: severity
s_ = cat.sort_values("class1_rate")
fig, ax = plt.subplots(figsize=(9, 4.5))
ax.barh(s_["category"], s_["class1_rate"] * 100, xerr=s_["rate_ci95"] * 100, capsize=3,
        color=[BLUE if v > overall_rate else GREY for v in s_["class1_rate"]])
ax.axvline(overall_rate * 100, color=RED, ls="--", lw=1)
ax.text(overall_rate * 100 + 0.2, len(s_) - 0.45, f"average {overall_rate:.1%}", color=RED, fontsize=8)
for i, v in enumerate(s_["class1_rate"]):
    ax.text(v * 100 + 0.3, i + 0.18, f"{v:.1%}", fontsize=8)
ax.set_xlabel("Share of recall events classified Class I, 2013-2025 (with 95% confidence interval)")
ax.set_title("Which root causes lead to the most serious recalls?", loc="left", fontweight="bold")
fig.tight_layout(); fig.savefig(f"{OUT}/chart3_severity.png", dpi=200); plt.close(fig)

# Chart 4: p-chart
fig, ax = plt.subplots(figsize=(10, 4))
xs = np.arange(len(q))
ax.plot(xs, q["p"] * 100, marker="o", ms=3, lw=1.1, color=BLUE, label="Class I share of events")
ax.step(xs, q["UCL"] * 100, where="mid", color=RED, ls="--", lw=1, label="Control limits (from 2013-2018)")
ax.step(xs, q["LCL"] * 100, where="mid", color=RED, ls="--", lw=1)
ax.axhline(pbar * 100, color="#333", lw=1, label=f"Center line {pbar:.1%}")
oi = np.where(q["above_UCL"])[0]
ax.scatter(oi, q["p"].iloc[oi] * 100, color=RED, zorder=5, s=28, label="Above UCL")
ticks = [i for i, lab in enumerate(q.index) if lab.endswith("Q1")]
ax.set_xticks(ticks); ax.set_xticklabels([q.index[i][:4] for i in ticks])
ax.axvline(ticks[q.index[ticks].tolist().index("2019Q1")] - 0.5, color=GREY, lw=0.8)
ax.set_ylabel("% of recall events that are Class I")
ax.set_title("p-chart: share of recall events classified Class I, by quarter", loc="left", fontweight="bold")
ax.legend(frameon=False, fontsize=8, ncol=2, loc="upper left")
fig.tight_layout(); fig.savefig(f"{OUT}/chart4_pchart.png", dpi=200); plt.close(fig)

# Chart 5: priority
s_ = cat[cat["category"] != "Regulatory"].sort_values("RPN")
fig, ax = plt.subplots(figsize=(9, 4))
ax.barh(s_["category"], s_["RPN"], color=[BLUE if r_ <= 3 else GREY for r_ in s_["priority_rank"]])
for i, v in enumerate(s_["RPN"]):
    ax.text(v + 3, i - 0.1, f"{v:.0f}", fontsize=8)
ax.set_xlabel("Risk priority number = Occurrence x Severity x Trend (each scored 1-10)")
ax.set_title("Where manufacturers should focus recall prevention first", loc="left", fontweight="bold")
fig.tight_layout(); fig.savefig(f"{OUT}/chart5_priority.png", dpi=200); plt.close(fig)

# Chart 6: logistic regression odds ratios (root cause and period, adjusted)
sel = lt[lt["variable"].isin(["Root cause", "Period"]) & ~lt["level"].str.contains("reference")].iloc[::-1]
fig, ax = plt.subplots(figsize=(9, 4.5))
yy_ = np.arange(len(sel))
ax.errorbar(sel["odds_ratio"], yy_, xerr=[sel["odds_ratio"] - sel["or_low95"], sel["or_high95"] - sel["odds_ratio"]],
            fmt="o", color=BLUE, ecolor=GREY, capsize=3)
ax.axvline(1, color=RED, ls="--", lw=1)
ax.set_xscale("log"); ax.set_yticks(yy_)
ticks_ = [0.2, 0.3, 0.5, 1, 2, 3]
ax.set_xticks(ticks_); ax.set_xticklabels([str(t) for t in ticks_]); ax.minorticks_off()
ax.set_yticklabels([f"{v}: {l}" for v, l in zip(sel["variable"], sel["level"])])
for i, v in enumerate(sel["odds_ratio"]):
    ax.text(v, i + 0.22, f"{v:.2f}", ha="center", fontsize=8)
ax.set_xlabel("Odds ratio of a Class I recall (log scale)\n1 = same as reference (Production process; 2013-2018). Bars = 95% CI")
ax.set_title("Logistic regression: adjusted for device class and specialty", loc="left", fontweight="bold")
fig.tight_layout(); fig.savefig(f"{OUT}/chart6_logistic_odds_ratios.png", dpi=200); plt.close(fig)

with open(f"{OUT}/results.json", "w") as f:
    json.dump(results, f, indent=2, default=lambda o: o.item() if hasattr(o, "item") else str(o))

# ------------------------------------------------------------------ console summary
print(f"Recall events studied (2013-2025): {results['study_events']:,} "
      f"({results['study_product_records']:,} recalled products)")
print(f"Class I events 2020 -> 2025: {y.loc[2020, 'class1']} -> {y.loc[2025, 'class1']} "
      f"({results['class1_change_2020_to_2025_pct']:+.0f}%)")
print(f"Class I share {results['share_test']['baseline_share']:.1%} (2016-18) -> "
      f"{results['share_test']['recent_share']:.1%} (2023-25), z = {results['share_test']['z']}")
print(f"Root cause not determined: {results['share_not_determined']:.1%} of events")
print(f"Chi-square (cause vs Class I): {results['chi_square']}")
print(cat[["category", "events", "share", "cum_share", "class1_rate", "class1_rate_recent", "class1_growth",
           "RPN", "priority_rank"]].round(3).to_string(index=False))
print("Top priorities:", top2, "| sensitivity:", sens)
print("p-chart quarters above UCL:", results["pchart"]["quarters_above_UCL"])
lr = results["logistic_regression"]
print(f"Logistic regression: {lr['events']:,} events, McFadden R2 = {lr['mcfadden_r2']}, AUC = {lr['auc']}, "
      f"root-cause LR test chi2 = {lr['lr_test_root_cause']['chi2']} (df {lr['lr_test_root_cause']['df']})")
print(lt[["variable", "level", "odds_ratio", "or_low95", "or_high95", "p_value"]].round(3).to_string(index=False))
print("Done. Results are in the outputs folder.")
