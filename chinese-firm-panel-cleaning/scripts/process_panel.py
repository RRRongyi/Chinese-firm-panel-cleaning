import argparse
from pathlib import Path
import pandas as pd

def clean_code(s):
    s = s.astype(str).str.strip().str.replace(r"\\.0$", "", regex=True)
    s = s.str.extract(r"(\\d+)", expand=False).fillna("").str.lstrip("0")
    return s.replace("", "0")

def clean_one(path, header_row, out):
    df = pd.read_excel(path, header=header_row)
    df.columns = (df.columns.astype(str).str.strip().str.replace("\\n", "", regex=False)
                  .str.replace("\\r", "", regex=False).str.replace(" ", "", regex=False))
    if not {"std", "date"}.issubset(df.columns):
        raise ValueError(f"{path.name}必须包含std和date列")
    df["std"] = clean_code(df["std"])
    date_text = (df["date"].astype(str).str.strip().str.replace("年", "-", regex=False)
                 .str.replace("月", "-", regex=False).str.replace("日", "", regex=False)
                 .str.replace("/", "-", regex=False).str.replace(".", "-", regex=False))
    df["date"] = pd.to_datetime(date_text, errors="coerce")
    df = df.dropna(subset=["date"])
    df = df[(df["date"].dt.month == 12) & (df["date"].dt.day == 31)].copy()
    pos = df.columns.get_loc("date")
    df.insert(pos + 1, "year", df["date"].dt.year.astype(int))
    df["yearstd"] = df["year"].astype(str) + df["std"]
    dup = df[df.duplicated("yearstd", keep=False)]
    if not dup.empty:
        dup.to_excel(out / f"{path.stem}_duplicates.xlsx", index=False)
        raise ValueError(f"{path.name}存在重复yearstd")
    df.to_excel(out / f"{path.stem}_clean.xlsx", index=False)
    return df

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--main", required=True)
    ap.add_argument("--others-dir", required=True)
    ap.add_argument("--analysis-vars", nargs="+", required=True)
    ap.add_argument("--winsor-vars", nargs="+", required=True)
    ap.add_argument("--header-row", type=int, default=0)
    ap.add_argument("--output-dir", default="panel_output")
    a = ap.parse_args()
    out = Path(a.output_dir); out.mkdir(parents=True, exist_ok=True)
    main_path = Path(a.main)
    paths = [main_path] + [p for p in Path(a.others_dir).glob("*.xls*") if p.resolve() != main_path.resolve()]
    data = {p.name: clean_one(p, a.header_row, out) for p in paths}
    merged = data[main_path.name].copy(); merge_log = []
    for name, other in data.items():
        if name == main_path.name: continue
        merged = merged.merge(other, on="yearstd", how="left", validate="one_to_one", suffixes=("", f"_{Path(name).stem}"), indicator=True)
        matched = int((merged["_merge"] == "both").sum())
        merge_log.append({"file": name, "matched": matched, "rows": len(merged), "match_rate": matched / len(merged)})
        merged = merged.drop(columns="_merge")
    vars_found = [v for v in a.analysis_vars if v in merged.columns]
    before = len(merged); merged = merged.dropna(subset=vars_found).copy()
    pd.DataFrame([{"before": before, "after": len(merged), "dropped": before - len(merged)}]).to_excel(out / "missing_log.xlsx", index=False)
    win_log = []
    for var in a.winsor_vars:
        if var not in merged.columns: continue
        merged[var] = pd.to_numeric(merged[var], errors="coerce")
        support = set(merged[var].dropna().unique())
        if support and support.issubset({0, 1}):
            win_log.append({"variable": var, "status": "skipped_binary"}); continue
        lo, hi = merged[var].quantile(.01), merged[var].quantile(.99)
        merged[f"{var}_w"] = merged[var].clip(lo, hi)
        win_log.append({"variable": var, "status": "winsorized", "p1": lo, "p99": hi})
    if not merged["yearstd"].is_unique: raise ValueError("最终yearstd不唯一")
    merged.to_excel(out / "final_panel_winsorized.xlsx", index=False)
    pd.DataFrame(merge_log).to_excel(out / "merge_log.xlsx", index=False)
    pd.DataFrame(win_log).to_excel(out / "winsor_log.xlsx", index=False)

if __name__ == "__main__":
    main()
