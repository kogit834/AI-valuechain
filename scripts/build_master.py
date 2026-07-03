"""data/*.csv からバリューチェーン順の企業一覧 output/valuechain_master.xlsx を生成する。"""
import pandas as pd
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
OUTPUT = ROOT / "output"

CONFIDENCE_ORDER = {"高": 0, "中": 1, "低": 2}


def main():
    segments = pd.read_csv(DATA / "segments.csv")
    companies = pd.read_csv(DATA / "companies.csv")
    sources = pd.read_csv(DATA / "sources.csv")

    merged = companies.merge(
        segments[["segment_id", "segment_name", "category", "value_chain_order", "priority"]],
        on="segment_id",
        how="left",
    )

    src_agg = (
        sources.groupby("company_id")
        .apply(lambda g: " / ".join(f"{r.title}（{r.url}）" for r in g.itertuples()))
        .rename("sources_combined")
    )
    merged = merged.merge(src_agg, on="company_id", how="left")

    merged["confidence_rank"] = merged["confidence"].map(CONFIDENCE_ORDER).fillna(9)
    merged = merged.sort_values(
        ["value_chain_order", "confidence_rank", "company_name_ja"]
    ).drop(columns=["confidence_rank"])

    master_cols = [
        "value_chain_order",
        "category",
        "segment_id",
        "segment_name",
        "priority",
        "company_id",
        "company_name_ja",
        "company_name_en",
        "ticker",
        "market",
        "listed_status",
        "country",
        "growth_driver",
        "confidence",
        "sources_combined",
        "added_date",
    ]
    master = merged.reindex(columns=master_cols)

    OUTPUT.mkdir(exist_ok=True)
    out_path = OUTPUT / "valuechain_master.xlsx"
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        master.to_excel(writer, sheet_name="valuechain_master", index=False)
        segments.to_excel(writer, sheet_name="segments", index=False)
        sources.to_excel(writer, sheet_name="sources", index=False)

        ws = writer.sheets["valuechain_master"]
        for col_cells in ws.columns:
            length = max(len(str(c.value)) if c.value is not None else 0 for c in col_cells)
            ws.column_dimensions[col_cells[0].column_letter].width = min(max(length + 2, 10), 60)
        ws.freeze_panes = "A2"

    print(f"Wrote {len(master)} rows to {out_path}")


if __name__ == "__main__":
    main()
