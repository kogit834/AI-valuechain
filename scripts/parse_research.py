"""research/NN_*.md (segment-researcher出力フォーマット) を data/companies.csv, data/sources.csv に追記する。

使い方: python scripts/parse_research.py research/02_heavy_electric_generators.md SEG002
"""
import csv
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"


def max_id_num(path, prefix, col):
    with open(path, newline="", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    nums = [int(r[col][1:]) for r in rows if r[col].startswith(prefix)]
    return max(nums) if nums else 0


def parse_company_header(line):
    # "### 企業名(EN)（上場/非上場、市場・ティッカー、国）"
    text = line[4:].strip()
    m = re.match(r"^(.*)（([^（）]*)）\s*$", text)
    if not m:
        return text, {}
    name_part, meta_part = m.group(1).strip(), m.group(2).strip()
    fields = [x.strip() for x in meta_part.split("、")]
    listed_status = fields[0] if len(fields) > 0 else ""
    market_ticker = fields[1] if len(fields) > 1 else ""
    country = fields[2] if len(fields) > 2 else ""
    return name_part, {
        "listed_status": listed_status,
        "market_ticker": market_ticker,
        "country": country,
    }


def parse_file(path):
    text = Path(path).read_text(encoding="utf-8")
    lines = text.splitlines()
    companies = []
    cur = None
    for line in lines:
        if line.startswith("### "):
            if cur:
                companies.append(cur)
            name, meta = parse_company_header(line)
            cur = {"name_raw": name, "meta": meta, "growth_driver": "", "confidence": "", "sources": []}
        elif cur is not None:
            if line.strip().startswith("- 成長ドライバー:") or line.strip().startswith("- 成長ドライバー："):
                cur["growth_driver"] = line.split(":", 1)[-1].split("：", 1)[-1].strip()
            elif line.strip().startswith("- 確度:") or line.strip().startswith("- 確度："):
                cur["confidence"] = line.split(":", 1)[-1].split("：", 1)[-1].strip()
            elif line.strip().startswith("- 出典:"):
                m = re.search(r"\[(.*?)\]\((.*?)\)\s*(.*)", line)
                if m:
                    cur["sources"].append({"title": m.group(1), "url": m.group(2), "date": m.group(3).strip()})
    if cur:
        companies.append(cur)
    return companies


def main():
    md_path, segment_id = sys.argv[1], sys.argv[2]
    companies = parse_file(md_path)

    companies_csv = DATA / "companies.csv"
    sources_csv = DATA / "sources.csv"

    next_cnum = max_id_num(companies_csv, "C", "company_id") + 1
    with open(companies_csv, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        for c in companies:
            cid = "C" + str(next_cnum).zfill(3)
            next_cnum += 1
            name = re.sub(r"\([^)]*\)", "", c["name_raw"]).strip()
            en_match = re.search(r"\(([^)]*)\)", c["name_raw"])
            name_en = en_match.group(1) if en_match else ""
            market_ticker = c["meta"].get("market_ticker", "")
            parts = market_ticker.rsplit(" ", 1) if " " in market_ticker else [market_ticker, ""]
            market = parts[0] if len(parts) > 0 else ""
            ticker = parts[1] if len(parts) > 1 else ""
            writer.writerow([
                cid, segment_id, name, name_en, ticker, market,
                c["meta"].get("listed_status", ""), c["meta"].get("country", ""),
                c["growth_driver"], c["confidence"], "", "2026-07-03",
            ])
            c["_id"] = cid

    # second pass for sources, now that company_ids assigned in same order
    next_snum = max_id_num(sources_csv, "S", "source_id") + 1
    with open(sources_csv, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        for c in companies:
            for s in c["sources"]:
                sid = "S" + str(next_snum).zfill(3)
                next_snum += 1
                writer.writerow([
                    sid, c["_id"], "", s["title"], "", s["url"], s["date"], "2026-07-03",
                ])

    print(f"Added {len(companies)} companies from {md_path} ({segment_id})")


if __name__ == "__main__":
    main()
