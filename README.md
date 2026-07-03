# AI Infra Value Chain Research

AI需要拡大に伴い歴史的な増収増益が見込める業種・企業を、バリューチェーンに沿って
網羅的に整理するプロジェクト。

## セットアップ
```bash
cd AI-valuechain
pip install pandas openpyxl
claude
```

## 使い方（Claude Code内で）
1. `data/segments.csv` を見て、未着手の優先度「高」セグメントを確認
2. 例: 「SEG004（データセンター建設・ゼネコン）をsegment-researcherで調査して」
3. サブエージェントが research/ にメモを作成し、data/companies.csv・data/sources.csv に追記
4. 全セグメント、または区切りのよいところで:
   ```bash
   python scripts/build_master.py
   ```
   → output/valuechain_master.xlsx が生成される
5. 定期的に見直し・更新（決算発表シーズンごとの更新を推奨）

## ディレクトリ構成
```
AI-valuechain/
├── CLAUDE.md                       # 調査方針・ルール
├── data/
│   ├── segments.csv                # バリューチェーンのセグメントマスタ
│   ├── companies.csv               # 企業マスタ（メインデータ）
│   └── sources.csv                 # 出典マスタ
├── research/                       # セグメント別の調査メモ
├── scripts/build_master.py         # xlsx集計スクリプト
├── output/valuechain_master.xlsx   # 最終成果物
└── .claude/agents/segment-researcher.md  # 調査サブエージェント
```

## 運用ルール
CLAUDE.md参照。特に「一次情報を最低1件」「確度の明記」「思惑先行の情報を安易に含めない」
の3点は精度維持のため厳守。
