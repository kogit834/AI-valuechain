"""国内企業を追加登録する（2026-07 リサーチ分）。
companies.csv / sources.csv に追記する。既存行は変更しない。
一次情報（決算・IR・中計・報道）で裏付けの取れた国内企業を、国内が薄い
セグメントを中心に追加。確度は裏付けの強さに応じて 高/中/低 を付与。
"""
import csv
from pathlib import Path

DATA = Path(__file__).resolve().parent.parent / "data"
TODAY = "2026-07-07"

# (segment_id, name_ja, name_en, ticker, market, listed_status, growth_driver, confidence,
#  source_url, source_title, source_publisher, source_pubdate)  ※source_url が空なら出典未添付
ADDS = [
    # --- SEG001 特別高圧・大型変圧器 ---
    ("SEG001", "愛知電機", "Aichi Electric Co., Ltd.", "6623", "東証スタンダード", "上場",
     "柱上変圧器で国内大手（中部電力系）。生成AI向けパッケージ基板用コアが急伸し26/3期3Qは営業利益+23.0%・通期上方修正。変圧器も電力インフラ投資で堅調。",
     "高", "https://biz.chunichi.co.jp/news/article/10/124732/",
     "愛知電機は増収増益 26年3月期", "中日BIZナビ", "2026"),

    # --- SEG002 発電機・タービン・非常用電源 ---
    ("SEG002", "澤藤電機", "Sawafuji Electric Co., Ltd.", "6901", "東証スタンダード", "上場",
     "エンジン発電機・可搬型電源を手がける。データセンター建設現場・非常用/仮設電源の需要を取り込む。26/3期決算を公表。",
     "低", "https://www.sawafuji.co.jp/",
     "2026年3月期 決算短信を公表", "澤藤電機", "2026"),
    ("SEG002", "北越工業", "Hokuetsu Industries Co., Ltd.", "6364", "東証スタンダード", "上場",
     "可搬型エンジン発電機・エンジンコンプレッサ（AIRMAN）大手。データセンター建設・非常用電源向けの需要が追い風。",
     "低", "", "", "", ""),

    # --- SEG003 送配電（開閉装置・遮断器・地中線） ---
    ("SEG003", "東光高岳", "Takaoka Toko Co., Ltd.", "6617", "東証プライム", "上場",
     "特別高圧受変電機器（変圧器・開閉装置）がデータセンター向けで好調。26/3期は営業利益97.6億円と統合以降で過去最高。小山事業所の生産能力を約1.5倍へ（新工場2030年代初頭）。",
     "高", "https://finance.logmi.jp/articles/384901",
     "東光高岳、営業利益60.2％増で統合以降最高益", "ログミーファイナンス", "2026"),
    ("SEG003", "戸上電機製作所", "Togami Electric Mfg. Co., Ltd.", "6643", "東証スタンダード", "上場",
     "産業用配電機器が好調で26/3期は増収増益（売上307億円+11.2%・営業益+11.2%）。データセンター・工場向け受変電/配電設備の需要を取り込む。",
     "中", "https://finance.yahoo.co.jp/quote/6643.T/financials",
     "戸上電機製作所 決算情報", "Yahoo!ファイナンス", "2026"),

    # --- SEG006 産業冷却・空調（液冷・チラー・CRAC/CRAH） ---
    ("SEG006", "ニデック", "Nidec Corporation", "6594", "東証プライム", "上場",
     "データセンターの液冷（CDU・水冷モジュール）に本格参入。生成AIサーバの高発熱化で液冷需要が急拡大し冷却を成長領域に位置づけ。",
     "中", "https://www.sbbit.jp/article/cont1/159638",
     "データセンターの液浸冷却とは？富士通・ニデックら開発競争", "ビジネス+IT", "2026"),
    ("SEG006", "オリオン機械", "Orion Machinery Co., Ltd.", "6417", "東証スタンダード", "上場",
     "精密チラー大手。AI液冷向けにCDU一次側のチルド水供給設備を40ftコンテナにパッケージ化するなどデータセンター向けを強化。",
     "中", "https://www.orionkikai.co.jp/industries/data-centers/",
     "データセンター向け | 業界別製品情報", "オリオン機械", "2026"),
    ("SEG006", "木村工機", "Kimura Kohki Co., Ltd.", "6231", "東証スタンダード", "上場",
     "データセンター向け中温冷水・外気冷房空調機を展開。省エネ空調でDC需要を取り込む。",
     "中", "", "", "", ""),
    ("SEG006", "前川製作所", "Mayekawa Mfg. Co., Ltd.", "", "", "非上場",
     "産業用冷凍・大型冷却の国内大手（非上場）。データセンターの高効率冷却・ヒートポンプ技術を展開。",
     "低", "", "", "", ""),

    # --- SEG008 半導体前工程（ロジック・メモリ・デバイス） ---
    ("SEG008", "ソニーグループ", "Sony Group Corporation", "6758", "東証プライム", "上場",
     "CMOSイメージセンサ世界首位（シェア約49.5%）。設備投資を増額し先端プロセスの微細化を前倒し。生成AI・車載向けでファブ投資が拡大。",
     "中", "https://news.mynavi.jp/techplus/article/20250616-3353314/",
     "ソニーがCMOSイメージセンサ向け設備投資を増額", "マイナビニュース", "2025"),
    ("SEG008", "ルネサスエレクトロニクス", "Renesas Electronics Corp.", "6723", "東証プライム", "上場",
     "AIサーバ・DC電源向けを含むパワー半導体に投資（900億円規模）。三菱電機・東芝とパワー半導体事業の統合協議も。※足元の市況は軟調で確度は限定的。",
     "低", "https://www.mitsubishielectric.co.jp/ja/pr/2026/pdf/0327.pdf",
     "パワー半導体事業の統合に向けた基本合意", "三菱電機", "2026"),
    ("SEG008", "ローム", "ROHM Co., Ltd.", "6963", "東証プライム", "上場",
     "SiCパワー半導体に7年累計5,100億円を投資し福岡・宮崎に新棟。AI・DC電源の高効率化需要が中長期の柱。※足元の市況は軟調。",
     "低", "https://www.nikkei.com/article/DGXZQOUF078JS0X00C22A6000000/",
     "ローム、次世代パワー半導体の投資3倍 福岡に新工場棟", "日本経済新聞", "2022"),

    # --- SEG010 半導体製造装置・部材 ---
    ("SEG010", "TOWA", "TOWA Corporation", "6315", "東証プライム", "上場",
     "半導体モールディング装置大手。AI・データセンター向け半導体需要増で27/3期は純利益+52%（70億円）見込み。",
     "高", "https://www.nikkei.com/article/DGXZQOUF1169U0R10C26A5000000/",
     "TOWAの27年3月期、純利益5割増 半導体装置がAI向け堅調", "日本経済新聞", "2026"),
    ("SEG010", "キヤノン", "Canon Inc.", "7751", "東証プライム", "上場",
     "半導体露光装置で後工程向け業界標準機を展開し販売台数を伸ばす。生成AI・先端パッケージング需要で露光装置事業を拡大・生産能力を増強。",
     "中", "https://news.mynavi.jp/techplus/article/20260216-4131591/",
     "半導体露光機3社の決算まとめ", "マイナビニュース", "2026"),
    ("SEG010", "東京精密", "Tokyo Seimitsu Co., Ltd.", "7729", "東証プライム", "上場",
     "半導体プローバ・ダイシング装置大手。AI半導体（HBM・先端パッケージ）の検査・加工需要増を取り込む。",
     "中", "", "", "", ""),
    ("SEG010", "コクサイエレクトリック", "Kokusai Electric Corp.", "6525", "東証プライム", "上場",
     "バッチ式成膜（ALD/CVD）装置大手。AIメモリ（DRAM/HBM）増産で成膜装置需要が拡大。",
     "中", "", "", "", ""),

    # --- SEG011 電力インフラ（発送電・系統増強） ---
    ("SEG011", "電源開発（Jパワー）", "J-POWER (Electric Power Development)", "9513", "東証プライム", "上場",
     "2024-26中計で3年約3兆円（2030年度まで7兆円）を投資。脱炭素電源・送電網増強を推進し、AI・DCによる電力需要増を見据える。",
     "中", "https://www.jpower.co.jp/ir/management/plan/",
     "中期経営計画", "J-POWER", "2024"),
    ("SEG011", "東北電力", "Tohoku Electric Power Co., Inc.", "9506", "東証プライム", "上場",
     "中長期ビジョン『よりそうnext+PLUS』で送配電・再エネ投資を拡大。半導体・DC立地に伴う需要増に対応。",
     "中", "https://www.tohoku-epco.co.jp/comp/keiei/vision.html",
     "東北電力グループ中長期ビジョン", "東北電力", "2024"),
    ("SEG011", "中国電力", "Chugoku Electric Power Co., Inc.", "9504", "東証プライム", "上場",
     "『アクションプラン2030』（2026年度〜）で系統増強・設備投資を計画。データセンター誘致地域の需要増を取り込む。",
     "低", "https://www.energia.co.jp/ir/irkeiei/gaiyou.html",
     "中期経営計画（アクションプラン）", "中国電力", "2026"),

    # --- SEG012 非常用発電機・UPS ---
    ("SEG012", "山洋電気", "Sanyo Denki Co., Ltd.", "6516", "東証プライム", "上場",
     "UPS・エンジン発電機ブランド『SANUPS』を展開。データセンター・IT設備の無停電電源需要を取り込む。",
     "中", "https://products.sanyodenki.com/ja/sanups/",
     "SANUPS プロダクトサイト", "山洋電気", "2026"),
    ("SEG012", "オムロン", "OMRON Corporation", "6645", "東証プライム", "上場",
     "IT〜産業用まで幅広いUPSをラインナップ。データセンター・設備向けバックアップ電源需要が追い風。※UPSは同社の一事業。",
     "低", "https://socialsolution.omron.com/jp/ja/products_service/ups/",
     "無停電電源装置（UPS）", "オムロン", "2026"),
]


def read(path):
    with open(path, encoding="utf-8", newline="") as f:
        r = csv.reader(f)
        return next(r), list(r)


chead, crows = read(DATA / "companies.csv")
shead, srows = read(DATA / "sources.csv")
cid = max(int(r[0][1:]) for r in crows)
sid = max(int(r[0][1:]) for r in srows)

new_c, new_s = [], []
for (seg, ja, en, tk, mk, ls, drv, conf, url, stitle, spub, spdate) in ADDS:
    cid += 1
    cnum = f"C{cid:03d}"
    src_id = ""
    if url:
        sid += 1
        src_id = f"S{sid:03d}"
        new_s.append([src_id, cnum, "報道・IR", stitle, spub, url, spdate, TODAY])
    new_c.append([cnum, seg, ja, en, tk, mk, ls, "日本", drv, conf, src_id, TODAY])

with open(DATA / "companies.csv", "a", encoding="utf-8", newline="") as f:
    csv.writer(f).writerows(new_c)
with open(DATA / "sources.csv", "a", encoding="utf-8", newline="") as f:
    csv.writer(f).writerows(new_s)

print(f"added companies: {len(new_c)} (through C{cid:03d})")
print(f"added sources:   {len(new_s)} (through S{sid:03d})")
