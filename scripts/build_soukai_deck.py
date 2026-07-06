"""ANDPAD 本部総会 向けプレゼン資料（5分・全7枚）を生成する。
デザイン方針: 箱・塗り・角丸を排し、白地・余白・細い罫線・タイポグラフィで見せる
エディトリアル調。赤はアクセント1色（キー数字/語・細い罫線）に限定。
出力: output/andpad_soukai_ai_deck.pptx
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pathlib import Path

# ---- palette ----
RED   = RGBColor(0xE6,0x00,0x12)
INK   = RGBColor(0x17,0x1A,0x1E)   # 主要テキスト（ほぼ黒）
SUB   = RGBColor(0x3C,0x41,0x48)   # 本文
MUTE  = RGBColor(0x7A,0x82,0x8B)   # 副次テキスト
FAINT = RGBColor(0xA5,0xAB,0xB2)   # ごく薄い注記
HAIR  = RGBColor(0xDC,0xDF,0xE4)   # 罫線
G1    = RGBColor(0xCF,0xD3,0xD8)   # バー（淡）
G2    = RGBColor(0xAD,0xB2,0xB9)   # バー（中）
G_DK  = RGBColor(0x5A,0x60,0x68)   # バー（濃）
WHITE = RGBColor(0xFF,0xFF,0xFF)

FONT   = "Noto Sans JP"
FONT_M = "Consolas"

LM, RM = 0.9, 12.43          # 左右マージン
CW = RM - LM                 # コンテンツ幅 = 11.53

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]
SW, SH = prs.slide_width, prs.slide_height


def slide():
    s = prs.slides.add_slide(BLANK)
    r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, SH)
    r.fill.solid(); r.fill.fore_color.rgb = WHITE; r.line.fill.background()
    r.shadow.inherit = False
    s.shapes._spTree.remove(r._element); s.shapes._spTree.insert(2, r._element)
    return s


def box(s, x, y, w, h, fill):
    sh = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.shadow.inherit = False
    sh.fill.solid(); sh.fill.fore_color.rgb = fill; sh.line.fill.background()
    return sh


def hrule(s, x, y, w, color=HAIR, wt=1.0):
    box(s, x, y, w, wt/72.0, color)


def vrule(s, x, y, h, color=HAIR, wt=1.0):
    box(s, x, y, wt/72.0, h, color)


def text(s, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
         space_after=4, line_spacing=1.08):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align; p.space_before = Pt(0); p.space_after = Pt(space_after)
        p.line_spacing = line_spacing
        for (txt, size, color, bold, font, tracking) in para:
            r = p.add_run(); r.text = txt
            r.font.size = Pt(size); r.font.bold = bold
            r.font.color.rgb = color; r.font.name = font
            rPr = r._r.get_or_add_rPr()
            ea = rPr.find(qn('a:ea'))
            if ea is None:
                ea = rPr.makeelement(qn('a:ea'), {}); rPr.append(ea)
            ea.set('typeface', font)
            if tracking: rPr.set('spc', str(int(tracking * 100)))
    return tb


def R(txt, size, color=INK, bold=False, font=FONT, tracking=0):
    return (txt, size, color, bold, font, tracking)


def header(s, page, eyebrow, title):
    box(s, LM, 0.64, 0.12, 0.12, RED)
    text(s, 1.11, 0.57, 9.5, 0.26, [[R(eyebrow, 10.5, RED, True, FONT_M, 1.5)]])
    text(s, RM-0.6, 0.57, 0.6, 0.26, [[R(page, 10.5, FAINT, True, FONT_M)]], align=PP_ALIGN.RIGHT)
    text(s, LM-0.02, 0.9, CW, 0.5, [[R(title, 21, INK, True)]])
    hrule(s, LM, 1.52, CW, wt=1.2)


# =========================================================
# SLIDE 1 — 表紙
# =========================================================
s = slide()
box(s, LM, 1.55, 0.12, 0.12, RED)
text(s, 1.11, 1.48, 10.0, 0.28,
     [[R("ANDPAD 本部総会   ·   WHY MANUFACTURING NOW", 11.5, RED, True, FONT_M, 2)]])
text(s, LM-0.03, 2.35, 11.7, 2.2,
     [[R("建設SaaSの我々が、", 40, INK, True)],
      [R("なぜ、いま", 40, INK, True), R("製造業", 40, RED, True), R("なのか。", 40, INK, True)]],
     line_spacing=1.12)
hrule(s, LM, 4.55, 3.0, color=RED, wt=2.4)
text(s, LM-0.02, 4.85, 11.6, 1.0,
     [[R("AIをはじめとする世界のメガトレンドは、膨大な“設備”を必要とする。", 15, SUB)],
      [R("そして設備を作り・建てるのは、建設業だけでなく製造業がコア——その市場が、いま圧倒的に広がっている。", 15, SUB)]],
     line_spacing=1.4)
hrule(s, LM, 6.85, CW)
text(s, LM-0.02, 6.98, 11.6, 0.3,
     [[R("2026-07   /   AIインフラ・バリューチェーン調査（国内外154社・決算/IR一次情報）＋公開データ", 10, FAINT, False, FONT_M)]])

# =========================================================
# SLIDE 2 — メガトレンド → 設備 → 製造業がコア → AIにフォーカス
# =========================================================
s = slide()
header(s, "01", "世界のメガトレンドと、その“共通の土台”",
       "世界が動く先には、必ず「設備」がいる。その主役は、製造業だ。")

cols = [
    ("01", "AI", True,  "生成AI・データセンター・半導体", "電源・冷却・建屋・チップ工場が大量に要る"),
    ("02", "脱炭素（GX）", False, "再エネ・送電網・電化・蓄電池", "発電・変電・ケーブル・電池／EV工場"),
    ("03", "経済安保・国内回帰", False, "半導体・重要物資の国産化", "国内に工場を新設し、供給網を再構築"),
]
pitch = CW / 3
ctop = 1.95
for i, (no, nm, hot, what, need) in enumerate(cols):
    x = LM + i * pitch
    if i > 0:
        vrule(s, x-0.22, ctop, 1.95)
    text(s, x, ctop, pitch-0.5, 0.24, [[R(no, 10, FAINT, True, FONT_M, 1)]])
    text(s, x, ctop+0.28, pitch-0.5, 0.44, [[R(nm, 19, (RED if hot else INK), True)]])
    if hot:
        text(s, x, ctop+0.78, pitch-0.5, 0.24, [[R("● 本日フォーカス", 9.5, RED, True, FONT_M)]])
    else:
        text(s, x, ctop+0.78, pitch-0.5, 0.24, [[R("　", 9.5, MUTE)]])
    text(s, x, ctop+1.12, pitch-0.55, 0.4, [[R(what, 12, MUTE, True)]], line_spacing=1.14)
    text(s, x, ctop+1.5, pitch-0.55, 0.5, [[R("→ ", 12, RED, True), R(need, 12, SUB)]], line_spacing=1.2)

hrule(s, LM, 4.28, CW)
text(s, LM-0.02, 4.55, 11.6, 0.5,
     [[R("どのメガトレンドも、実現するには", 17, INK, True),
       R("膨大な“設備”", 17, RED, True), R("がいる。", 17, INK, True)]])
text(s, LM-0.02, 5.15, 11.6, 0.45,
     [[R("そして設備を作る・建てるのは、建設業だけではない——", 13.5, SUB),
       R("重電・機械・電機など製造業がコア。", 13.5, INK, True)]])
hrule(s, LM, 5.95, CW)
text(s, LM-0.02, 6.15, 11.6, 0.5,
     [[R("→ 今日はその中で、最も勢いのある「", 15, INK, True),
       R("AI", 15, RED, True), R("」に絞って話す。", 15, INK, True)]])

# =========================================================
# SLIDE 3 — トレンド：これまでの何倍か
# =========================================================
s = slide()
header(s, "02", "規模より“勢い”——トレンドとして、どれだけ急か",
       "AI投資は“大きい”だけじゃない。数年で“何倍”に跳ねている。")


def vbars(s, x0, base_y, area_w, max_h, values, years, val_label):
    n = len(values); mx = max(values)
    slot = area_w / n; bw = min(0.42, slot * 0.5)
    hrule(s, x0, base_y, area_w, color=HAIR, wt=1.2)
    for i, v in enumerate(values):
        h = max_h * v / mx
        bxx = x0 + i * slot + (slot - bw) / 2
        box(s, bxx, base_y - h, bw, h, RED if i == n-1 else G1)
        text(s, x0 + i*slot, base_y + 0.06, slot, 0.22,
             [[R(years[i], 8.5, MUTE, True, FONT_M)]], align=PP_ALIGN.CENTER)
        if i == n-1:
            text(s, bxx-0.6, base_y-h-0.24, bw+1.2, 0.22,
                 [[R(val_label, 9, MUTE, True, FONT_M)]], align=PP_ALIGN.CENTER)


rows = [
    ("AIチップの需要", "NVIDIA データセンター売上", [15, 48, 115, 196],
     ["FY23", "FY24", "FY25", "FY26"], "約2,000億ドル", "約13倍"),
    ("AIインフラ投資", "Big Tech 4社の設備投資（年間）", [180, 230, 410, 725],
     ["’23", "’24", "’25", "’26"], "7,250億ドル", "約4倍"),
    ("国内DC建設投資", "日本のデータセンター建設（年間）", [3222, 5000, 10000],
     ["’23", "’24", "’28"], "1兆円超", "約3倍"),
]
ty, rh = 1.75, 1.28
for i, (what, note, vals, yrs, vlab, mult) in enumerate(rows):
    y = ty + i * rh
    if i > 0:
        hrule(s, LM, y-0.02, CW)
    text(s, LM, y+0.28, 3.0, 0.85,
         [[R(what, 15, INK, True)], [R(note, 9.5, MUTE, True, FONT_M)]],
         space_after=3, line_spacing=1.14)
    vbars(s, 4.15, y+rh-0.4, 3.85, 0.6, vals, yrs, vlab)
    text(s, 8.7, y+0.2, 1.6, 0.85, [[R(mult, 32, RED, True)]], anchor=MSO_ANCHOR.MIDDLE)
    text(s, 10.35, y+0.2, 2.1, 0.85,
         [[R("に伸びた", 11, MUTE, True)]], anchor=MSO_ANCHOR.MIDDLE)

hrule(s, LM, 5.75, CW, wt=1.2)
text(s, LM-0.02, 5.95, 11.6, 0.5,
     [[R("投資のピークは2027〜2028年。この波は、まだ“", 16, INK, True),
       R("序盤", 16, RED, True), R("”だ。", 16, INK, True)]])
text(s, LM-0.02, 6.72, 11.6, 0.28,
     [[R("出典：NVIDIA IR／Big Tech各社IR・報道／IDC Japan（国内DC建設投資）。", 8.5, FAINT, False, FONT_M)]])

# =========================================================
# SLIDE 4 — バリューチェーン概観
# =========================================================
s = slide()
header(s, "03", "AI需要 → 連鎖して伸びる、日本のバリューチェーン",
       "AIが伸びれば、この“川”がまるごと潤う。")

text(s, LM, 1.85, 2.0, 0.3, [[R("AI・DC需要の拡大", 12, RED, True)]])
text(s, LM, 2.12, 2.0, 0.3, [[R("この一手が、川下までまるごと波及する", 9.5, MUTE, True)]])
hrule(s, LM, 2.55, CW, color=RED, wt=1.6)
text(s, LM, 2.62, CW, 0.24, [[R("↓", 12, RED, True)]])

stages = [
    ("01", "発電・電源", ["三菱重工", "川崎重工", "IHI", "デンヨー"]),
    ("02", "送変電・電線", ["日立", "三菱電機", "ダイヘン", "フジクラ"]),
    ("03", "DC建設・設備", ["鹿島建設", "大林組", "きんでん", "高砂熱学"]),
    ("04", "冷却・電源保護", ["ダイキン", "GSユアサ", "荏原製作所"]),
    ("05", "半導体", ["東京エレクトロン", "ディスコ", "キオクシア", "信越化学"]),
]
spitch = CW / 5
stop = 3.05
for i, (no, stg, cos) in enumerate(stages):
    x = LM + i * spitch
    if i > 0:
        vrule(s, x-0.12, stop, 2.15)
    text(s, x, stop, spitch-0.3, 0.2, [[R(no, 9, FAINT, True, FONT_M, 1)]])
    text(s, x, stop+0.22, spitch-0.28, 0.5, [[R(stg, 12, INK, True)]], line_spacing=1.05)
    for j, co in enumerate(cos):
        text(s, x, stop+0.74 + j*0.32, spitch-0.28, 0.3, [[R(co, 10.5, SUB, True)]])

hrule(s, LM, 5.55, CW, wt=1.2)
text(s, LM-0.02, 5.78, 11.6, 0.5,
     [[R("AIの“源流”が、川下の日本メーカーまで、まるごと潤す。", 16, INK, True)]])
text(s, LM-0.02, 6.28, 11.6, 0.4,
     [[R("狙える現場は、この一社一社にある。", 13, RED, True)]])
text(s, LM-0.02, 6.85, 11.6, 0.28,
     [[R("※ 各段階の代表企業を抜粋（社名表記＝ロゴのイメージ、実ロゴへ差し替え可）。数値・詳細は次頁以降。", 8.5, FAINT, False, FONT_M)]])

# =========================================================
# SLIDE 5 — 建設費の中身：ANDPADの市場はどこか
# =========================================================
s = slide()
header(s, "04", "建設費の中身：どこにANDPADの市場があるか",
       "お金の3/4は設備。効くのは機器代ではなく“工”＝据付・試運転・保守。")

# Bar A：建設費 100%
bx, by, bw, bh = LM, 2.05, CW, 0.82
wg = bw*0.25; we = bw*0.50; wm = bw*0.25
box(s, bx, by, wg, bh, G_DK)
box(s, bx+wg, by, we, bh, G1)
box(s, bx+wg+we, by, wm, bh, G2)
box(s, bx+wg-0.01, by, 0.02, bh, WHITE)
box(s, bx+wg+we-0.01, by, 0.02, bh, WHITE)
text(s, bx, by, wg, bh, [[R("建築（躯体）", 10, WHITE, True)], [R("約25%", 16, WHITE, True)]],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, space_after=0, line_spacing=1.0)
text(s, bx+wg, by, we, bh, [[R("電気設備", 10.5, INK, True)], [R("約50%", 18, INK, True)]],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, space_after=0, line_spacing=1.0)
text(s, bx+wg+we, by, wm, bh, [[R("空調・機械", 10, INK, True)], [R("約25%", 16, INK, True)]],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, space_after=0, line_spacing=1.0)
text(s, bx, by-0.28, wg, 0.24, [[R("建設（ゼネコン）", 9.5, MUTE, True, FONT_M)]])
text(s, bx+wg, by-0.28, we+wm, 0.24, [[R("設備＝製造業（電気・機械）  約75%", 9.5, RED, True, FONT_M)]])

# Bar B：材 vs 工
text(s, LM, 3.15, CW, 0.26,
     [[R("その「設備 約75%」を  ", 11, SUB, True),
       R("材（機器・材料）", 11, MUTE, True), R("  と  ", 11, SUB),
       R("工（据付・施工＝人が動く）", 11, RED, True), R("  に分けると", 11, SUB, True)]])
b2y, b2h = 3.48, 0.78
wmat = CW*0.60; wwork = CW*0.40
box(s, LM, b2y, wmat, b2h, G1)
box(s, LM+wmat, b2y, wwork, b2h, RED)
box(s, LM+wmat-0.01, b2y, 0.02, b2h, WHITE)
text(s, LM, b2y, wmat, b2h, [[R("材：機器・材料（変圧器・冷凍機・UPS 等）  約6割", 11, INK, True)]],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text(s, LM+wmat, b2y, wwork, b2h, [[R("工：据付・施工  約4割", 11, WHITE, True)]],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text(s, LM+wmat, b2y+b2h+0.05, wwork, 0.22,
     [[R("↑ ここがANDPADの市場（工）", 9, RED, True, FONT_M)]], align=PP_ALIGN.CENTER)

# 工の3工程（罫線区切り・箱なし）
wy = 4.85
work = [
    ("据付・施工", "建設費の約3割", "国内の設備工事そのもの＝中核市場"),
    ("試運転", "建設費の1〜3%", "世界の試運転市場 $2.1B → $4.3B"),
    ("保守・運用", "毎年・運用費の約4割", "世界のDC運用保守 $15.8B → $45.6B"),
]
wpitch = CW / 3
for i, (nm, pct, mk) in enumerate(work):
    x = LM + i*wpitch
    if i > 0:
        vrule(s, x-0.18, wy, 0.66)
    text(s, x, wy, wpitch-0.4, 0.28, [[R(nm, 12, INK, True), R("  "+pct, 11, RED, True)]])
    text(s, x, wy+0.32, wpitch-0.4, 0.3, [[R(mk, 9.5, MUTE, True)]], line_spacing=1.12)

hrule(s, LM, 5.85, CW, wt=1.2)
text(s, LM-0.02, 6.05, 11.6, 0.5,
     [[R("これまで建設業だけを見てきた。その周辺の“工”（据付・試運転・保守）に、", 14, INK, True),
       R("広大な市場", 14, RED, True), R("がある。", 14, INK, True)]])
text(s, LM-0.02, 6.78, 11.6, 0.26,
     [[R("出典：日経xTECH（設備工事≒建設費の75%）／材工比・試運転1〜3%は積算・業界目安／市場：DataIntelo・DataHorizon。", 8.5, FAINT, False, FONT_M)]])

# =========================================================
# SLIDE 6 — 主要企業：国内のどこに効いているか（裏付け）
# =========================================================
s = slide()
header(s, "05", "バリューチェーン別・国内主要プレーヤーと伸び",
       "国内のAI投資は、製造業のこの現場に着地している。")

flow = [
    ("01", "発電・電源", "電気をつくる", "受注残 5兆円超", "三菱重工"),
    ("02", "送変電・電線", "電気を送る・変える", "変圧器 生産能力2倍", "日立・ダイヘン・フジクラ"),
    ("03", "DC建設・設備工事", "箱を建てる", "空調工事 受注 +25.5%", "ダイダン・高砂熱学"),
    ("04", "冷却・電源保護", "冷やす・止めない", "北米DC冷却 約13倍", "ダイキン・GSユアサ"),
    ("05", "半導体", "計算する頭脳", "設備投資 +66%", "キオクシア・東京ｴﾚｸﾄﾛﾝ"),
]
fpitch = CW / 5
ftop = 1.95
for i, (no, stg, sb, stat, cos) in enumerate(flow):
    x = LM + i * fpitch
    if i > 0:
        vrule(s, x-0.12, ftop, 2.85)
    text(s, x, ftop, fpitch-0.28, 0.2, [[R(no, 9, FAINT, True, FONT_M, 1)]])
    text(s, x, ftop+0.22, fpitch-0.26, 0.46, [[R(stg, 12, INK, True)]], line_spacing=1.04)
    text(s, x, ftop+0.7, fpitch-0.26, 0.24, [[R(sb, 9.5, MUTE, True)]])
    text(s, x, ftop+1.12, fpitch-0.3, 0.7, [[R(stat, 14.5, RED, True)]], line_spacing=1.08)
    text(s, x, ftop+2.15, fpitch-0.28, 0.6, [[R(cos, 9.5, SUB, True)]], line_spacing=1.16)

hrule(s, LM, 5.05, CW)
text(s, LM-0.02, 5.28, 11.6, 0.5,
     [[R("我々の入り方：", 11, MUTE, True, FONT_M),
       R("  A 発注者＝工場増設・自社DCを“建てる側”で管理", 12, SUB, True),
       R("   /   ", 11, FAINT),
       R("B 請負＝据付・試運転・保守を“納める側”で管理（本命）", 12, RED, True)]])
hrule(s, LM, 5.95, CW, wt=1.2)
text(s, LM-0.02, 6.15, 11.6, 0.5,
     [[R("川上から川下まで、全段が同時に増収増益。", 15, INK, True),
       R("狙える現場が、日本中に生まれている。", 15, RED, True)]])

# =========================================================
# SLIDE 7 — クロージング
# =========================================================
s = slide()
box(s, LM, 1.15, 0.12, 0.12, RED)
text(s, 1.11, 1.08, 10.0, 0.28,
     [[R("SO, LET'S GO", 11.5, RED, True, FONT_M, 2)]])
text(s, LM-0.03, 1.7, 11.7, 1.1,
     [[R("次の主戦場は、", 38, INK, True), R("製造業", 38, RED, True), R("だ。", 38, INK, True)]])
hrule(s, LM, 3.15, 3.0, color=RED, wt=2.4)

pts = [
    ("市場は建設の3倍広い", "DC建設費の約3/4は電気・機械設備＝製造業。据付・試運転・保守まで現場が続く。"),
    ("追い風は本物", "AIチップ需要は3年で約13倍、投資ピークは2027-28。受注残に裏打ちされた構造需要。"),
    ("武器は完成済み", "建設で磨いた現場管理は、製造業の据付・保守フィールドに“そのまま効く”。"),
]
py = 3.6
for i, (h, b) in enumerate(pts):
    y = py + i * 0.86
    if i > 0:
        hrule(s, LM, y-0.06, CW)
    text(s, LM, y, 0.55, 0.6, [[R(f"0{i+1}", 15, RED, True, FONT_M)]], anchor=MSO_ANCHOR.MIDDLE)
    text(s, LM+0.7, y+0.06, 4.4, 0.6, [[R(h, 15, INK, True)]], anchor=MSO_ANCHOR.MIDDLE)
    text(s, LM+5.3, y+0.06, 6.3, 0.6, [[R(b, 10.5, SUB)]], anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.14)

hrule(s, LM, 6.5, CW, color=RED, wt=1.6)
text(s, LM-0.02, 6.68, 11.6, 0.5,
     [[R("建設の“周辺”に広がる巨大市場を、100人でつかみにいこう。", 18, INK, True)]])

out = Path(__file__).resolve().parent.parent / "output" / "andpad_soukai_ai_deck.pptx"
prs.save(str(out))
print("saved:", out, "/ slides:", len(prs.slides._sldIdLst))
