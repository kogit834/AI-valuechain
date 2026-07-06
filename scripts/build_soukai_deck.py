"""ANDPAD 本部総会 向けプレゼン資料（5分・全7枚）を生成する。
テーマ: 建設SaaSの我々が、なぜいま製造業か。AI市場の規模を“体感”に変え、
        DC建設費の内訳（ゼネコン vs 製造業）から“周辺の広大な市場”を示す。
配色: 白 × ANDPADレッド × グレー（既存デッキとトンマナ統一）。
出力: output/andpad_soukai_ai_deck.pptx
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pathlib import Path

# ---- palette (white / red / grey) ----
RED    = RGBColor(0xE6,0x00,0x12)   # ANDPAD red
RED2   = RGBColor(0xF2,0x55,0x60)   # lighter red (2nd tone)
REDWA  = RGBColor(0xFD,0xEC,0xEE)   # pink wash
REDBD  = RGBColor(0xF3,0xC7,0xCD)   # pink border
INK    = RGBColor(0x1A,0x1D,0x21)   # near-black heading
BODY   = RGBColor(0x3C,0x41,0x48)
MUTE   = RGBColor(0x8A,0x90,0x98)
CHAR   = RGBColor(0x45,0x4B,0x54)   # charcoal
CHARWA = RGBColor(0xEC,0xEE,0xF1)
PANEL  = RGBColor(0xF3,0xF4,0xF6)   # light grey panel
LINE   = RGBColor(0xE2,0xE4,0xE9)
WHITE  = RGBColor(0xFF,0xFF,0xFF)

FONT   = "Noto Sans JP"
FONT_M = "Consolas"

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]
SW, SH = prs.slide_width, prs.slide_height


def slide(bg=WHITE):
    s = prs.slides.add_slide(BLANK)
    r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, SH)
    r.fill.solid(); r.fill.fore_color.rgb = bg; r.line.fill.background()
    r.shadow.inherit = False
    s.shapes._spTree.remove(r._element); s.shapes._spTree.insert(2, r._element)
    return s


def rect(s, x, y, w, h, fill=None, line=None, line_w=1.0, radius=None):
    t = MSO_SHAPE.ROUNDED_RECTANGLE if radius is not None else MSO_SHAPE.RECTANGLE
    sh = s.shapes.add_shape(t, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.shadow.inherit = False
    if fill is None: sh.fill.background()
    else: sh.fill.solid(); sh.fill.fore_color.rgb = fill
    if line is None: sh.line.fill.background()
    else: sh.line.color.rgb = line; sh.line.width = Pt(line_w)
    if radius is not None:
        try: sh.adjustments[0] = radius
        except Exception: pass
    return sh


def text(s, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
         space_after=4, line_spacing=1.06):
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


def header(s, tile, title, eyebrow=None):
    rect(s, 0.85, 0.5, 0.52, 0.52, fill=RED, radius=0.18)
    text(s, 0.85, 0.5, 0.52, 0.52, [[R(tile, 15, WHITE, True, FONT_M)]],
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    if eyebrow:
        text(s, 1.55, 0.5, 11.2, 0.28, [[R(eyebrow, 11, RED, True, FONT_M, 1.5)]])
        text(s, 1.53, 0.74, 11.4, 0.5, [[R(title, 20, INK, True)]])
    else:
        text(s, 1.55, 0.5, 11.2, 0.55, [[R(title, 23, INK, True)]],
             anchor=MSO_ANCHOR.MIDDLE)
    rect(s, 0.85, 1.28, 11.63, 0.02, fill=LINE)


# =========================================================
# SLIDE 1 — 表紙
# =========================================================
s = slide()
rect(s, 0, 0, SW.inches, SH.inches, fill=WHITE)
rect(s, 0, 0, 0.30, SH.inches, fill=RED)
rect(s, 0.85, 1.35, 0.62, 0.62, fill=RED, radius=0.18)
text(s, 0.85, 1.35, 0.62, 0.62, [[R("AI", 18, WHITE, True, FONT_M)]],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text(s, 1.68, 1.43, 10.6, 0.5,
     [[R("ANDPAD 本部総会  ·  WHY MANUFACTURING NOW", 12.5, RED, True, FONT_M, 2)]],
     anchor=MSO_ANCHOR.MIDDLE)
text(s, 0.83, 2.25, 11.8, 2.3,
     [[R("建設SaaSの我々が、", 39, INK, True)],
      [R("なぜ、いま", 39, INK, True), R("製造業", 39, RED, True),
       R("なのか。", 39, INK, True)]],
     line_spacing=1.1)
text(s, 0.86, 4.45, 11.4, 0.7,
     [[R("AIが、日本の製造業を", 19, BODY, False),
       R("“歴史的な建設・据付ラッシュ”", 19, RED, True),
       R("に変えている。", 19, BODY, False)]])
rect(s, 0.88, 5.5, 4.2, 0.03, fill=RED)
text(s, 0.86, 5.7, 11.4, 0.95,
     [[R("しかもお金の大半は、建物ではなく“電気・機械設備”＝製造業に流れる。据付・試運転・保守まで人が動き続ける。", 13.5, BODY)],
      [R("増える現場、足りない人手。その差を埋めるのが、我々ANDPADの現場管理だ。", 13.5, BODY)]],
     line_spacing=1.28)
text(s, 0.86, 7.0, 11.4, 0.35,
     [[R("2026-07  /  出典：AIインフラ・バリューチェーン調査（国内外154社・決算/IR一次情報）＋公開データ", 10.5, MUTE, False, FONT_M)]])

# =========================================================
# SLIDE 2 — WHY：なぜ製造業がアツいのか（メカニズム）
# =========================================================
s = slide()
header(s, "1", "AIブームは“クラウドの話”ではない。電気と鉄と現場の話だ。", "なぜ、いま製造業がアツいのか")

steps = [
    ("きっかけ", "AIは“電気と設備の塊”", "DC（データセンター）1棟で数万世帯分の電力。膨大な発電・変圧器・冷却・半導体が要る。"),
    ("連鎖", "製造業が一斉に動き出す", "重電・電線・空調・UPS・半導体——各社が工場を増設し、機器の製造・納入・据付が全国で立ち上がる。"),
    ("結果", "“現場”が全国で急増する", "工場を建てる現場も、機器を据え付け・試運転・保守する現場も、同時多発で増えていく。"),
]
cx, cw, gap, cy, ch = 0.85, 3.75, 0.30, 1.62, 1.95
for i, (k, h, b) in enumerate(steps):
    x = cx + i * (cw + gap)
    rect(s, x, cy, cw, ch, fill=WHITE, line=REDBD, line_w=1.2, radius=0.06)
    rect(s, x, cy, cw, 0.10, fill=RED, radius=0.0)
    text(s, x+0.28, cy+0.28, cw-0.56, ch-0.4,
         [[R(k, 10.5, RED, True, FONT_M, 1)],
          [R(h, 16, INK, True)],
          [R(b, 11.5, BODY)]], space_after=7, line_spacing=1.16)
    if i < 2:
        text(s, x+cw-0.03, cy+0.55, 0.36, 0.6, [[R("→", 22, RED, True, FONT_M)]], align=PP_ALIGN.CENTER)

rect(s, 0.85, 3.95, 11.63, 1.35, fill=CHARWA, line=CHAR, line_w=1.2, radius=0.05)
text(s, 1.15, 4.18, 11.0, 0.5,
     [[R("いま製造業は、建設業のように", 17, INK, True),
       R("“現場だらけ”", 17, RED, True), R("になっている。", 17, INK, True)]])
text(s, 1.15, 4.78, 11.0, 0.45,
     [[R("＝ 我々が建設業で磨いた「現場管理」が、そのまま効く土俵が、製造業に新しく生まれている。", 12.5, BODY)]])

rect(s, 0.85, 5.55, 11.63, 1.25, fill=REDWA, line=REDBD, line_w=1, radius=0.05)
text(s, 1.15, 5.76, 11.0, 0.5,
     [[R("しかもこれは、一過性のブームではない。", 15.5, INK, True)]])
text(s, 1.15, 6.28, 11.0, 0.45,
     [[R("各社の受注残・設備投資計画・売上ガイダンスに裏打ちされた、数年〜十数年続く", 12.5, BODY),
       R("構造需要", 12.5, RED, True), R("だ。", 12.5, BODY)]])

# =========================================================
# SLIDE 3 — 規模を“体感”する（身近なものに換算）
# =========================================================
s = slide()
header(s, "2", "その金額、ピンとこない。だから“身近なもの”に換算する。", "AI投資の規模を、体感する")

# 左：換算行  /  right column removed for full-width rows
comp = [
    ("110兆円", "／年", "Big Tech 4社のAI投資（'26）",
     "日本の国家予算（115兆円）まるごと1年分。それを“毎年”、たった4社が。"),
    ("780兆円", "累計", "世界のAI DC投資（〜'30・McKinsey）",
     "日本のGDP（約600兆円）を超える規模。国家予算に換算すると 約7年分。"),
    ("4.8兆円", "対日", "米クラウド大手の対日DC投資（〜'30）",
     "大型DC（1棟 約300億円）を 150棟以上 建てられる額が、日本に落ちる。"),
    ("13倍", "電力", "国内DC・半導体の最大電力（10年で）",
     "56万→715万kW。増える分だけで 原発 約7基分 の新しい電気が要る（OCCTO）。"),
]
ty, rh = 1.55, 1.14
for i, (v, tag, lab, eq) in enumerate(comp):
    y = ty + i * rh
    if i % 2 == 1:
        rect(s, 0.85, y, 11.63, rh-0.06, fill=PANEL, radius=0.04)
    # left number block
    rect(s, 0.95, y+0.12, 0.62, 0.6, fill=RED, radius=0.14)
    text(s, 0.95, y+0.12, 0.62, 0.6, [[R(tag, 9, WHITE, True, FONT_M)]],
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, 1.75, y+0.06, 3.3, 0.78, [[R(v, 34, RED, True)]], anchor=MSO_ANCHOR.MIDDLE)
    text(s, 1.78, y+0.82, 3.3, 0.24, [[R(lab, 9.5, MUTE, True, FONT_M)]])
    # equivalence
    text(s, 5.2, y+0.06, 0.5, rh-0.2, [[R("≒", 22, CHAR, True, FONT_M)]],
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, 5.85, y, 6.55, rh-0.06, [[R(eq, 15, INK, True)]],
         anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.18)

rect(s, 0.85, 6.28, 11.63, 0.6, fill=RED, radius=0.05)
text(s, 0.85, 6.28, 11.63, 0.6,
     [[R("要するに——", 14, RGBColor(0xFF,0xCF,0xD4), True),
       R("国家予算級のお金が、日本の“建設と設備の現場”に流れ込んでくる。", 15.5, WHITE, True)]],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# =========================================================
# SLIDE 4 — DC1棟の建設費は、誰に入るのか（ゼネコン vs 製造業）
# =========================================================
s = slide()
header(s, "3", "データセンターは“建物”より“設備”。お金の3/4は製造業側へ。",
       "DC建設費の中身：建設（ゼネコン）と、電気・機械設備（製造業）の比率")

# 100%バー
bx, by, bw, bh = 0.85, 1.95, 11.63, 1.25
w_gen = bw * 0.25   # 建築（ゼネコン）
w_ele = bw * 0.50   # 電気設備
w_mec = bw * 0.25   # 機械（空調・冷却）
rect(s, bx, by, w_gen, bh, fill=CHAR)
rect(s, bx+w_gen, by, w_ele, bh, fill=RED)
rect(s, bx+w_gen+w_ele, by, w_mec, bh, fill=RED2)
# 白い区切り
rect(s, bx+w_gen-0.012, by, 0.024, bh, fill=WHITE)
rect(s, bx+w_gen+w_ele-0.012, by, 0.024, bh, fill=WHITE)
# セグメント内ラベル
text(s, bx, by, w_gen, bh, [[R("建築（躯体）", 11, WHITE, True)], [R("約25%", 20, WHITE, True)]],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, space_after=1, line_spacing=1.0)
text(s, bx+w_gen, by, w_ele, bh, [[R("電気設備", 11.5, WHITE, True)], [R("約50%", 22, WHITE, True)]],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, space_after=1, line_spacing=1.0)
text(s, bx+w_gen+w_ele, by, w_mec, bh, [[R("空調・機械ほか", 10.5, WHITE, True)], [R("約25%", 20, WHITE, True)]],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, space_after=1, line_spacing=1.0)
# 上部ブラケット
rect(s, bx, by-0.34, w_gen, 0.26, fill=CHARWA, line=CHAR, line_w=1, radius=0.1)
text(s, bx, by-0.34, w_gen, 0.26, [[R("建設（ゼネコン）", 10, CHAR, True, FONT_M)]],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
rect(s, bx+w_gen, by-0.34, w_ele+w_mec, 0.26, fill=REDWA, line=RED, line_w=1, radius=0.1)
text(s, bx+w_gen, by-0.34, w_ele+w_mec, 0.26,
     [[R("設備＝製造業（電気・機械メーカー＋設備工事）  約75%", 10.5, RED, True, FONT_M)]],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# 参考：他用途との比較
rect(s, 0.85, 3.55, 5.6, 1.35, fill=WHITE, line=LINE, line_w=1, radius=0.05)
text(s, 1.1, 3.72, 5.2, 0.3, [[R("設備工事の比率は、用途で全く違う", 10.5, RED, True, FONT_M)]])
bars = [("オフィス", 0.30), ("マンション", 0.25), ("データセンター", 0.75)]
bxx, bww = 3.35, 2.85
for i, (nm, r) in enumerate(bars):
    yy = 4.02 + i * 0.28
    hot = nm == "データセンター"
    text(s, 1.1, yy-0.02, 2.15, 0.26, [[R(nm, 10, (RED if hot else BODY), hot)]], anchor=MSO_ANCHOR.MIDDLE)
    rect(s, bxx, yy+0.02, bww, 0.18, fill=PANEL, radius=0.3)
    rect(s, bxx, yy+0.02, bww*r, 0.18, fill=(RED if hot else MUTE), radius=0.3)
    text(s, bxx+bww+0.06, yy-0.02, 0.7, 0.26, [[R(f"{int(r*100)}%", 10, (RED if hot else MUTE), hot, FONT_M)]], anchor=MSO_ANCHOR.MIDDLE)

# パンチライン
rect(s, 6.65, 3.55, 5.83, 1.35, fill=CHARWA, line=CHAR, line_w=1.2, radius=0.05)
text(s, 6.9, 3.72, 5.4, 1.1,
     [[R("建設業だけを見るのは、", 15.5, INK, True),
       R("氷山の一角。", 15.5, RED, True)],
      [R("お金の“3/4”は、電気・機械設備＝製造業の世界にある。", 12.5, BODY)]],
     space_after=6, line_spacing=1.16)

# 下段：金額の実感
rect(s, 0.85, 5.15, 11.63, 1.0, fill=REDWA, line=REDBD, line_w=1, radius=0.05)
text(s, 1.15, 5.15, 11.0, 1.0,
     [[R("大型DC1棟＝数百億〜1,000億円級（IT機器は別）。", 14, INK, True),
       R("その約3/4、1棟あたり数百億円が、製造業側の“設備”に向かう。", 14, RED, True)]],
     anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.2)
text(s, 0.9, 6.3, 11.5, 0.4,
     [[R("出典：日経xTECH（DC建設費は電気設備 約50%＋空調等 約25%＋建築 約25%）／建設費/MW：Cushman & Wakefield・Turner & Townsend。", 9, MUTE, False, FONT_M)]])

# =========================================================
# SLIDE 5 — 設備は“人が現場で動き続ける”（周辺市場の正体）
# =========================================================
s = slide()
header(s, "4", "しかも設備は、作って終わりじゃない。据付・試運転・保守で人が動く。",
       "周辺市場の正体：モノ売りではなく、“現場”がずっと続く")

phases = [
    ("①", "製造", "○ 工場のなか", False, "機器をつくる。ここはメーカーの工場の中。"),
    ("②", "据付・施工", "● 現場ではじまる", True, "搬入・配線・配管・据付。ここから“現場”が始まる。"),
    ("③", "試運転・調整", "● 現場でつづく", True, "コミッショニング。建設費の1〜3%が試運転に。"),
    ("④", "アフター保守・運用", "● 現場でずっと", True, "引き渡し後も点検・更新が続く。建設は一度、保守は毎年。"),
]
py, ph = 1.6, 1.85
pcol0 = 0.85
pcw = (12.48 - pcol0) / 4
for i, (no, nm, badge, onsite, desc) in enumerate(phases):
    x = pcol0 + i * pcw
    bxw = pcw - 0.14
    col = RED if onsite else CHAR
    rect(s, x, py, bxw, ph, fill=WHITE, line=(REDBD if onsite else CHAR),
         line_w=1.2, radius=0.06)
    rect(s, x, py, bxw, 0.10, fill=col)
    text(s, x+0.2, py+0.24, bxw-0.4, 0.9,
         [[R(no + "  " + nm, 14, INK, True)],
          [R(badge, 9.5, (RED if onsite else MUTE), True, FONT_M)],
          [R(desc, 10, BODY)]], space_after=5, line_spacing=1.16)
    if i < 3:
        text(s, x+bxw-0.02, py+0.5, 0.28, 0.6, [[R("→", 18, RED, True, FONT_M)]], align=PP_ALIGN.CENTER)

# 4工程のうち3つが“現場”
rect(s, 0.85, 3.62, 11.63, 0.5, fill=PANEL, line=LINE, line_w=1, radius=0.06)
text(s, 1.1, 3.62, 11.2, 0.5,
     [[R("4工程のうち ", 12, BODY),
       R("3つが“現場”", 12, RED, True),
       R("——製造業の設備ビジネスは、モノを売って終わりではなく、人が現場で動き続ける労働集約の世界。", 12, INK, True)]],
     anchor=MSO_ANCHOR.MIDDLE)

# O&M市場の伸び
rect(s, 0.85, 4.3, 5.6, 1.55, fill=WHITE, line=REDBD, line_w=1.2, radius=0.05)
text(s, 1.1, 4.48, 5.2, 0.3, [[R("保守（O&M）は“毎年”積み上がる別市場", 10.5, RED, True, FONT_M)]])
text(s, 1.1, 4.82, 5.2, 0.9,
     [[R("$15.8B → $45.6B", 22, RED, True, FONT_M)],
      [R("DC運用保守サービス市場（世界・2024→2033）。年率+11%で拡大。", 10, BODY)]],
     space_after=4, line_spacing=1.15)

# パンチライン（＝ANDPADの主戦場）
rect(s, 6.65, 4.3, 5.83, 1.55, fill=REDWA, line=REDBD, line_w=1.2, radius=0.05)
text(s, 6.9, 4.5, 5.4, 1.2,
     [[R("＝ 建設の“周辺”に、より大きく・より長く続く", 13.5, INK, True)],
      [R("“人が動く現場”の市場", 15.5, RED, True), R("が広がっている。", 13.5, INK, True)],
      [R("ここはANDPADの現場管理が、そのまま効く。", 12.5, BODY)]],
     space_after=5, line_spacing=1.16)

rect(s, 0.85, 6.02, 11.63, 0.78, fill=CHARWA, line=CHAR, line_w=1.2, radius=0.05)
text(s, 1.15, 6.02, 11.0, 0.78,
     [[R("これまで建設業だけを相手にしてきた。", 14.5, INK, True),
       R("その周辺に、設備＋据付＋試運転＋保守という広大な市場がある。", 14.5, RED, True)]],
     anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.15)

# =========================================================
# SLIDE 6 — その現場は、日本のこの企業に立ち上がっている（バリューチェーン）
# =========================================================
s = slide()
header(s, "5", "国内のAI投資は、製造業のこの現場に着地している。", "AI需要は、バリューチェーンのどこに効くのか（国内主要プレーヤー）")

flow = [
    ("01", "発電・電源", "電気をつくる", "受注残 5兆円超", "三菱重工 ガスタービン"),
    ("02", "送変電・電線", "電気を送る・変える", "生産能力2倍/予約数年先", "変圧器各社・フジクラ"),
    ("03", "DC建設・設備工事", "箱を建てる", "空調工事 受注 +25.5%", "ダイダン・高砂熱学ほか"),
    ("04", "冷却・電源保護", "冷やす・止めない", "北米DC冷却 約13倍", "ダイキン 230→3,000億円"),
    ("05", "半導体", "計算する頭脳", "設備投資 +66%", "キオクシア・東京ｴﾚｸﾄﾛﾝ"),
]
fy, fh = 1.55, 0.86
by, bh = 2.52, 2.28
col0 = 1.95
cw = (12.48 - col0) / 5
rect(s, 0.85, fy, 0.98, fh, fill=CHAR, radius=0.09)
text(s, 0.85, fy, 0.98, fh,
     [[R("起点", 8, RGBColor(0xD7,0xDB,0xE0), True, FONT_M, 1)],
      [R("AI需要", 12, WHITE, True)]],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, space_after=0, line_spacing=1.0)
for i, (no, nm, sb, sv, chips) in enumerate(flow):
    cx = col0 + i * cw
    chv = s.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(cx-0.02), Inches(fy), Inches(cw+0.14), Inches(fh))
    chv.shadow.inherit = False
    chv.fill.solid(); chv.fill.fore_color.rgb = RED; chv.line.fill.background()
    chv.text_frame.text = ""
    text(s, cx-0.04, fy, cw+0.06, fh,
         [[R(no, 8, RGBColor(0xFF,0xC9,0xCE), True, FONT_M, 1)],
          [R(nm, 11, WHITE, True)]],
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, space_after=0, line_spacing=1.0)
    bx = cx + 0.05; bw = cw - 0.10
    rect(s, bx, by, bw, bh, fill=WHITE, line=LINE, line_w=1, radius=0.05)
    pad = 0.13
    text(s, bx+pad, by+0.12, bw-2*pad, 0.3, [[R(sb, 11, INK, True)]])
    rect(s, bx+pad, by+0.56, bw-2*pad, 0.01, fill=REDBD)
    text(s, bx+pad, by+0.66, bw-2*pad, 0.66, [[R(sv, 12.5, RED, True)]], line_spacing=1.04)
    text(s, bx+pad, by+1.5, bw-2*pad, 0.22, [[R("代表企業", 7, MUTE, True, FONT_M)]])
    text(s, bx+pad, by+1.7, bw-2*pad, 0.55, [[R(chips, 8.8, INK, True)]], line_spacing=1.14)

# 2モデル凡例
rect(s, 0.85, 5.0, 11.63, 0.62, fill=PANEL, line=LINE, line_w=1, radius=0.06)
text(s, 1.1, 5.0, 11.2, 0.62,
     [[R("我々の入り方：", 10.5, MUTE, True, FONT_M),
       R(" A 発注者＝工場増設・自社DCを“建てる側”で管理", 11, CHAR, True),
       R("　／　", 10.5, MUTE, False),
       R("B 請負＝据付・試運転・保守を“納める側”で管理（本命）", 11, RED, True)]],
     anchor=MSO_ANCHOR.MIDDLE)

rect(s, 0.85, 5.8, 11.63, 1.0, fill=CHARWA, line=CHAR, line_w=1.2, radius=0.05)
text(s, 1.15, 5.8, 11.0, 1.0,
     [[R("川上（発電）から川下（半導体）まで全段が同時に増収増益。", 14.5, INK, True)],
      [R("＝ 攻めどころは1つではなく、チェーン全体にある。狙える現場が、日本中に生まれている。", 13, RED, True)]],
     anchor=MSO_ANCHOR.MIDDLE, space_after=5, line_spacing=1.16)

# =========================================================
# SLIDE 7 — クロージング（モチベーション）
# =========================================================
s = slide()
rect(s, 0, 0, SW.inches, SH.inches, fill=INK)
rect(s, 0, 0, 0.30, SH.inches, fill=RED)
rect(s, 0.85, 0.9, 0.62, 0.62, fill=RED, radius=0.18)
text(s, 0.85, 0.9, 0.62, 0.62, [[R("6", 18, WHITE, True, FONT_M)]],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text(s, 1.68, 0.98, 10.6, 0.5,
     [[R("SO, LET'S GO  ·  次の主戦場は、製造業だ", 12.5, RGBColor(0xFF,0x8A,0x97), True, FONT_M, 2)]],
     anchor=MSO_ANCHOR.MIDDLE)
text(s, 0.83, 1.95, 11.8, 1.7,
     [[R("いま製造業に行くことは、", 34, WHITE, True)],
      [R("“次の主戦場”の一番槍", 34, RGBColor(0xFF,0x6B,0x7A), True),
       R("だ。", 34, WHITE, True)]],
     line_spacing=1.12)

pts = [
    ("市場は建設の3倍広い", "DC建設費の約3/4は電気・機械設備＝製造業。据付・試運転・保守まで現場が続く。"),
    ("追い風は本物", "国家予算級のAIマネーが国内に流入。受注残・設備投資計画に裏打ちされた構造需要。"),
    ("武器は完成済み", "建設で磨いた現場管理は、製造業の据付・保守フィールドに“そのまま効く”。"),
]
py = 3.95
pw = (11.63 - 0.6) / 3
for i, (h, b) in enumerate(pts):
    x = 0.85 + i * (pw + 0.3)
    rect(s, x, py, pw, 1.65, fill=RGBColor(0x24,0x28,0x2E), line=RGBColor(0x3A,0x40,0x48), line_w=1, radius=0.06)
    rect(s, x, py, pw, 0.09, fill=RED)
    text(s, x+0.26, py+0.28, pw-0.52, 1.3,
         [[R(h, 14.5, WHITE, True)],
          [R(b, 10.5, RGBColor(0xC7,0xD0,0xDC))]], space_after=8, line_spacing=1.2)

rect(s, 0.85, 5.98, 11.63, 0.86, fill=RED, radius=0.06)
text(s, 0.85, 5.98, 11.63, 0.86,
     [[R("建設の“周辺”に広がる巨大市場を、100人でつかみにいこう。", 21, WHITE, True)]],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

out = Path(__file__).resolve().parent.parent / "output" / "andpad_soukai_ai_deck.pptx"
prs.save(str(out))
print("saved:", out, "/ slides:", len(prs.slides._sldIdLst))
