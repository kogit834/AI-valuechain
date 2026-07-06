"""ANDPAD 本部総会 向けプレゼン資料（5分・全6枚）を生成する。
テーマ: 建設SaaSの我々が、なぜいま製造業か。AI市場の規模と、我々の出番。
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
        text(s, 1.53, 0.74, 11.4, 0.5, [[R(title, 21, INK, True)]])
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
     [[R("世界のAIマネーが、発電・変圧器・冷却・半導体の“設備投資と現場”になって日本に降ってくる。", 13.5, BODY)],
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

# 中段：建設業とのアナロジー
rect(s, 0.85, 3.95, 11.63, 1.35, fill=CHARWA, line=CHAR, line_w=1.2, radius=0.05)
text(s, 1.15, 4.18, 11.0, 0.5,
     [[R("いま製造業は、建設業のように", 17, INK, True),
       R("“現場だらけ”", 17, RED, True),
       R("になっている。", 17, INK, True)]])
text(s, 1.15, 4.78, 11.0, 0.45,
     [[R("＝ 我々が建設業で磨いた「現場管理」が、そのまま効く土俵が、製造業に新しく生まれている。", 12.5, BODY)]])

# 下段：これは“思惑”ではない
rect(s, 0.85, 5.55, 11.63, 1.25, fill=REDWA, line=REDBD, line_w=1, radius=0.05)
text(s, 1.15, 5.76, 11.0, 0.5,
     [[R("しかもこれは、一過性のブームではない。", 15.5, INK, True)]])
text(s, 1.15, 6.28, 11.0, 0.45,
     [[R("各社の受注残・設備投資計画・売上ガイダンスに裏打ちされた、数年〜十数年続く", 12.5, BODY),
       R("構造需要", 12.5, RED, True), R("だ。", 12.5, BODY)]])

# =========================================================
# SLIDE 3 — HOW BIG（グローバル）：世界のAI投資の“規模”
# =========================================================
s = slide()
header(s, "2", "AI投資は、桁が違う。国家予算級のマネーが動いている。", "① グローバル：AI市場の盛り上がりは、どれくらいの規模か")

macro = [
    ("約110兆円", "Big Tech 4社のAI投資（2026年・年間）",
     "MS・Google・Amazon・Metaの設備投資合計 約7,250億ドル。前年比 +77%。ほぼ全額がAIインフラ。"),
    ("約780兆円", "世界のAI DC投資（2030年までの累計）",
     "生成AI対応DCへの設備投資は2030年までに約5.2兆ドル（総DC投資は約6.7兆ドル）に達する見込み（McKinsey）。"),
    ("日本1国分", "世界のDC電力需要（2030年）",
     "世界のデータセンター/AIの電力需要は2030年に約1兆kWh超へ。日本の年間電力消費に匹敵（IEA）。"),
]
mw = 11.63 / 3
for i, (v, lab, b) in enumerate(macro):
    x = 0.85 + i * mw
    rect(s, x+0.06, 1.6, mw-0.12, 2.5, fill=WHITE, line=REDBD, line_w=1.3, radius=0.05)
    rect(s, x+0.06, 1.6, mw-0.12, 0.11, fill=RED)
    text(s, x+0.32, 1.9, mw-0.6, 0.9, [[R(v, 40, RED, True)]], anchor=MSO_ANCHOR.MIDDLE)
    text(s, x+0.32, 2.86, mw-0.6, 0.5, [[R(lab, 12.5, INK, True)]], line_spacing=1.1)
    text(s, x+0.32, 3.36, mw-0.62, 0.7, [[R(b, 9.8, BODY)]], line_spacing=1.16)

# 解釈バンド
rect(s, 0.85, 4.4, 11.63, 1.15, fill=CHARWA, line=CHAR, line_w=1.2, radius=0.05)
text(s, 1.15, 4.6, 11.0, 0.5,
     [[R("“AIが盛り上がっている”の正体は、", 16.5, INK, True),
       R("電力・建物・機械への史上最大級の設備投資", 16.5, RED, True),
       R("だ。", 16.5, INK, True)]])
text(s, 1.15, 5.14, 11.0, 0.4,
     [[R("ソフトウェアの話に見えて、その足元は徹底的に“物理”——発電所、変電所、建屋、冷却、半導体工場。", 12, BODY)]])

# 下段：だから製造業に効く（→次スライドで日本へ）
rect(s, 0.85, 5.78, 11.63, 1.05, fill=REDWA, line=REDBD, line_w=1, radius=0.05)
text(s, 1.15, 5.9, 11.0, 0.9,
     [[R("では、これは", 15, INK, True),
       R("日本", 15, RED, True), R("にどう効くのか？", 15, INK, True),
       R("——実は、国内市場も“桁”が変わっている。", 15, INK, True)]],
     anchor=MSO_ANCHOR.MIDDLE)

# =========================================================
# SLIDE 4 — 日本ではどうか（国内市場の規模）
# =========================================================
s = slide()
header(s, "3", "世界のAIマネーは、実額で日本にも流れ込んでいる。", "② 日本ではどうか — 国内市場も、桁が変わった")

jp = [
    ("4〜4.8兆円", "米クラウド大手の対日DC投資（〜2030年）",
     "AWS 約2.26兆円('27迄)・MS 約1.6兆円('29迄)・Oracle 約1.2兆円。合計で2030年までに300億ドル超（日経）。"),
    ("2.7→5.6兆円", "国内DCサービス市場（2023→2030年）",
     "IDC/JEITA見通し。DC“建設投資”だけでも2028年に年1兆円超へ拡大する見込み（IDC）。"),
    ("約3兆＋5兆円", "半導体：TSMC熊本＋ラピダス",
     "TSMC熊本 第1+2で約3兆円、ラピダス総投資 約5兆円。関連投資を含め国内は9兆円規模（日経）。"),
]
jw = 11.63 / 3
for i, (v, lab, b) in enumerate(jp):
    x = 0.85 + i * jw
    rect(s, x+0.06, 1.55, jw-0.12, 2.4, fill=WHITE, line=REDBD, line_w=1.3, radius=0.05)
    rect(s, x+0.06, 1.55, jw-0.12, 0.11, fill=RED)
    text(s, x+0.30, 1.82, jw-0.56, 0.85, [[R(v, 33, RED, True)]], anchor=MSO_ANCHOR.MIDDLE)
    text(s, x+0.30, 2.72, jw-0.58, 0.5, [[R(lab, 12, INK, True)]], line_spacing=1.1)
    text(s, x+0.30, 3.22, jw-0.6, 0.7, [[R(b, 9.6, BODY)]], line_spacing=1.16)

# 電力・送配電バンド（国内固有の物理制約）
rect(s, 0.85, 4.25, 11.63, 1.25, fill=CHARWA, line=CHAR, line_w=1.2, radius=0.05)
text(s, 1.15, 4.45, 11.0, 0.5,
     [[R("そして電気が足りない：", 15.5, INK, True),
       R("国内DC・半導体の新増設で、最大需要電力は10年で約13倍", 15.5, RED, True),
       R("（56万→715万kW／OCCTO）。", 14, BODY, False)]])
text(s, 1.15, 5.0, 11.0, 0.45,
     [[R("→ 発電所・変電所・送配電網の増強も国内で同時進行（大阪の変電所群＋首都圏66kV網に1,500億円超）。", 12, BODY)]])

# クロージング：国内の現場になる
rect(s, 0.85, 5.72, 11.63, 1.1, fill=REDWA, line=REDBD, line_w=1, radius=0.05)
text(s, 1.15, 5.72, 11.0, 1.1,
     [[R("これらは全て、", 15, INK, True),
       R("“国内での建設・据付・保守の現場”", 15, RED, True), R("になる。", 15, INK, True)],
      [R("＝ 我々の営業対象が、海外の話ではなく、日本国内で今まさに立ち上がっている。", 13, BODY)]],
     anchor=MSO_ANCHOR.MIDDLE, space_after=6, line_spacing=1.2)

# =========================================================
# SLIDE 5 — WHERE：その金が“どこ”に着地するか（バリューチェーン×企業数字）
# =========================================================
s = slide()
header(s, "4", "国内のAI投資は、製造業のこの現場に着地している。", "AI需要は、バリューチェーンのどこに効くのか")

flow = [
    ("01", "発電・電源", "電気をつくる", "受注残 5兆円超", "三菱重工 ガスタービン"),
    ("02", "送変電・電線", "電気を送る・変える", "生産能力2倍/予約数年先", "変圧器 各社・フジクラ"),
    ("03", "DC建設・設備工事", "箱を建てる", "空調工事 受注 +25.5%", "ダイダン・高砂熱学ほか"),
    ("04", "冷却・電源保護", "冷やす・止めない", "北米DC冷却 約13倍", "ダイキン 230→3,000億円"),
    ("05", "半導体", "計算する頭脳", "設備投資 +66%", "キオクシア・東京ｴﾚｸﾄﾛﾝ"),
]
fy, fh = 1.55, 0.86
by, bh = 2.52, 2.5
col0 = 1.95
cw = (12.48 - col0) / 5
# start cap
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
    # body card
    bx = cx + 0.05; bw = cw - 0.10
    rect(s, bx, by, bw, bh, fill=WHITE, line=LINE, line_w=1, radius=0.05)
    pad = 0.13
    text(s, bx+pad, by+0.14, bw-2*pad, 0.3, [[R(sb, 11.5, INK, True)]])
    rect(s, bx+pad, by+0.62, bw-2*pad, 0.01, fill=REDBD)
    text(s, bx+pad, by+0.74, bw-2*pad, 0.7, [[R(sv, 13.5, RED, True)]], line_spacing=1.04)
    text(s, bx+pad, by+1.62, bw-2*pad, 0.24, [[R("代表企業", 7, MUTE, True, FONT_M)]])
    text(s, bx+pad, by+1.84, bw-2*pad, 0.6, [[R(chips, 9.2, INK, True)]], line_spacing=1.16)

# support band
rect(s, 0.85, 5.2, 11.63, 0.62, fill=REDWA, line=REDBD, line_w=1, radius=0.05)
rect(s, 0.85, 5.2, 2.35, 0.62, fill=RED, radius=0.05)
text(s, 0.95, 5.2, 2.2, 0.62,
     [[R("さらに基盤層で", 8.5, RGBColor(0xFF,0xE2,0xE5), False)],
      [R("部品・材料・ガス", 11, WHITE, True)]],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, space_after=0, line_spacing=1.05)
text(s, 3.4, 5.2, 8.9, 0.62,
     [[R("村田（MLCC・AIサーバ1台に約3万個）／SMC・アドバンテスト・イビデン・日本酸素 …", 11, INK, True)]],
     anchor=MSO_ANCHOR.MIDDLE)

# closing line
rect(s, 0.85, 6.0, 11.63, 0.85, fill=CHARWA, line=CHAR, line_w=1.2, radius=0.05)
text(s, 1.15, 6.0, 11.0, 0.85,
     [[R("川上（発電）から川下（半導体）まで全段が同時に増収増益。", 14.5, INK, True),
       R("＝ 攻めどころは1つではなく、チェーン全体にある。", 14.5, RED, True)]],
     anchor=MSO_ANCHOR.MIDDLE)

# =========================================================
# SLIDE 5 — WHAT VALUE：どの業界に、どんな価値を届けるか
# =========================================================
s = slide()
header(s, "5", "建設で磨いた“現場管理”を、そのまま製造業のフィールドへ。", "我々は、どの業界に・どんな価値を提供するのか")

# 2モデル
my = 1.55
rect(s, 0.85, my, 5.75, 1.35, fill=CHARWA, line=CHAR, line_w=1.2, radius=0.08)
text(s, 1.08, my+0.18, 5.35, 1.0,
     [[R("A  発注者として", 14, CHAR, True)],
      [R("設備投資を“建てる側”のプロジェクト管理", 11.5, INK, True)],
      [R("工場の新増設・生産ライン更新・自社DC/電源・冷却設備の導入", 10.5, BODY)]],
     space_after=4, line_spacing=1.14)
rect(s, 6.73, my, 5.75, 1.35, fill=REDWA, line=RED, line_w=1.2, radius=0.08)
text(s, 6.96, my+0.18, 5.35, 1.0,
     [[R("B  請負として", 14, RED, True)],
      [R("据付・試運転・保守を“納める側”の現場管理", 11.5, INK, True)],
      [R("発電機・変圧器・空調・液冷・UPS・半導体装置のフィールド現場（本命）", 10.5, BODY)]],
     space_after=4, line_spacing=1.14)

# どの業界（優先ターゲット）
rect(s, 0.85, 3.15, 11.63, 1.35, fill=WHITE, line=LINE, line_w=1, radius=0.05)
text(s, 1.1, 3.3, 11.2, 0.3, [[R("価値を届ける主要業界", 10.5, RED, True, FONT_M, 1),
     R("　★＝最優先", 8.5, MUTE, False, FONT_M)]])
tgts = [("重電（変圧器・発電機）", True), ("DC建設サブコン（電気・空調）", True),
        ("産業冷却・液冷・UPS", True), ("半導体・製造装置", False)]
tw = 11.63 / 4
for i, (t, top) in enumerate(tgts):
    x = 0.85 + i * tw
    if i > 0:
        rect(s, x, 3.7, 0.012, 0.65, fill=LINE)
    lab = "★ 最優先" if top else "有望"
    lcol = RED if top else CHAR
    text(s, x+0.24, 3.66, tw-0.36, 0.7,
         [[R(lab, 8.5, lcol, True, FONT_M)],
          [R(t, 12, INK, True)]], space_after=3, line_spacing=1.12)

# 相手の痛み → 我々の価値
rect(s, 0.85, 4.72, 11.63, 1.0, fill=REDWA, line=REDBD, line_w=1, radius=0.05)
text(s, 1.15, 4.72, 11.0, 1.0,
     [[R("相手の痛み：", 14, RED, True),
       R("案件は激増、でも人は増えない。紙・Excel・電話では、増えた据付／試運転／保守現場をさばけない。", 14, INK, True)]],
     anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.2)
rect(s, 0.85, 5.9, 11.63, 0.92, fill=CHARWA, line=CHAR, line_w=1.2, radius=0.05)
text(s, 1.15, 5.9, 11.0, 0.92,
     [[R("提供価値：", 14, CHAR, True),
       R("多拠点・多職種の現場を1つにつなぎ、進捗・図面・検査・写真・報告を可視化する。", 14, INK, True),
       R("——建設で実証済みの我々の主戦力。", 14, RED, True)]],
     anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.2)

# =========================================================
# SLIDE 6 — クロージング（モチベーション）
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
    ("追い風は本物", "受注残・設備投資計画・売上ガイダンスに裏打ちされた、数年続く構造需要。思惑ではない。"),
    ("武器は完成済み", "建設で磨いた現場管理は、製造業の据付・保守フィールドに“そのまま効く”最強の武器。"),
    ("市場は今まさに開いた", "川上から川下まで、狙える現場がチェーン全体に生まれている。先に動いた者が獲る。"),
]
py = 3.95
pw = (11.63 - 0.6) / 3
for i, (h, b) in enumerate(pts):
    x = 0.85 + i * (pw + 0.3)
    rect(s, x, py, pw, 1.65, fill=RGBColor(0x24,0x28,0x2E), line=RGBColor(0x3A,0x40,0x48), line_w=1, radius=0.06)
    rect(s, x, py, pw, 0.09, fill=RED)
    text(s, x+0.26, py+0.28, pw-0.52, 1.3,
         [[R(h, 15, WHITE, True)],
          [R(b, 10.8, RGBColor(0xC7,0xD0,0xDC))]], space_after=8, line_spacing=1.2)

rect(s, 0.85, 5.98, 11.63, 0.86, fill=RED, radius=0.06)
text(s, 0.85, 5.98, 11.63, 0.86,
     [[R("増える現場を、100人でつかみにいこう。", 21, WHITE, True)]],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

out = Path(__file__).resolve().parent.parent / "output" / "andpad_soukai_ai_deck.pptx"
prs.save(str(out))
print("saved:", out, "/ slides:", len(prs.slides._sldIdLst))
