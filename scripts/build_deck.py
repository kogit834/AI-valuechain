"""ANDPAD 製造業事業部 向けプレゼン資料（4枚, 16:9）を生成する。
出力: output/andpad_ai_opportunity_deck.pptx
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pathlib import Path

# ---- palette ----
INK   = RGBColor(0x13,0x20,0x38)
INKSF = RGBColor(0x4C,0x5B,0x74)
FAINT = RGBColor(0x7C,0x8A,0xA1)
PAPER = RGBColor(0xF6,0xF8,0xFB)
CARD  = RGBColor(0xFF,0xFF,0xFF)
LINE  = RGBColor(0xD3,0xDB,0xE6)
OWNER = RGBColor(0xB0,0x60,0x2A)   # 発注者
OWNBG = RGBColor(0xF3,0xE6,0xDA)
VEND  = RGBColor(0x0F,0x7A,0x83)   # 請負
VENBG = RGBColor(0xDD,0xEC,0xEE)
WHITE = RGBColor(0xFF,0xFF,0xFF)

FONT   = "Noto Sans JP"
FONT_M = "Consolas"

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]
SW, SH = prs.slide_width, prs.slide_height


def slide():
    s = prs.slides.add_slide(BLANK)
    r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, SH)
    r.fill.solid(); r.fill.fore_color.rgb = PAPER; r.line.fill.background()
    r.shadow.inherit = False
    s.shapes._spTree.remove(r._element); s.shapes._spTree.insert(2, r._element)
    return s


def rect(s, x, y, w, h, fill=None, line=None, line_w=1.0, radius=None):
    shp_type = MSO_SHAPE.ROUNDED_RECTANGLE if radius is not None else MSO_SHAPE.RECTANGLE
    sh = s.shapes.add_shape(shp_type, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.shadow.inherit = False
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid(); sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line; sh.line.width = Pt(line_w)
    if radius is not None:
        try:
            sh.adjustments[0] = radius
        except Exception:
            pass
    return sh


def text(s, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
         space_after=4, line_spacing=1.06):
    """runs: list of paragraphs; each paragraph = list of (txt, size, color, bold, font, tracking)."""
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align; p.space_before = Pt(0); p.space_after = Pt(space_after)
        p.line_spacing = line_spacing
        for (txt, size, color, bold, font, tracking) in para:
            r = p.add_run(); r.text = txt
            r.font.size = Pt(size); r.font.bold = bold
            r.font.color.rgb = color; r.font.name = font
            # set east-asian font too
            rPr = r._r.get_or_add_rPr()
            ea = rPr.find(qn('a:ea'))
            if ea is None:
                ea = rPr.makeelement(qn('a:ea'), {}); rPr.append(ea)
            ea.set('typeface', font)
            if tracking:
                rPr.set('spc', str(int(tracking * 100)))
    return tb


def R(txt, size, color=INK, bold=False, font=FONT, tracking=0):
    return (txt, size, color, bold, font, tracking)


# =========================================================
# SLIDE 1 — 表紙
# =========================================================
s = slide()
rect(s, 0, 0, 0.32, SH.inches, fill=OWNER)
rect(s, 0.32, 0, 0.14, SH.inches, fill=VEND)
text(s, 1.15, 1.5, 11.0, 0.5,
     [[R("ANDPAD 製造業事業部  ·  AI DEMAND OPPORTUNITY", 12.5, OWNER, True, FONT_M, 2)]])
text(s, 1.1, 2.05, 11.3, 2.6,
     [[R("AI需要の拡大が生む、", 40, INK, True)],
      [R("製造業の現場管理", 40, INK, True), R("オポチュニティ", 40, OWNER, True)]],
     line_spacing=1.1)
text(s, 1.15, 4.35, 11.0, 0.8,
     [[R("「", 19, INKSF, False), R("発注者", 19, OWNER, True),
       R("」と「", 19, INKSF, False), R("請負", 19, VEND, True),
       R("」——ふたつの入り方で、増える現場をつかむ。", 19, INKSF, False)]])
# divider
rect(s, 1.17, 5.55, 4.2, 0.02, fill=LINE)
text(s, 1.15, 5.75, 11.0, 0.9,
     [[R("生成AI → データセンター建設ラッシュ → 電力・重電・冷却・半導体まで、", 13.5, INKSF)],
      [R("製造業のあらゆる領域で「設備投資」と「据付・保守の現場」が同時に急増している。", 13.5, INKSF)]],
     line_spacing=1.25)
text(s, 1.15, 6.95, 11.0, 0.4,
     [[R("2026-07  /  出典：AIインフラ・バリューチェーン調査（国内外128社・決算/IR一次情報）", 10.5, FAINT, False, FONT_M)]])

# =========================================================
# SLIDE 2 — なぜ今
# =========================================================
s = slide()
text(s, 0.9, 0.55, 11.5, 0.4, [[R("01", 13, OWNER, True, FONT_M, 2), R("   なぜ、いま製造業に追い風なのか", 13, FAINT, True, FONT_M, 1)]])
text(s, 0.85, 1.0, 11.6, 0.8, [[R("AI需要の連鎖が、製造業に投資と現場をもたらす。", 27, INK, True)]])

steps = [
    ("きっかけ", "AIは“電気と設備の塊”", "DC1棟で数万世帯分の電力。膨大な電源・冷却・建屋・半導体が要る。"),
    ("連鎖", "製造業に投資と案件が波及", "重電・電線・空調・UPS・半導体装置——各社が工場を増設し、機器の納入・据付案件が全国で立ち上がる。"),
    ("結果", "管理すべき“現場”が急増", "設備投資プロジェクトも、据付・試運転・アフター保守の現場も同時に増える。"),
]
cx, cw, gap = 0.85, 3.75, 0.30
cy, ch = 2.05, 1.95
for i, (k, h, b) in enumerate(steps):
    x = cx + i * (cw + gap)
    rect(s, x, cy, cw, ch, fill=CARD, line=LINE, line_w=1, radius=0.06)
    text(s, x+0.28, cy+0.24, cw-0.56, ch-0.4,
         [[R(k, 10.5, FAINT, True, FONT_M, 1)],
          [R(h, 15.5, INK, True)],
          [R(b, 11, INKSF)]], space_after=6, line_spacing=1.14)
    if i < 2:
        text(s, x+cw-0.02, cy+0.55, 0.34, 0.6, [[R("→", 22, OWNER, True, FONT_M)]], align=PP_ALIGN.CENTER)

# number band
by, bh = 4.35, 1.45
rect(s, 0.85, by, 11.63, bh, fill=INK, radius=0.05)
nums = [
    ("5兆円超", "ガスタービン受注残", "三菱重工"),
    ("2倍へ", "変圧器の生産能力", "ダイヘン・東芝ES"),
    ("約13倍", "北米DC冷却の売上計画", "ダイキン 230→3,000億円超"),
    ("+66%", "半導体の設備投資", "キオクシア"),
]
nw = 11.63 / 4
for i, (v, lab, src) in enumerate(nums):
    x = 0.85 + i * nw
    if i > 0:
        rect(s, x, by+0.28, 0.014, bh-0.56, fill=RGBColor(0x2C,0x3B,0x54))
    text(s, x+0.28, by+0.22, nw-0.4, bh-0.3,
         [[R(v, 25, WHITE, True)],
          [R(lab, 12, RGBColor(0xC7,0xD3,0xE4), True)],
          [R(src, 9.5, RGBColor(0x8A,0x9C,0xB8), False, FONT_M)]], space_after=3, line_spacing=1.05)

text(s, 0.9, 6.08, 11.6, 0.9,
     [[R("→ ", 14, OWNER, True), R("案件は急増、人手は増えない。", 15, INK, True),
       R(" その差を、現場管理サービスで埋める。", 15, INKSF, False)]])
text(s, 0.9, 6.62, 11.6, 0.5,
     [[R("紙・Excel・電話では、増えた据付/試運転/保守の現場をさばききれない。", 11.5, FAINT)]])

# =========================================================
# SLIDE 3 — 2モデル × 領域マップ
# =========================================================
s = slide()
text(s, 0.9, 0.5, 11.5, 0.4, [[R("02", 13, OWNER, True, FONT_M, 2), R("   どの領域で、どんなテーマが生まれるか", 13, FAINT, True, FONT_M, 1)]])
text(s, 0.85, 0.93, 11.6, 0.7, [[R("ふたつの活用モデルで、バリューチェーンを攻める。", 25, INK, True)]])

# two model chips
my = 1.8
rect(s, 0.85, my, 5.75, 0.92, fill=OWNBG, line=OWNER, line_w=1.2, radius=0.09)
text(s, 1.08, my+0.13, 5.3, 0.7,
     [[R("A  発注者", 13.5, OWNER, True), R("   自社の設備投資を、建てる側として管理", 11.5, INKSF, False)],
      [R("工場新設・増産、生産ライン更新、自社DC・電源/冷却設備の導入", 10.5, INKSF)]], space_after=3, line_spacing=1.1)
rect(s, 6.73, my, 5.75, 0.92, fill=VENBG, line=VEND, line_w=1.2, radius=0.09)
text(s, 6.96, my+0.13, 5.3, 0.7,
     [[R("B  請負", 13.5, VEND, True), R("   据付・試運転・保守を、納める側として管理", 11.5, INKSF, False)],
      [R("発電機・変圧器・空調・液冷・UPS・半導体装置のフィールド現場", 10.5, INKSF)]], space_after=3, line_spacing=1.1)

# table: 領域 | A発注者テーマ | B請負テーマ
rows = [
    ("領域", "Model A ・ 発注者テーマ", "Model B ・ 請負テーマ", None),
    ("01 発電・電源設備", "メーカーの工場新増設を管理", "非常用電源の設置・試運転・保守", "B"),
    ("02 受変電・送配電・電線", "変圧器工場の増産・変電所建設", "機器の据付・更新工事・保守", "AB"),
    ("03 DC建設・設備工事", "自社拠点を建てる事業者の管理", "機器を納める製造業への横展開接点", "B"),
    ("04 冷却・電源保護", "空調メーカーの増産投資", "液冷・UPS・免震の据付〜アフター（最有力）", "B"),
    ("05 半導体", "巨大ファブ・クリーンルーム新設", "製造装置の据付・立ち上げ・保守", "AB"),
]
ty = 3.0
rh = 0.62
col_x = [0.85, 3.75, 8.1]
col_w = [2.9, 4.35, 4.38]
for ri, (c0, c1, c2, hot) in enumerate(rows):
    y = ty + ri * rh
    header = (ri == 0)
    if header:
        rect(s, col_x[0], y, sum(col_w), rh, fill=INK)
    elif ri % 2 == 0:
        rect(s, col_x[0], y, sum(col_w), rh, fill=RGBColor(0xEE,0xF2,0xF7))
    # cells
    def cell(cx, cw, txt, col, bold, size=11.5, font=FONT):
        text(s, cx+0.16, y, cw-0.3, rh, [[R(txt, size, col, bold, font)]],
             anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.05)
    if header:
        cell(col_x[0], col_w[0], c0, WHITE, True, 11, FONT_M)
        cell(col_x[1], col_w[1], c1, RGBColor(0xE9,0xC9,0xB2), True, 11, FONT_M)
        cell(col_x[2], col_w[2], c2, RGBColor(0xBF,0xE0,0xE3), True, 11, FONT_M)
    else:
        cell(col_x[0], col_w[0], c0, INK, True, 11.5)
        cell(col_x[1], col_w[1], c1, INKSF, False)
        cell(col_x[2], col_w[2], c2, INKSF, False)
# grid lines
for ri in range(len(rows)+1):
    rect(s, col_x[0], ty+ri*rh, sum(col_w), 0.012, fill=LINE)
rect(s, col_x[1]-0.0, ty, 0.012, rh*len(rows), fill=LINE)
rect(s, col_x[2]-0.0, ty, 0.012, rh*len(rows), fill=LINE)

text(s, 0.9, ty+rh*len(rows)+0.16, 11.6, 0.5,
     [[R("代表的な国内プレーヤー：", 10.5, FAINT, True, FONT_M),
       R(" 三菱重工 / 日立 / 三菱電機 / ダイヘン / フジクラ / きんでん・高砂熱学 / ダイキン・GSユアサ / 東京エレクトロン・ディスコ・キオクシア", 10.5, INKSF)]])

# =========================================================
# SLIDE 4 — どこから攻めるか
# =========================================================
s = slide()
text(s, 0.9, 0.55, 11.5, 0.4, [[R("03", 13, OWNER, True, FONT_M, 2), R("   どこから攻めるか — 狙い目と次の一歩", 13, FAINT, True, FONT_M, 1)]])
text(s, 0.85, 1.0, 11.6, 0.7, [[R("据付・保守は「請負」で、工場増設は「発注者」で。", 25, INK, True)]])

# priority matrix (left)
mx = [
    ("領域", "A 発注者", "B 請負"),
    ("01 発電・電源", "○", "◎"),
    ("02 受変電・送配電", "◎", "◎"),
    ("03 DC建設・設備", "○", "◎"),
    ("04 冷却・電源保護", "○", "◎"),
    ("05 半導体", "◎", "◎"),
]
mxx, mxy = 0.85, 2.1
cwid = [3.0, 1.5, 1.5]
mrh = 0.56
for ri, (c0, c1, c2) in enumerate(mx):
    y = mxy + ri*mrh
    header = ri == 0
    if header:
        rect(s, mxx, y, sum(cwid), mrh, fill=INK)
    elif ri % 2 == 0:
        rect(s, mxx, y, sum(cwid), mrh, fill=RGBColor(0xEE,0xF2,0xF7))
    text(s, mxx+0.16, y, cwid[0]-0.2, mrh,
         [[R(c0, 11 if header else 12, (WHITE if header else INK), True, (FONT_M if header else FONT))]],
         anchor=MSO_ANCHOR.MIDDLE)
    text(s, mxx+cwid[0], y, cwid[1], mrh,
         [[R(c1, 11 if header else 15, (RGBColor(0xE9,0xC9,0xB2) if header else OWNER), True, (FONT_M if header else FONT))]],
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, mxx+cwid[0]+cwid[1], y, cwid[2], mrh,
         [[R(c2, 11 if header else 15, (RGBColor(0xBF,0xE0,0xE3) if header else VEND), True, (FONT_M if header else FONT))]],
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
for ri in range(len(mx)+1):
    rect(s, mxx, mxy+ri*mrh, sum(cwid), 0.012, fill=LINE)
text(s, mxx, mxy+mrh*len(mx)+0.12, 6.0, 0.4,
     [[R("◎ = 相性が特に高い / ○ = 有望", 10, FAINT, False, FONT_M)]])

# conclusions (right)
concl = [
    (OWNER, "最有力の入口", "据付・保守は「請負 × フィールド」", "冷却・UPS・半導体装置など、設置→試運転→保守が命の機器メーカー。Model Bの本命。"),
    (VEND, "大型の伸びしろ", "工場増設ラッシュは「発注者」", "変圧器・半導体の増産投資は桁違い。建てる側のプロジェクト管理としてModel Aで。"),
    (OWNER, "横展開の型", "建設接点から機器メーカーへ", "DC建設の設備工事はANDPADの中核。そこに機器を納める製造業へ請負活用を広げる。"),
]
rx, rw = 7.35, 5.15
ry, rhh, rgap = 2.1, 1.15, 0.18
for i, (col, k, h, b) in enumerate(concl):
    y = ry + i*(rhh+rgap)
    rect(s, rx, y, rw, rhh, fill=CARD, line=LINE, line_w=1, radius=0.07)
    rect(s, rx, y, 0.09, rhh, fill=col)
    text(s, rx+0.3, y+0.15, rw-0.5, rhh-0.25,
         [[R(k+"　", 10, FAINT, True, FONT_M, 1), R(h, 14.5, INK, True)],
          [R(b, 10.5, INKSF)]], space_after=3, line_spacing=1.12)

text(s, 0.85, 6.75, 11.6, 0.5,
     [[R("次の一歩：", 12.5, OWNER, True),
       R(" 領域04（冷却・UPS）と領域05（半導体装置）の機器メーカーを起点に、「据付・試運転・アフター現場」の課題ヒアリングから着手する。", 12.5, INK)]])

out = Path(__file__).resolve().parent.parent / "output" / "andpad_ai_opportunity_deck.pptx"
prs.save(str(out))
print("saved:", out, "/ slides:", len(prs.slides._sldIdLst))
