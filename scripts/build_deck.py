"""ANDPAD 製造業事業部 向けプレゼン資料（本編4枚＋企業別参考）を生成する。
配色: 白 × ANDPADレッド × グレー。 出力: output/andpad_ai_opportunity_deck.pptx
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
CHAR   = RGBColor(0x45,0x4B,0x54)   # charcoal (model A)
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
    """ANDPAD風: 赤い角丸タイル + 見出し。"""
    rect(s, 0.85, 0.5, 0.52, 0.52, fill=RED, radius=0.18)
    text(s, 0.85, 0.5, 0.52, 0.52, [[R(tile, 15, WHITE, True, FONT_M)]],
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    if eyebrow:
        text(s, 1.55, 0.5, 11.0, 0.28, [[R(eyebrow, 11, RED, True, FONT_M, 1.5)]])
        text(s, 1.53, 0.74, 11.2, 0.5, [[R(title, 21, INK, True)]])
    else:
        text(s, 1.55, 0.5, 11.0, 0.55, [[R(title, 23, INK, True)]],
             anchor=MSO_ANCHOR.MIDDLE)
    rect(s, 0.85, 1.28, 11.63, 0.02, fill=LINE)


# =========================================================
# SLIDE 1 — 表紙
# =========================================================
s = slide()
rect(s, 0, 0, SW.inches, SH.inches, fill=WHITE)
rect(s, 0, 0, 0.30, SH.inches, fill=RED)
rect(s, 0.85, 1.42, 0.62, 0.62, fill=RED, radius=0.18)
text(s, 0.85, 1.42, 0.62, 0.62, [[R("AI", 18, WHITE, True, FONT_M)]],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text(s, 1.68, 1.5, 10.6, 0.5,
     [[R("ANDPAD 製造業事業部  ·  AI DEMAND OPPORTUNITY", 12.5, RED, True, FONT_M, 2)]],
     anchor=MSO_ANCHOR.MIDDLE)
text(s, 0.83, 2.35, 11.6, 2.3,
     [[R("AI需要の拡大が生む、", 40, INK, True)],
      [R("製造業の現場管理", 40, INK, True), R("オポチュニティ", 40, RED, True)]],
     line_spacing=1.1)
text(s, 0.86, 4.55, 11.2, 0.7,
     [[R("「", 19, BODY, False), R("発注者", 19, CHAR, True),
       R("」と「", 19, BODY, False), R("請負", 19, RED, True),
       R("」——ふたつの入り方で、増える現場をつかむ。", 19, BODY, False)]])
rect(s, 0.88, 5.6, 4.2, 0.03, fill=RED)
text(s, 0.86, 5.8, 11.2, 0.95,
     [[R("生成AI → データセンター建設ラッシュ → 電力・重電・冷却・半導体まで、", 13.5, BODY)],
      [R("製造業のあらゆる領域で「設備投資」と「据付・保守の現場」が同時に急増している。", 13.5, BODY)]],
     line_spacing=1.28)
text(s, 0.86, 7.0, 11.2, 0.35,
     [[R("2026-07  /  出典：AIインフラ・バリューチェーン調査（国内外128社・決算/IR一次情報）", 10.5, MUTE, False, FONT_M)]])

# =========================================================
# SLIDE 2 — バリューチェーン全体像（横フロー）
# =========================================================
s = slide()
header(s, "1", "電気をつくる川上から、計算する半導体の川下まで。", "AIインフラのバリューチェーン（価値の流れ）")

#  no, 段階名, 一言, tag, 数値, 出典, [代表企業], desc
flow = [
    ("01", "発電・電源", "電気をつくる", "B", "受注残 5兆円超", "三菱重工",
     "三菱重工・IHI・川崎重工・ヤンマー・デンヨー", "電力が全ての起点。ガスタービンや非常用発電機の需要が急増。"),
    ("02", "送変電・電線", "電気を送る・変える", "AB", "変圧器 生産能力2倍", "ダイヘン・東芝ES",
     "日立・三菱電機・ダイヘン・フジクラ・電力4社", "高電圧で送りDC向けに変換。変圧器は世界的に品薄で受注は数年先まで。"),
    ("03", "DC建設・設備工事", "箱を建てる", "B", "空調工事 受注 +25.5%", "ダイダン",
     "鹿島・大林・きんでん・関電工・高砂熱学", "巨大な建屋を建設。電気・空調のサブコンが主役で、営業の本命。"),
    ("04", "冷却・電源保護", "冷やす・止めない", "B", "北米DC冷却 約13倍", "ダイキン",
     "ダイキン・GSユアサ・オイレス・エア・ウォーター", "液冷で冷やし、UPSで止めない。据付・保守が伴う機器の宝庫。"),
    ("05", "半導体", "計算する頭脳", "AB", "設備投資 +66%", "キオクシア",
     "東京エレクトロン・ディスコ・信越化学・キオクシア", "AIチップ本体は海外勢が強いが、装置・材料で日本が世界的シェア。"),
]

fy, fh = 1.5, 0.98         # chevron band
by, bh = 2.66, 3.84        # body cards
col0 = 2.02
cw = (10.62 - col0) / 5    # 5 columns

# start cap（起点）
rect(s, 0.85, fy, 1.02, fh, fill=CHAR, radius=0.09)
text(s, 0.85, fy, 1.02, fh,
     [[R("起点", 8, RGBColor(0xD7,0xDB,0xE0), True, FONT_M, 1)],
      [R("AI需要", 12, WHITE, True)], [R("の拡大", 9.5, WHITE, True)]],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, space_after=0, line_spacing=1.0)
# end cap（価値実現）
rect(s, 10.66, fy, 1.82, fh, fill=INK, radius=0.06)
text(s, 10.66, fy, 1.82, fh,
     [[R("価値実現", 8, RGBColor(0xC7,0xCF,0xDB), True, FONT_M, 1)],
      [R("AIサービス", 12.5, WHITE, True)], [R("稼働", 10, WHITE, True)]],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, space_after=0, line_spacing=1.0)

for i, (no, nm, sb, tag, sv, ss, chips, desc) in enumerate(flow):
    cx = col0 + i * cw
    # chevron（矢羽根、テキストは別ボックスで重ねて改行崩れを防ぐ）
    ch = s.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(cx-0.02), Inches(fy), Inches(cw+0.16), Inches(fh))
    ch.shadow.inherit = False
    ch.fill.solid(); ch.fill.fore_color.rgb = RED; ch.line.fill.background()
    ch.text_frame.text = ""
    text(s, cx-0.06, fy, cw+0.10, fh,
         [[R(no, 8, RGBColor(0xFF,0xC9,0xCE), True, FONT_M, 1)],
          [R(nm, 11, WHITE, True)]],
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, space_after=0, line_spacing=1.0)

    # body card
    bx = cx + 0.05; bw = cw - 0.10
    rect(s, bx, by, bw, bh, fill=WHITE, line=LINE, line_w=1, radius=0.045)
    tagtxt = {"A": "A 発注者", "B": "B 請負", "AB": "A＋B 両取り"}[tag]
    tagcol = CHAR if tag == "A" else RED
    pad = 0.14
    # tag + 一言
    text(s, bx+pad, by+0.14, bw-2*pad, 0.24, [[R(tagtxt, 8.5, tagcol, True, FONT_M)]])
    text(s, bx+pad, by+0.42, bw-2*pad, 0.3, [[R(sb, 11.5, INK, True)]])
    # desc
    text(s, bx+pad, by+0.82, bw-2*pad, 1.0, [[R(desc, 8.8, BODY)]], line_spacing=1.16)
    # stat block
    rect(s, bx+pad, by+1.86, bw-2*pad, 0.01, fill=REDBD)
    text(s, bx+pad, by+1.96, bw-2*pad, 0.66,
         [[R(sv, 12, RED, True)], [R(ss, 7.5, MUTE, False, FONT_M)]], space_after=0, line_spacing=1.06)
    # companies
    text(s, bx+pad, by+2.72, bw-2*pad, 0.22, [[R("代表企業（国内）", 7.5, MUTE, True, FONT_M)]])
    text(s, bx+pad, by+2.96, bw-2*pad, 0.82, [[R(chips, 9, INK, True)]], line_spacing=1.22)

# legend
rect(s, 0.85, 6.55, 11.63, 0.5, fill=PANEL, line=LINE, line_w=1, radius=0.06)
text(s, 1.1, 6.55, 11.2, 0.5,
     [[R("現場管理の入り方：", 10, MUTE, True, FONT_M),
       R(" A 発注者＝自社の設備投資を建てる側で管理", 10.5, CHAR, True),
       R("　／　", 10, MUTE, False),
       R("B 請負＝据付・試運転・保守を納める側で管理", 10.5, RED, True),
       R("　※各段階は代表企業のみ", 9, MUTE, False)]],
     anchor=MSO_ANCHOR.MIDDLE)
text(s, 0.9, 7.12, 11.5, 0.28,
     [[R("さらに川下へ波及：", 9, RED, True, FONT_M),
       R("半導体製造装置の部品（SMC・堀場・アドバンテスト等）／電子部品・パッケージ基板（村田・イビデン等）／産業ガスにも拡大 → 付録参照", 9, BODY)]])

# =========================================================
# SLIDE 3 — なぜ今
# =========================================================
s = slide()
header(s, "2", "AI需要の連鎖が、製造業に投資と現場をもたらす。", "なぜ、いま製造業に追い風なのか")

steps = [
    ("きっかけ", "AIは“電気と設備の塊”", "DC1棟で数万世帯分の電力。膨大な電源・冷却・建屋・半導体が要る。"),
    ("連鎖", "製造業に投資と案件が波及", "重電・電線・空調・UPS・半導体装置——各社が工場を増設し、機器の納入・据付案件が全国で立ち上がる。"),
    ("結果", "管理すべき“現場”が急増", "設備投資プロジェクトも、据付・試運転・アフター保守の現場も同時に増える。"),
]
cx, cw, gap, cy, ch = 0.85, 3.75, 0.30, 1.6, 1.85
for i, (k, h, b) in enumerate(steps):
    x = cx + i * (cw + gap)
    rect(s, x, cy, cw, ch, fill=WHITE, line=REDBD, line_w=1.2, radius=0.06)
    rect(s, x, cy, cw, 0.10, fill=RED, radius=0.0)
    text(s, x+0.28, cy+0.26, cw-0.56, ch-0.4,
         [[R(k, 10.5, RED, True, FONT_M, 1)],
          [R(h, 15.5, INK, True)],
          [R(b, 11, BODY)]], space_after=6, line_spacing=1.14)
    if i < 2:
        text(s, x+cw-0.03, cy+0.5, 0.36, 0.6, [[R("→", 22, RED, True, FONT_M)]], align=PP_ALIGN.CENTER)

by, bh = 3.85, 1.5
rect(s, 0.85, by, 11.63, bh, fill=PANEL, line=LINE, line_w=1, radius=0.04)
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
        rect(s, x, by+0.3, 0.014, bh-0.6, fill=LINE)
    text(s, x+0.30, by+0.24, nw-0.42, bh-0.3,
         [[R(v, 26, RED, True)],
          [R(lab, 12, INK, True)],
          [R(src, 9.5, MUTE, False, FONT_M)]], space_after=3, line_spacing=1.05)

rect(s, 0.85, 5.65, 11.63, 1.2, fill=REDWA, line=REDBD, line_w=1, radius=0.05)
text(s, 1.15, 5.86, 11.0, 0.5,
     [[R("案件は急増、人手は増えない。", 16, INK, True),
       R("　その差を、現場管理サービスで埋める。", 16, RED, True)]])
text(s, 1.15, 6.4, 11.0, 0.4,
     [[R("紙・Excel・電話では、増えた据付／試運転／保守の現場をさばききれない。", 12, BODY)]])

# =========================================================
# SLIDE 3 — 2モデル × 領域マップ
# =========================================================
s = slide()
header(s, "3", "ふたつの活用モデルで、バリューチェーンを攻める。", "どの領域で、どんなテーマが生まれるか")

my = 1.55
rect(s, 0.85, my, 5.75, 0.95, fill=CHARWA, line=CHAR, line_w=1.2, radius=0.09)
text(s, 1.08, my+0.15, 5.35, 0.7,
     [[R("A  発注者", 13.5, CHAR, True), R("　自社の設備投資を、建てる側として管理", 11, BODY, False)],
      [R("工場新設・増産、生産ライン更新、自社DC・電源/冷却設備の導入", 10.5, BODY)]], space_after=3, line_spacing=1.12)
rect(s, 6.73, my, 5.75, 0.95, fill=REDWA, line=RED, line_w=1.2, radius=0.09)
text(s, 6.96, my+0.15, 5.35, 0.7,
     [[R("B  請負", 13.5, RED, True), R("　据付・試運転・保守を、納める側として管理", 11, BODY, False)],
      [R("発電機・変圧器・空調・液冷・UPS・半導体装置のフィールド現場", 10.5, BODY)]], space_after=3, line_spacing=1.12)

rows = [
    ("領域", "Model A ・ 発注者テーマ", "Model B ・ 請負テーマ"),
    ("01 発電・電源設備", "メーカーの工場新増設を管理", "非常用電源の設置・試運転・保守"),
    ("02 受変電・送配電・電線", "変圧器工場の増産・変電所建設", "機器の据付・更新工事・保守"),
    ("03 DC建設・設備工事", "自社拠点を建てる事業者の管理", "機器を納める製造業への横展開接点"),
    ("04 冷却・電源保護", "空調メーカーの増産投資", "液冷・UPS・免震の据付〜アフター（最有力）"),
    ("05 半導体", "巨大ファブ・クリーンルーム新設", "製造装置の据付・立ち上げ・保守"),
]
ty, rh = 2.85, 0.62
col_x = [0.85, 3.75, 8.1]; col_w = [2.9, 4.35, 4.38]
for ri, (c0, c1, c2) in enumerate(rows):
    y = ty + ri * rh; head = (ri == 0)
    if head: rect(s, col_x[0], y, sum(col_w), rh, fill=RED)
    elif ri % 2 == 0: rect(s, col_x[0], y, sum(col_w), rh, fill=PANEL)
    def cell(cx, cw, txt, col, bold, size=11.5, font=FONT, align=PP_ALIGN.LEFT):
        text(s, cx+0.16, y, cw-0.3, rh, [[R(txt, size, col, bold, font)]],
             anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.05, align=align)
    if head:
        cell(col_x[0], col_w[0], c0, WHITE, True, 11, FONT_M)
        cell(col_x[1], col_w[1], c1, WHITE, True, 11, FONT_M)
        cell(col_x[2], col_w[2], c2, RGBColor(0xFF,0xDD,0xE1), True, 11, FONT_M)
    else:
        cell(col_x[0], col_w[0], c0, INK, True, 11.5)
        cell(col_x[1], col_w[1], c1, BODY, False)
        cell(col_x[2], col_w[2], c2, BODY, False)
for ri in range(len(rows)+1):
    rect(s, col_x[0], ty+ri*rh, sum(col_w), 0.012, fill=LINE)
rect(s, col_x[1], ty+rh, 0.012, rh*(len(rows)-1), fill=LINE)
rect(s, col_x[2], ty+rh, 0.012, rh*(len(rows)-1), fill=LINE)

text(s, 0.9, ty+rh*len(rows)+0.14, 11.6, 0.5,
     [[R("代表的な国内プレーヤー：", 10.5, RED, True, FONT_M),
       R(" 三菱重工 / 日立 / 三菱電機 / ダイヘン / フジクラ / きんでん・高砂熱学 / ダイキン・GSユアサ / 東京エレクトロン・ディスコ・キオクシア", 10.5, BODY)]])

# =========================================================
# SLIDE 4 — どこから攻めるか
# =========================================================
s = slide()
header(s, "4", "据付・保守は「請負」で、工場増設は「発注者」で。", "どこから攻めるか — 狙い目と次の一歩")

mx = [
    ("領域", "A 発注者", "B 請負"),
    ("01 発電・電源", "○", "◎"),
    ("02 受変電・送配電", "◎", "◎"),
    ("03 DC建設・設備", "○", "◎"),
    ("04 冷却・電源保護", "○", "◎"),
    ("05 半導体", "◎", "◎"),
]
mxx, mxy = 0.85, 1.7
cwid = [3.0, 1.5, 1.5]; mrh = 0.57
for ri, (c0, c1, c2) in enumerate(mx):
    y = mxy + ri*mrh; head = ri == 0
    if head: rect(s, mxx, y, sum(cwid), mrh, fill=RED)
    elif ri % 2 == 0: rect(s, mxx, y, sum(cwid), mrh, fill=PANEL)
    text(s, mxx+0.16, y, cwid[0]-0.2, mrh,
         [[R(c0, 11 if head else 12, (WHITE if head else INK), True, (FONT_M if head else FONT))]],
         anchor=MSO_ANCHOR.MIDDLE)
    text(s, mxx+cwid[0], y, cwid[1], mrh,
         [[R(c1, 11 if head else 15, (WHITE if head else CHAR), True, (FONT_M if head else FONT))]],
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, mxx+cwid[0]+cwid[1], y, cwid[2], mrh,
         [[R(c2, 11 if head else 15, (RGBColor(0xFF,0xDD,0xE1) if head else RED), True, (FONT_M if head else FONT))]],
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
for ri in range(len(mx)+1):
    rect(s, mxx, mxy+ri*mrh, sum(cwid), 0.012, fill=LINE)
text(s, mxx, mxy+mrh*len(mx)+0.12, 6.0, 0.4,
     [[R("◎ = 相性が特に高い / ○ = 有望", 10, MUTE, False, FONT_M)]])

concl = [
    (RED, "最有力の入口", "据付・保守は「請負 × フィールド」", "冷却・UPS・半導体装置など、設置→試運転→保守が命の機器メーカー。Model Bの本命。"),
    (CHAR, "大型の伸びしろ", "工場増設ラッシュは「発注者」", "変圧器・半導体の増産投資は桁違い。建てる側のプロジェクト管理としてModel Aで。"),
    (RED, "横展開の型", "建設接点から機器メーカーへ", "DC建設の設備工事はANDPADの中核。そこに機器を納める製造業へ請負活用を広げる。"),
]
rx, rw, ry, rhh, rgap = 7.35, 5.15, 1.7, 1.15, 0.18
for i, (col, k, h, b) in enumerate(concl):
    y = ry + i*(rhh+rgap)
    rect(s, rx, y, rw, rhh, fill=WHITE, line=LINE, line_w=1, radius=0.07)
    rect(s, rx, y, 0.10, rhh, fill=col)
    text(s, rx+0.3, y+0.16, rw-0.5, rhh-0.25,
         [[R(k+"　", 10, (RED if col==RED else CHAR), True, FONT_M, 1), R(h, 14.5, INK, True)],
          [R(b, 10.5, BODY)]], space_after=3, line_spacing=1.12)

rect(s, 0.85, 6.7, 11.63, 0.62, fill=REDWA, line=REDBD, line_w=1, radius=0.06)
text(s, 1.1, 6.7, 11.1, 0.62,
     [[R("次の一歩：", 12.5, RED, True),
       R(" 領域04（冷却・UPS）と領域05（半導体装置）の機器メーカーを起点に、「据付・試運転・アフター現場」の課題ヒアリングから着手する。", 12.5, INK)]],
     anchor=MSO_ANCHOR.MIDDLE)

# =========================================================
# 参考（巻末） — 企業別 AI需要の影響
# =========================================================
def appendix_divider():
    s = slide()
    rect(s, 0, 0, SW.inches, SH.inches, fill=INK)
    rect(s, 0, 0, 0.30, SH.inches, fill=RED)
    rect(s, 0.9, 3.0, 0.62, 0.62, fill=RED, radius=0.18)
    text(s, 0.9, 3.0, 0.62, 0.62, [[R("参", 17, WHITE, True)]], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, 1.75, 3.0, 10.5, 0.7, [[R("APPENDIX", 13, RED, True, FONT_M, 3)]], anchor=MSO_ANCHOR.MIDDLE)
    text(s, 0.9, 3.85, 11.4, 1.2,
         [[R("参考：企業別 AI需要の影響", 34, WHITE, True)]])
    text(s, 0.92, 4.95, 11.2, 0.6,
         [[R("バリューチェーンの領域ごとに、国内主要企業がAI需要をどう受けるかを整理。", 13.5, RGBColor(0xC7,0xD0,0xDC))]])
    text(s, 0.92, 5.5, 11.2, 0.5,
         [[R("数値は各社の決算・IR等の直近開示にもとづく（対象期間・単位は出典元を要確認）。", 11, RGBColor(0x9A,0xA6,0xB6), False, FONT_M)]])


def detail_slide(tile, eyebrow, title, model, rows):
    """model: 'A'/'B'/'AB' でヘッダー帯色。rows: [(name, code, desc, conf)]"""
    s = slide()
    header(s, tile, title, eyebrow)
    # model tag
    tagcol = RED if model in ("B",) else (CHAR if model == "A" else RED)
    tagtxt = {"A":"Model A ・ 発注者が主","B":"Model B ・ 請負が主","AB":"Model A / B 両取り"}[model]
    text(s, 0.9, 1.34, 11.5, 0.3, [[R(tagtxt, 10.5, tagcol, True, FONT_M, 1)]])
    ty = 1.78
    n = len(rows)
    avail = 6.95 - ty
    rh = min(0.72, avail / n)
    cx = [0.85, 3.35, 4.35]; cw = [2.5, 1.0, 7.4]  # name, code, desc  (conf inline)
    fsz = 11 if rh >= 0.62 else 10
    for ri, (name, code, desc, conf) in enumerate(rows):
        y = ty + ri*rh
        if ri % 2 == 1:
            rect(s, cx[0], y, sum(cw), rh, fill=PANEL)
        # confidence dot color
        cc = RED if conf == "高" else (CHAR if conf == "中" else MUTE)
        text(s, cx[0]+0.12, y, cw[0]-0.2, rh,
             [[R(name, fsz+0.5, INK, True)]], anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.02)
        text(s, cx[1]+0.05, y, cw[1]-0.1, rh,
             [[R(code, fsz-0.5, MUTE, False, FONT_M)]], anchor=MSO_ANCHOR.MIDDLE)
        text(s, cx[2]+0.12, y, cw[2]-0.9, rh,
             [[R(desc, fsz, BODY)]], anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.08)
        text(s, cx[2]+cw[2]-0.78, y, 0.7, rh,
             [[R("確度 ", fsz-2, MUTE, False, FONT_M), R(conf, fsz-1, cc, True)]],
             anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.RIGHT)
    for ri in range(n+1):
        rect(s, cx[0], ty+ri*rh, sum(cw), 0.01, fill=LINE)


appendix_divider()

detail_slide("参", "領域01 ・ 発電・電源設備", "電気をつくる — 重工・エンジン・発電機メーカー", "B", [
    ("三菱重工業", "7011", "北米中心にガスタービン受注が急拡大。大型ガスタービンの生産能力を2倍へ増強。", "高"),
    ("川崎重工業", "7012", "ガスタービン・ガスエンジンが増収増益。非常用「PUシリーズ」8,000台超の納入実績。", "中"),
    ("IHI", "7013", "子会社IHI原動機がDC向け非常用発電装置の引き合い増で、発電機を高出力化。", "中"),
    ("デンヨー", "6517", "非常用発電機の国内需要が堅調、米国向け輸出も好調で増収増益。", "中"),
    ("ヤンマーHD", "非上場", "DC向け大容量非常用発電機「GY175」を商品化、北九州に専用新工場（28年稼働）。", "中"),
])

detail_slide("参", "領域02a ・ 受変電・重電・送配電機器", "電気を変える・送る — 変圧器/GISメーカー", "AB", [
    ("日立製作所", "6501", "DC・半導体工場向け変圧器/送配電の受注が世界的に急拡大、米国に変圧器新工場。3年ぶり最高益。", "高"),
    ("三菱電機", "6503", "特別高圧変圧器（赤穂）の生産予約が数年先まで。DC/防衛で5年連続最高益。C-GISも増産。", "高"),
    ("東芝エネルギーシステムズ", "非上場", "送変電機器の増産投資を拡大（GIS等に約550億円、30年度に生産能力2倍超）。", "高"),
    ("富士電機", "6504", "DC・半導体向け特高変圧器とUPS・施設電源が伸長し5年連続最高益。生産体制を再編・増強。", "高"),
    ("明電舎", "6508", "系統増強を見据え沼津の変圧器工場に160億円投資、生産能力を約1.5倍に拡大。", "中"),
    ("ダイヘン", "6622", "大形変圧器の需要急増で三重に新工場、生産能力を29年度までに2倍へ。", "高"),
    ("日新電機", "非上場", "レベニューキャップ・半導体工場向けでGIS需要が急増、国内にGIS第二工場を新設。", "中"),
])

detail_slide("参", "領域02b ・ 電線・ケーブル / 電力会社", "電気を送る — 光/電力ケーブルと送配電投資", "AB", [
    ("古河電気工業", "5801", "DC/AI向け高付加価値光ファイバの供給体制を強化、DC事業の営業利益を5年で8.5倍計画。", "高"),
    ("住友電気工業", "5802", "AI DC向け光デバイス/配線/コネクタ需要増で〜28年度に約1,000億円投資（光コネクタ7倍）。", "高"),
    ("フジクラ", "5803", "北米AI DC向け光ケーブル需要が爆発、情通事業が全社営業利益の8割超。最大3,000億円増産。", "高"),
    ("SWCC", "5805", "米AI DC向け超多心光ケーブル「e-Ribbon」需要が拡大、通信・部品事業を牽引。", "高"),
    ("東京電力HD", "9501", "千葉印西等のDC集積で電力需要急増、送電網増強へ〜27年度に4,700億円投資。", "中"),
    ("関西電力", "9503", "DC新増設で電力需要急増、送配電子会社が変電所・送電線に1,500億円超を追加投資。", "中"),
    ("中部電力", "9502", "DC・半導体の需要増を見据え、連系線・系統増強など次世代電力網に重点投資。", "中"),
    ("九州電力", "9508", "TSMC進出等でDC/半導体需要増、送配電5カ年投資を約6,500億円（+1割）に。", "高"),
])

detail_slide("参", "領域03a ・ データセンター建設ゼネコン", "箱を建てる — 元請けゼネコン", "A", [
    ("鹿島建設", "1812", "DC等大型工事の施工量増で建築が急拡大、建設業初の単体売上高3兆円を突破。", "中"),
    ("大成建設", "1801", "中期経営計画でDCを注力分野に明記、運営ノウハウを商品企画へ反映。", "中"),
    ("清水建設", "1803", "DCを建築の主要需要分野に位置づけ、適正工期・価格で選別受注を強化。", "中"),
    ("大林組", "1802", "米DC/半導体工場専門の建設会社GCONを買収、北米DC受注拡大を狙う。", "高"),
    ("竹中工務店", "非上場", "業界初のDC特化型設計支援ツールを自社開発、提案スピードを強化。", "中"),
    ("戸田建設", "1860", "関西けいはんなの大型DCキャンパス「KIX01A」（総電力80MW）を受注。", "中"),
    ("西松建設", "1820", "中計でDCを注力分野に明記、国内建築の高収益体質への転換を推進。", "中"),
])

detail_slide("参", "領域03b ・ 電気・空調 設備工事（サブコン）", "箱を建てる — 設備工事の専門会社（営業本命）", "B", [
    ("きんでん", "1944", "DC/物流向け一般電気工事が急拡大、一般電気工事の受注高が+21.0%（3,194億円）。", "高"),
    ("関電工", "1942", "DC・半導体工場向け電気設備工事とビル再開発で、26/3期は連続最高益の見通し。", "高"),
    ("クラフティア(旧九電工)", "1959", "DC関連工事中心に受注を強化、中間期受注高が+20.8%（2,752億円）。", "高"),
    ("ダイダン", "1980", "DC新設で大型空調工事が急増、受注工事高が+25.5%（3,531億円）。", "高"),
    ("高砂熱学工業", "1969", "半導体/AI DC向け高難度空調を独占的に受注、4期連続最高益（営業+47.3%）。", "高"),
    ("新日本空調", "1952", "半導体CR・DC向け「産業」が完成工事高の57%超、受注・繰越とも過去最高。", "高"),
    ("三機工業", "1961", "DC向け大容量空調「L-LAC」で大型受注が継続、営業利益が過去最高。", "高"),
    ("住友電設", "1949", "国内外の大型DC案件と再エネ工事が好調、営業利益+42.5%。", "高"),
    ("日本電設工業", "1950", "民間DC投資拡大で一般電気工事が伸長、受注・売上とも過去最高。", "中"),
])

detail_slide("参", "領域04a ・ 冷却・空調 / 電源保護", "冷やす・止めない — 液冷/チラー/UPS", "B", [
    ("ダイキン工業", "6367", "北米DC冷却の売上を約230億→3,000億円超（30年度）へ、液冷企業を相次ぎ買収。", "高"),
    ("三菱重工サーマルシステムズ", "非上場", "国内ターボ冷凍機シェア約7割、半導体/DC向け受注増で増産。", "中"),
    ("荏原冷熱システム", "非上場", "北米DC向けターボ冷凍機・冷却塔がAI/DC拡大で好調。", "低"),
    ("東洋熱工業", "非上場", "DC向け高効率冷却ユニット「グリーンアイル」など独自技術で冷却需要を取り込む。", "中"),
    ("ジーエス・ユアサ", "6674", "DC/半導体工場向け非常用電池電源（UPS蓄電池）の大型受注で増収増益。", "高"),
    ("エア・ウォーター", "4088", "子会社Hitec中心の高出力UPSがDC/半導体向けでアジア・北米受注を拡大。", "高"),
])

detail_slide("参", "領域04b ・ 免震・防音・セキュリティ・建材", "守る — データセンターの特殊設備", "B", [
    ("オイレス工業", "6282", "都市型DC/物流向け免震装置のスペックインを重点化、大型案件の受注に注力。", "中"),
    ("日創グループ", "3440", "子会社製のDC向け非常用発電機用の防音筐体・消音ダクトの受注が拡大。", "高"),
    ("オプテックスグループ", "6914", "AI/DC投資拡大で、北米中心にDC等向け防犯センサーの受注・需要が拡大。", "高"),
    ("ヤマックス", "5285", "熊本の半導体集積で工場建設向けプレキャストが想定超、TSMC熊本のCR床材に採用。", "高"),
])

detail_slide("参", "領域05 ・ 半導体（装置・材料・メモリ）", "頭脳をつくる — 日本は装置・材料に世界的強み", "AB", [
    ("東京エレクトロン", "8035", "生成AI DC向けの微細化投資でロジック・DRAM向け装置売上が急拡大。", "高"),
    ("ディスコ", "6146", "GPU/HBM等の先端半導体需要で精密加工装置が拡大、6期連続で営業最高益。", "高"),
    ("SCREENホールディングス", "7735", "半導体微細化・チップレット向け需要で、27/3期の増収増益を見込む。", "低"),
    ("SUMCO", "3436", "AIサーバ/HBM向けシリコンウエハ需要拡大で、既存工場（伊万里）の増強・高精度化に投資。", "中"),
    ("信越化学工業", "4063", "シリコンウエハ・電子材料のAI用途比率が2割超、顧客の在庫積み増しで追加受注。", "中"),
    ("レゾナック", "4004", "先端パッケージング材料（TIM/NCF）のAI向け販売増で半導体材料の営業利益が過去最高。", "高"),
    ("キオクシア", "285A", "AI DC向けSSD需要で26年度の生産能力は完売、設備投資を前年比+66%へ。", "高"),
    ("ラピダス", "非上場", "2nmロジックの量産（27年度後半予定）へ資金を積み増し中、顧客受注確定は途上。", "低"),
])

detail_slide("参", "領域06a ・ 半導体製造装置 部品・サブシステム", "装置の“中身” — 流体制御・シール・搬送・計測（SMC等）", "B", [
    ("SMC", "6273", "空気圧制御で世界シェア約4割、半導体装置に多用。前工程投資回復で増収増益（売上8,425億円+6.4%）。", "中"),
    ("CKD", "6407", "空気圧・薬液用バルブ等の装置向け部品が主力。半導体関連は堅調でAI投資が追い風。", "中"),
    ("日本ピラー工業", "6490", "超純水・薬液用フッ素樹脂継手/シールで高シェア。27/3期に売上700億円へ大幅拡大を計画。", "高"),
    ("イーグル工業", "6486", "装置向けメカニカルシールが生成AIで回復、26/3期は営業利益+58.6%で過去最高。", "高"),
    ("堀場製作所", "6856", "マスフローコントローラ世界シェア約6割。AI/DC需要で通期見通しを上方修正（営業+28.2%）。", "高"),
    ("ローツェ", "6323", "搬送ロボット・EFEMで世界シェア25〜35%。Applied・TSMCが主要顧客で前工程需要が追い風。", "中"),
    ("アドバンテスト", "6857", "GPU/HBM向けテスタ需要が急拡大、26/3期は売上+44.7%・営業利益+118.8%で過去最高。", "高"),
    ("ダイフク", "6383", "クリーンルーム搬送(AMHS)で世界大手。先端半導体投資で受注+54.7%（Q1・過去最高）。", "高"),
    ("フェローテックHD", "6890", "装置向け真空シール・石英/セラミック部品が主力。生成AI・メモリ装置投資で増収増益。", "中"),
])

detail_slide("参", "領域06b ・ 電子部品・パッケージ基板・産業ガス", "AIサーバー/半導体の下流 — 基板・MLCC・特殊ガス", "AB", [
    ("イビデン", "4062", "生成AIサーバー向けICパッケージ基板が堅調。3年で5,000億円投資し基板能力を2.5倍へ。", "高"),
    ("新光電気工業", "6967", "GPU向けFC-BGA基板大手。AIパッケージング需要で成長投資（JIC傘下で非上場化）。", "中"),
    ("村田製作所", "6981", "MLCC世界首位。AIサーバー1台に約3万個必要で需要急拡大、26年4月に15〜35%値上げ。", "高"),
    ("太陽誘電", "6976", "AIサーバー向けMLCC（1台約2.8万個）が急拡大、26/3期は営業利益+91.2%・純利益5.4倍。", "高"),
    ("TDK", "6762", "AI DC向けHDD部品・パッシブ部品が牽引し26/3期は過去最高。パッシブ部品売上を約10倍へ。", "高"),
    ("日本酸素HD", "4091", "産業ガス国内首位・特殊ガス世界3強。AI半導体増産で特殊ガス供給が堅調、増収増益。", "中"),
    ("関東電化工業", "4047", "半導体エッチング用の含フッ素特殊ガスが主力。AI半導体増産で経常利益+47.1%、27/3期も拡大。", "中"),
    ("トリケミカル研究所", "4369", "High-k等の超高純度前駆体材料で高シェア。生成AI需要で売上+37.4%・営業+30.2%。", "高"),
])

out = Path(__file__).resolve().parent.parent / "output" / "andpad_ai_opportunity_deck.pptx"
prs.save(str(out))
print("saved:", out, "/ slides:", len(prs.slides._sldIdLst))
