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
# SLIDE 3 — トレンド：これまでの何倍か（時系列の勢い）
# =========================================================
s = slide()
header(s, "2", "AI投資は“大きい”だけじゃない。数年で“何倍”に跳ねている。",
       "規模より“勢い”——トレンドとして、どれだけ急か")

LT_RED = RGBColor(0xF4,0xB4,0xBD)   # 過去バー（淡赤）

def vbars(s, x0, base_y, area_w, max_h, values, last_label):
    """左→右に伸びる縦棒。最新バーのみ濃赤＋値ラベル。単一系列・magnitude。"""
    n = len(values); mx = max(values)
    slot = area_w / n; bw = min(0.46, slot * 0.62)
    rect(s, x0, base_y, area_w, 0.014, fill=LINE)  # baseline（控えめ）
    for i, v in enumerate(values):
        h = max_h * v / mx
        bxx = x0 + i * slot + (slot - bw) / 2
        col = RED if i == n - 1 else LT_RED
        rect(s, bxx, base_y - h, bw, h, fill=col, radius=0.12)
        if i == n - 1:
            text(s, bxx - 0.5, base_y - h - 0.28, bw + 1.0, 0.26,
                 [[R(last_label, 10, RED, True, FONT_M)]], align=PP_ALIGN.CENTER)

# rows: (what, note, values, last_label, mult, period, src)
rows = [
    ("AIチップの需要", "NVIDIA データセンター売上", [15, 48, 115, 196], "約2,000億ドル",
     "約13倍", "3年（FY23→FY26）", "NVIDIA IR"),
    ("AIインフラ投資", "Big Tech 4社の設備投資（年間）", [180, 230, 410, 725], "7,250億ドル",
     "約4倍", "3年（'23→'26）", "各社IR/報道"),
    ("国内DC建設投資", "日本のデータセンター建設（年間）", [3222, 6000, 10500], "1兆円超",
     "約3倍", "5年（'23→'28）", "IDC Japan"),
]
ty, rh = 1.5, 1.42
for i, (what, note, vals, llab, mult, period, src) in enumerate(rows):
    y = ty + i * rh
    if i % 2 == 1:
        rect(s, 0.85, y, 11.63, rh-0.08, fill=PANEL, radius=0.04)
    # left label
    text(s, 1.05, y+0.2, 3.05, 0.95,
         [[R(what, 15, INK, True)],
          [R(note, 9.5, MUTE, True, FONT_M)]], space_after=3, line_spacing=1.12)
    # bars
    vbars(s, 4.25, y+rh-0.28, 3.7, 0.82, vals, llab)
    # right multiple callout
    text(s, 8.35, y+0.16, 2.5, 0.72, [[R(mult, 33, RED, True)]], anchor=MSO_ANCHOR.MIDDLE)
    text(s, 10.85, y+0.28, 1.6, 0.9,
         [[R(period, 9.5, BODY, True)],
          [R(src, 8, MUTE, False, FONT_M)]], space_after=2, line_spacing=1.12, anchor=MSO_ANCHOR.MIDDLE)

rect(s, 0.85, 5.85, 11.63, 0.92, fill=RED, radius=0.05)
text(s, 1.1, 5.85, 11.2, 0.92,
     [[R("しかも投資のピークは2027〜2028年。", 16.5, WHITE, True),
       R("この波は、まだ“序盤”だ。", 16.5, RGBColor(0xFF,0xD2,0xD7), True)]],
     anchor=MSO_ANCHOR.MIDDLE)

# =========================================================
# SLIDE 4 — DC建設費の内訳：ゼネコン vs 製造業、さらに“材” vs “工”
# =========================================================
s = slide()
header(s, "3", "お金の3/4は設備（製造業）。しかも効くのは機器代ではなく“工”。",
       "DC建設費の中身：建設 vs 設備、そして 材（機器）vs 工（作業）")

# --- Bar A：建設費 100% の内訳 ---
bx, by, bw, bh = 0.85, 1.92, 11.63, 1.0
w_gen = bw * 0.25; w_ele = bw * 0.50; w_mec = bw * 0.25
rect(s, bx, by, w_gen, bh, fill=CHAR)
rect(s, bx+w_gen, by, w_ele, bh, fill=RED)
rect(s, bx+w_gen+w_ele, by, w_mec, bh, fill=RED2)
rect(s, bx+w_gen-0.012, by, 0.024, bh, fill=WHITE)
rect(s, bx+w_gen+w_ele-0.012, by, 0.024, bh, fill=WHITE)
text(s, bx, by, w_gen, bh, [[R("建築（躯体）", 10.5, WHITE, True)], [R("約25%", 18, WHITE, True)]],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, space_after=1, line_spacing=1.0)
text(s, bx+w_gen, by, w_ele, bh, [[R("電気設備", 11, WHITE, True)], [R("約50%", 20, WHITE, True)]],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, space_after=1, line_spacing=1.0)
text(s, bx+w_gen+w_ele, by, w_mec, bh, [[R("空調・機械", 10.5, WHITE, True)], [R("約25%", 18, WHITE, True)]],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, space_after=1, line_spacing=1.0)
rect(s, bx, by-0.32, w_gen, 0.25, fill=CHARWA, line=CHAR, line_w=1, radius=0.1)
text(s, bx, by-0.32, w_gen, 0.25, [[R("建設（ゼネコン）", 9.5, CHAR, True, FONT_M)]],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
rect(s, bx+w_gen, by-0.32, w_ele+w_mec, 0.25, fill=REDWA, line=RED, line_w=1, radius=0.1)
text(s, bx+w_gen, by-0.32, w_ele+w_mec, 0.25,
     [[R("設備＝製造業（電気・機械）  約75%", 10, RED, True, FONT_M)]],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# --- つなぎ線＋見出し ---
text(s, 0.85, 3.12, 11.63, 0.3,
     [[R("↓ その「設備 約75%」の中身を、", 11, BODY, True),
       R("材（機器・材料）", 11, MUTE, True), R(" と ", 11, BODY),
       R("工（据付・施工＝人が動く）", 11, RED, True), R(" に分けると", 11, BODY, True)]])

# --- Bar B：設備工事費を 材 vs 工 に分解 ---
b2x, b2y, b2w, b2h = 0.85, 3.5, 11.63, 0.92
w_mat = b2w * 0.60; w_work = b2w * 0.40
rect(s, b2x, b2y, w_mat, b2h, fill=MUTE)
rect(s, b2x+w_mat, b2y, w_work, b2h, fill=RED)
rect(s, b2x+w_mat-0.012, b2y, 0.024, b2h, fill=WHITE)
text(s, b2x, b2y, w_mat, b2h,
     [[R("材：機器・材料（変圧器・冷凍機・UPS等）", 10.5, WHITE, True)], [R("約6割", 17, WHITE, True)]],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, space_after=1, line_spacing=1.0)
text(s, b2x+w_mat, b2y, w_work, b2h,
     [[R("工：据付・施工（人が動く）", 10.5, WHITE, True)], [R("約4割", 17, WHITE, True)]],
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, space_after=1, line_spacing=1.0)
text(s, b2x+w_mat, b2y+b2h+0.04, w_work, 0.24,
     [[R("↑ ここがANDPADの市場（工）", 9.5, RED, True, FONT_M)]], align=PP_ALIGN.CENTER)

# --- パンチライン ---
rect(s, 0.85, 5.02, 11.63, 1.28, fill=REDWA, line=REDBD, line_w=1.2, radius=0.05)
text(s, 1.15, 5.2, 11.0, 0.5,
     [[R("機器そのものの価格（材）は、正直ANDPADには関係ない。", 14.5, INK, True)]])
text(s, 1.15, 5.72, 11.0, 0.5,
     [[R("効くのは“工”＝据付・施工。それでも建設費の約3割。大型DC1棟で数十〜百億円級が、現場作業に向かう。", 14.5, RED, True)]],
     line_spacing=1.15)
text(s, 0.9, 6.42, 11.5, 0.35,
     [[R("出典：日経xTECH（設備工事＝電気50%＋空調等25%）／材工比は積算の一般構成（直接工事費＝材料費＋労務費＋経費）にもとづく目安・物件で変動。", 8.5, MUTE, False, FONT_M)]])

# =========================================================
# SLIDE 5 — 据付・試運転・保守：割合と市場を、同じ物差しで
# =========================================================
s = slide()
header(s, "4", "「作って終わり」じゃない。建てる時も、建てた後も、現場が続く。",
       "“工”の3工程：据付・試運転・保守——割合と市場を同じ物差しで")

# table
tx = [0.85, 3.05, 6.35, 8.7]        # 工程 / やること / 建設費に対する割合 / 市場の広がり
twd = [2.2, 3.3, 2.35, 3.78]
thead = ["工程", "やること", "建設費に対する割合", "市場の広がり（世界）"]
trows = [
    ("① 据付・施工", "機器を搬入し、据付・配線・配管する", "約3割", "設備の“工”＝ 建設費の約30%",
     "国内の設備工事そのもの。ANDPADの中核市場", "建てる時"),
    ("② 試運転", "通電・調整・性能検証で“動かす”", "約1〜3%", "＋ 建設費の1〜3%",
     "試運転サービス市場 $2.15B→$4.25B（'24→'33）", "建てる時"),
    ("③ 保守・運用", "点検・更新・障害対応（稼働中ずっと）", "毎年発生", "運用費の約4割（毎年）",
     "DC運用保守市場 $15.8B→$45.6B（'24→'33）", "建てた後ずっと"),
]
ty0 = 1.55; hh = 0.46; drh = 1.02
# header
rect(s, tx[0], ty0, sum(twd), hh, fill=RED)
for j, htxt in enumerate(thead):
    text(s, tx[j]+0.14, ty0, twd[j]-0.2, hh, [[R(htxt, 10.5, WHITE, True, FONT_M)]],
         anchor=MSO_ANCHOR.MIDDLE)
for i, (proc, doing, pct, scale, market, when) in enumerate(trows):
    y = ty0 + hh + i * drh
    if i % 2 == 1:
        rect(s, tx[0], y, sum(twd), drh, fill=PANEL)
    # 工程 + when tag
    wtag_col = RED if when.startswith("建てる") else CHAR
    text(s, tx[0]+0.14, y+0.14, twd[0]-0.22, drh-0.2,
         [[R(proc, 13, INK, True)],
          [R(when, 8.5, wtag_col, True, FONT_M)]], space_after=3, line_spacing=1.1)
    # やること
    text(s, tx[1]+0.14, y, twd[1]-0.24, drh, [[R(doing, 11, BODY)]],
         anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.16)
    # 割合
    text(s, tx[2]+0.14, y+0.16, twd[2]-0.24, 0.5, [[R(pct, 20, RED, True)]])
    text(s, tx[2]+0.14, y+0.62, twd[2]-0.24, 0.32, [[R(scale, 8.5, MUTE, True, FONT_M)]], line_spacing=1.05)
    # 市場
    text(s, tx[3]+0.14, y, twd[3]-0.26, drh, [[R(market, 10.5, INK, True)]],
         anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.16)
# grid lines
for i in range(len(trows)+1):
    rect(s, tx[0], ty0+hh+i*drh, sum(twd), 0.01, fill=LINE)
for j in range(1, len(tx)):
    rect(s, tx[j], ty0+hh, 0.01, drh*len(trows), fill=LINE)

# 建てる時 / 建てた後 の帯
rect(s, 0.85, 4.75, 11.63, 0.5, fill=CHARWA, line=CHAR, line_w=1, radius=0.06)
text(s, 1.1, 4.75, 11.2, 0.5,
     [[R("①②「建てる時」は一度きり。", 11.5, CHAR, True),
       R(" だが ③「保守」は稼働中ずっと、毎年続く。", 11.5, RED, True),
       R(" → 一度つかんだ現場が、長く残る。", 11.5, INK, True)]],
     anchor=MSO_ANCHOR.MIDDLE)

# パンチライン
rect(s, 0.85, 5.42, 11.63, 0.88, fill=REDWA, line=REDBD, line_w=1.2, radius=0.05)
text(s, 1.15, 5.42, 11.0, 0.88,
     [[R("これまで建設業だけを相手にしてきた。", 15, INK, True),
       R("その周辺に、据付＋試運転＋保守という“人が動く”広大な市場がある。", 15, RED, True)]],
     anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.15)
text(s, 0.9, 6.42, 11.5, 0.35,
     [[R("出典：試運転＝建設費1〜3%（業界目安）・DC試運転市場 DataIntelo／DC運用保守市場 DataHorizon／保守は運用費の約4割 Thunder Said Energy。", 8.5, MUTE, False, FONT_M)]])

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
