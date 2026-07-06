"""pptxと同じ座標系(inch×96px)でHTMLプレビューを生成する（見た目確認用）。
出力: output/preview.html
"""
from pathlib import Path

RED="#E60012"; RED2="#F06670"; INK="#15181C"; SUB="#373C43"; MUTE="#6E767F"
FAINT="#9AA1A9"; HAIR="#D6DADF"; G1="#D2D6DB"; GDK="#535961"; WHITE="#fff"
PX=96
def i2p(v): return v*PX
def pt(v): return v*96/72

slides=[]
buf=[]
def S():
    buf.clear()
def flush():
    slides.append("".join(buf))

def rect(x,y,w,h,fill):
    buf.append(f'<div style="position:absolute;left:{i2p(x)}px;top:{i2p(y)}px;width:{i2p(w)}px;height:{i2p(h)}px;background:{fill}"></div>')
def hr(x,y,w,c=HAIR,wt=1.0):
    rect(x,y,w,wt/72.0,c)
def vr(x,y,h,c=HAIR,wt=1.0):
    rect(x,y,wt/72.0,h,c)
def T(x,y,w,h,spans,align="left",valign="top",ls=1.1):
    # spans: inline runs (text,size,color,bold,mono,tracking)
    inner=""
    for (tx,sz,col,bold,mono,tr) in spans:
        fam="Consolas,ui-monospace,monospace" if mono else "'Noto Sans JP','Hiragino Sans','Yu Gothic',sans-serif"
        inner+=f'<span style="font-size:{pt(sz):.1f}px;color:{col};font-weight:{700 if bold else 400};font-family:{fam};letter-spacing:{tr*0.7:.1f}px">{tx}</span>'
    ai={"top":"flex-start","middle":"center","bottom":"flex-end"}[valign]
    buf.append(f'<div style="position:absolute;left:{i2p(x)}px;top:{i2p(y)}px;width:{i2p(w)}px;height:{i2p(h)}px;'
               f'display:flex;align-items:{ai}">'
               f'<div style="width:100%;line-height:{ls};text-align:{align}">{inner}</div></div>')
def R(t,s,c=INK,b=False,m=False,tr=0): return (t,s,c,b,m,tr)

LM,RM=0.9,12.43; CW=RM-LM
def takeaway(y,spans,h=0.6):
    rect(LM,y+0.05,0.09,h-0.1,RED); T(LM+0.28,y,CW-0.3,h,spans,valign="middle",ls=1.16)
def header(page,eye,title):
    rect(LM,0.62,0.14,0.14,RED)
    T(1.14,0.55,9.5,0.28,[R(eye,11.5,RED,True,True,1.5)])
    T(RM-0.7,0.55,0.7,0.28,[R(page,11.5,FAINT,True,True)],align="right")
    T(LM-0.02,0.92,CW,0.55,[R(title,24,INK,True)])
    hr(LM,1.6,CW,wt=1.4)

# S1
S()
rect(LM,1.6,0.15,0.15,RED)
T(1.16,1.52,10.5,0.3,[R("ANDPAD 本部総会   ·   WHY MANUFACTURING NOW",13,RED,True,True,2)])
T(LM-0.04,2.45,11.8,1.1,[R("建設SaaSの我々が、",46,INK,True)],ls=1.12)
T(LM-0.04,3.35,11.8,1.1,[R("なぜ、いま",46,INK,True),R("製造業",46,RED,True),R("なのか。",46,INK,True)],ls=1.12)
rect(LM,4.85,3.2,0.05,RED)
T(LM-0.02,5.15,11.7,0.5,[R("AIをはじめ世界のメガトレンドは、膨大な“設備”を必要とする。",17,SUB)])
T(LM-0.02,5.72,11.7,0.6,[R("それを作り・建てるのは、建設業だけでなく製造業がコア——その市場が、いま圧倒的に広がっている。",17,SUB)])
hr(LM,6.95,CW)
T(LM-0.02,7.06,11.7,0.3,[R("2026-07   /   AIインフラ・バリューチェーン調査（国内外154社・決算/IR一次情報）＋公開データ",10.5,FAINT,False,True)])
flush()

# S2
S(); header("01","世界のメガトレンドと、その“共通の土台”","世界が動く先には、必ず「設備」がいる。その主役は、製造業だ。")
cols=[("01","AI",True,"生成AI・データセンター・半導体","電源・冷却・建屋・チップ工場が大量に要る"),
("02","脱炭素（GX）",False,"再エネ・送電網・電化・蓄電池","発電・変電・ケーブル・電池／EV工場"),
("03","経済安保・国内回帰",False,"半導体・重要物資の国産化","国内に工場を新設し、供給網を再構築")]
pitch=CW/3; ctop=1.98
for k,(no,nm,hot,what,need) in enumerate(cols):
    x=LM+k*pitch
    if k>0: vr(x-0.24,ctop+0.05,2.25,wt=1.2)
    T(x,ctop,pitch-0.5,0.26,[R(no,11,FAINT,True,True,1)])
    T(x,ctop+0.32,pitch-0.42,0.5,[R(nm,23,RED if hot else INK,True)])
    if hot: T(x,ctop+0.9,pitch-0.4,0.26,[R("●  本日フォーカス",11,RED,True,True)])
    T(x,ctop+1.34,pitch-0.5,0.44,[R(what,14.5,MUTE,True)],ls=1.18)
    T(x,ctop+1.82,pitch-0.5,0.55,[R("→ ",14.5,RED,True),R(need,14.5,SUB)],ls=1.22)
hr(LM,4.6,CW)
T(LM-0.02,4.85,11.7,0.5,[R("どのメガトレンドも、実現するには",19,INK,True),R("膨大な“設備”",19,RED,True),R("がいる。",19,INK,True)])
T(LM-0.02,5.5,11.7,0.5,[R("設備を作る・建てるのは、建設業だけではない——",15.5,SUB),R("重電・機械・電機など製造業がコア。",15.5,INK,True)])
takeaway(6.25,[R("→ 今日はその中で、最も勢いのある「",19,INK,True),R("AI",19,RED,True),R("」に絞って話す。",19,INK,True)],h=0.62)
flush()

# S3
S(); header("02","規模より“勢い”——トレンドとして、どれだけ急か","AI投資は“大きい”だけじゃない。数年で“何倍”に跳ねている。")
def vbars(x0,base_y,area_w,max_h,values,years,vlab):
    n=len(values); mx=max(values); slot=area_w/n; bw=min(0.5,slot*0.56)
    hr(x0,base_y,area_w,c=HAIR,wt=1.4)
    for k,v in enumerate(values):
        h=max_h*v/mx; bxx=x0+k*slot+(slot-bw)/2
        rect(bxx,base_y-h,bw,h,RED if k==n-1 else G1)
        T(x0+k*slot,base_y+0.07,slot,0.24,[R(years[k],9.5,MUTE,True,True)],align="center")
        if k==n-1: T(bxx-0.6,base_y-h-0.26,bw+1.2,0.24,[R(vlab,10,MUTE,True,True)],align="center")
rows=[("AIチップの需要","NVIDIA データセンター売上",[15,48,115,196],["FY23","FY24","FY25","FY26"],"約2,000億ドル","約13倍"),
("AIインフラ投資","Big Tech 4社の設備投資（年間）",[180,230,410,725],["’23","’24","’25","’26"],"7,250億ドル","約4倍"),
("国内DC建設投資","日本のデータセンター建設（年間）",[3222,5000,10000],["’23","’24","’28"],"1兆円超","約3倍")]
ty,rh=1.78,1.3
for k,(what,note,vals,yrs,vlab,mult) in enumerate(rows):
    y=ty+k*rh
    if k>0: hr(LM,y-0.03,CW)
    T(LM,y+0.28,3.1,0.5,[R(what,17,INK,True)])
    T(LM,y+0.62,3.1,0.3,[R(note,10.5,MUTE,True,True)])
    vbars(4.15,y+rh-0.42,3.9,0.66,vals,yrs,vlab)
    T(8.65,y+0.2,1.75,0.85,[R(mult,38,RED,True)],valign="middle")
    T(10.45,y+0.2,2.0,0.85,[R("に伸びた",13,MUTE,True)],valign="middle")
takeaway(5.95,[R("投資のピークは2027〜2028年。この波は、まだ“",19,INK,True),R("序盤",19,RED,True),R("”だ。",19,INK,True)],h=0.62)
T(LM,6.82,11.6,0.28,[R("出典：NVIDIA IR／Big Tech各社IR・報道／IDC Japan（国内DC建設投資）。",9,FAINT,False,True)])
flush()

# S4
S(); header("03","AI需要 → 連鎖して伸びる、日本のバリューチェーン","AIが伸びれば、この“川”がまるごと潤う。")
T(LM,1.9,8.0,0.34,[R("AI・DC需要の拡大",15,RED,True),R("　この一手が、川下までまるごと波及する",11,MUTE,True)])
rect(LM,2.4,CW,0.045,RED)
stages=[("01","発電・電源",["三菱重工","川崎重工","IHI","デンヨー"]),
("02","送変電・電線",["日立","三菱電機","ダイヘン","フジクラ"]),
("03","DC建設・設備",["鹿島建設","大林組","きんでん","高砂熱学"]),
("04","冷却・電源保護",["ダイキン","GSユアサ","荏原製作所"]),
("05","半導体",["東京エレクトロン","ディスコ","キオクシア","信越化学"])]
spitch=CW/5; stop=2.75
for k,(no,stg,cos) in enumerate(stages):
    x=LM+k*spitch
    if k>0: vr(x-0.12,stop,2.55,wt=1.2)
    T(x,stop,spitch-0.3,0.22,[R(no,10,FAINT,True,True,1)])
    T(x,stop+0.28,spitch-0.26,0.5,[R(stg,14,INK,True)],ls=1.06)
    for j,co in enumerate(cos):
        T(x,stop+0.9+j*0.42,spitch-0.26,0.34,[R(co,12,SUB,True)])
takeaway(5.7,[R("AIの“源流”が、川下の日本メーカーまで、",19,INK,True),R("まるごと潤す。",19,RED,True)],h=0.58)
T(LM+0.28,6.38,11.3,0.34,[R("狙える現場は、この一社一社にある。",14,SUB,True)])
T(LM,6.92,11.6,0.28,[R("※ 各段階の代表企業を抜粋（社名表記＝ロゴのイメージ、実ロゴへ差し替え可）。数値・詳細は次頁以降。",9,FAINT,False,True)])
flush()

# S5
S(); header("04","建設費の中身：どこにANDPADの市場があるか","お金の3/4は設備。効くのは機器代ではなく“工”＝据付・試運転・保守。")
bx,by,bw,bh=LM,2.2,CW,0.98
wg=bw*0.25; we=bw*0.5; wm=bw*0.25
rect(bx,by,wg,bh,GDK); rect(bx+wg,by,we,bh,RED); rect(bx+wg+we,by,wm,bh,RED2)
rect(bx+wg-0.012,by,0.024,bh,WHITE); rect(bx+wg+we-0.012,by,0.024,bh,WHITE)
def barlabel(x,w,l1,s1,l2,s2,col):
    buf.append(f'<div style="position:absolute;left:{i2p(x)}px;top:{i2p(by)}px;width:{i2p(w)}px;height:{i2p(bh)}px;'
               f'display:flex;flex-direction:column;justify-content:center;align-items:center;line-height:1.05">'
               f'<span style="font-size:{pt(s1):.1f}px;color:{col};font-weight:700;font-family:\'Noto Sans JP\',sans-serif">{l1}</span>'
               f'<span style="font-size:{pt(s2):.1f}px;color:{col};font-weight:700;font-family:\'Noto Sans JP\',sans-serif">{l2}</span></div>')
barlabel(bx,wg,"建築（躯体）",11,"約25%",19,WHITE)
barlabel(bx+wg,we,"電気設備",12,"約50%",22,WHITE)
barlabel(bx+wg+we,wm,"空調・機械",11,"約25%",19,WHITE)
T(bx,by-0.32,wg,0.26,[R("建設（ゼネコン）",10.5,MUTE,True,True)])
T(bx+wg,by-0.32,we+wm,0.26,[R("設備＝製造業（電気・機械）  約75%",11,RED,True,True)])
T(LM,3.45,CW,0.28,[R("その「設備 約75%」を  ",12.5,SUB,True),R("材（機器・材料）",12.5,MUTE,True),R("  と  ",12.5,SUB),R("工（据付・施工＝人が動く）",12.5,RED,True),R("  に分けると",12.5,SUB,True)])
b2y,b2h=3.82,0.88; wmat=CW*0.6; wwork=CW*0.4
rect(LM,b2y,wmat,b2h,G1); rect(LM+wmat,b2y,wwork,b2h,RED); rect(LM+wmat-0.012,b2y,0.024,b2h,WHITE)
T(LM,b2y,wmat,b2h,[R("材：機器・材料（変圧器・冷凍機・UPS 等）  約6割",12.5,INK,True)],align="center",valign="middle")
T(LM+wmat,b2y,wwork,b2h,[R("工：据付・施工  約4割",12.5,WHITE,True)],align="center",valign="middle")
T(LM+wmat,b2y+b2h+0.06,wwork,0.24,[R("↑ ここがANDPADの市場（工）",10,RED,True,True)],align="center")
wy=5.28; work=[("据付・施工","建設費の約3割","国内の設備工事そのもの＝中核市場"),("試運転","建設費の1〜3%","世界の試運転市場 $2.1B → $4.3B"),("保守・運用","毎年・運用費の約4割","世界のDC運用保守 $15.8B → $45.6B")]
wpitch=CW/3
for k,(nm,pct,mk) in enumerate(work):
    x=LM+k*wpitch
    if k>0: vr(x-0.2,wy,0.7,wt=1.2)
    T(x,wy,wpitch-0.4,0.3,[R(nm,13.5,INK,True),R("  "+pct,12,RED,True)])
    T(x,wy+0.36,wpitch-0.4,0.32,[R(mk,10.5,MUTE,True)],ls=1.14)
takeaway(6.32,[R("建設業だけを見てきた。その周辺の“工”（据付・試運転・保守）に、",16,INK,True),R("広大な市場",16,RED,True),R("がある。",16,INK,True)],h=0.56)
flush()

# S6
S(); header("05","バリューチェーン別・国内主要プレーヤーと伸び","国内のAI投資は、製造業のこの現場に着地している。")
flow=[("01","発電・電源","電気をつくる","受注残 5兆円超","三菱重工"),
("02","送変電・電線","電気を送る・変える","変圧器 生産能力2倍","日立・ダイヘン・フジクラ"),
("03","DC建設・設備工事","箱を建てる","空調工事 受注 +25.5%","ダイダン・高砂熱学"),
("04","冷却・電源保護","冷やす・止めない","北米DC冷却 約13倍","ダイキン・GSユアサ"),
("05","半導体","計算する頭脳","設備投資 +66%","キオクシア・東京ｴﾚｸﾄﾛﾝ")]
fpitch=CW/5; ftop=2.0
for k,(no,stg,sb,stat,cos) in enumerate(flow):
    x=LM+k*fpitch
    if k>0: vr(x-0.12,ftop,3.0,wt=1.2)
    T(x,ftop,fpitch-0.28,0.22,[R(no,10,FAINT,True,True,1)])
    T(x,ftop+0.28,fpitch-0.24,0.5,[R(stg,14,INK,True)],ls=1.05)
    T(x,ftop+0.82,fpitch-0.24,0.26,[R(sb,10.5,MUTE,True)])
    T(x,ftop+1.28,fpitch-0.3,0.78,[R(stat,16,RED,True)],ls=1.1)
    T(x,ftop+2.35,fpitch-0.26,0.6,[R(cos,10.5,SUB,True)],ls=1.18)
hr(LM,5.2,CW)
T(LM,5.42,11.6,0.5,[R("我々の入り方：",12,MUTE,True,True),R("  A 発注者＝“建てる側”で管理",13,SUB,True),R("   /   ",12,FAINT),R("B 請負＝据付・試運転・保守を“納める側”で管理（本命）",13,RED,True)])
takeaway(6.15,[R("川上から川下まで、全段が同時に増収増益。",17,INK,True),R("狙える現場が、日本中に生まれている。",17,RED,True)],h=0.6)
flush()

# S7
S()
rect(LM,1.2,0.15,0.15,RED)
T(1.16,1.12,10.5,0.3,[R("SO, LET'S GO",13,RED,True,True,2)])
T(LM-0.04,1.78,11.8,1.2,[R("次の主戦場は、",44,INK,True),R("製造業",44,RED,True),R("だ。",44,INK,True)])
rect(LM,3.35,3.2,0.05,RED)
pts=[("市場は建設の3倍広い","DC建設費の約3/4は電気・機械設備＝製造業。据付・試運転・保守まで現場が続く。"),
("追い風は本物","AIチップ需要は3年で約13倍、投資ピークは2027-28。受注残に裏打ちされた構造需要。"),
("武器は完成済み","建設で磨いた現場管理は、製造業の据付・保守フィールドに“そのまま効く”。")]
py=3.85
for k,(h,b) in enumerate(pts):
    y=py+k*0.92
    if k>0: hr(LM,y-0.08,CW)
    T(LM,y,0.6,0.65,[R(f"0{k+1}",17,RED,True,True)],valign="middle")
    T(LM+0.75,y+0.06,4.5,0.6,[R(h,17,INK,True)],valign="middle")
    T(LM+5.4,y+0.06,6.2,0.62,[R(b,11.5,SUB)],valign="middle",ls=1.16)
rect(LM,6.72,CW,0.05,RED)
T(LM-0.02,6.9,11.7,0.5,[R("建設の“周辺”に広がる巨大市場を、100人でつかみにいこう。",21,INK,True)])
flush()

cards="".join(f'<div class="slide">{c}</div>' for c in slides)
html=f"""<!doctype html><meta charset="utf-8"><style>
*{{margin:0;box-sizing:border-box}}
body{{background:#4a4d52;padding:20px;font-family:'Noto Sans JP','Hiragino Sans','Yu Gothic',sans-serif}}
.slide{{position:relative;width:1280px;height:720px;background:#fff;margin:0 auto 22px;overflow:hidden;box-shadow:0 2px 14px rgba(0,0,0,.4)}}
</style>{cards}"""
out=Path(__file__).resolve().parent.parent/"output"/"preview.html"
out.write_text(html,encoding="utf-8")
print("saved",out,"slides",len(slides))
