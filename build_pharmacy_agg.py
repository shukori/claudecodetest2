# -*- coding: utf-8 -*-
"""
在宅訪問調剤薬局 事業者集計ビルダー（A案）
 出力1: 福岡_事業者集計.xlsx   （福岡=基本薬局一覧 → 全保険薬局の事業者別店舗数。在宅フラグ無し）
 出力2: 九州在宅_事業者集計.xlsx（佐賀・長崎・熊本・沖縄=施設基準届出 → 在宅届出薬局の事業者別店舗数）
 設計: 集計はCOUNTIFSでローデータ参照／名寄せ辞書・名寄せ履歴シートで監査可能
"""
import openpyxl, re
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from collections import defaultdict, Counter, OrderedDict

SRC = "/root/.claude/uploads/5f0eed8e-ea94-5c75-b8ad-040dd274105a"
OUT = "/home/user/ClaudeCodetest2"

# ================= 名寄せ辞書 =================
# (ルールID, 正規化事業者名, タイプ, 親会社/メモ, 信頼度, [マッチ語(原文/屋号いずれかに部分一致)])
# 上から順に評価し、最初に一致したルールを採用（first match wins）。
RULES = [
 ("R01","アインホールディングス（アイン薬局）","大手全国","東証PRM・国内最大手","高",["アインファーマシーズ","アイン薬局","アインメディオ","アインHD"]),
 ("R02","総合メディカル（そうごう薬局）","大手全国","SOGO MEDICAL・福岡発祥の全国大手","高",["総合メディカル","そうごう薬局","ソウゴウ薬局"]),
 ("R03","日本調剤","大手全国","東証PRM","高",["日本調剤"]),
 ("R04","クオール（クオール薬局）","大手全国","東証PRM","高",["クオール"]),
 ("R05","メディカルシステムネットワーク（なの花薬局）","大手全国","札幌発祥・なの花","高",["なの花","メディカルシステムネットワーク"]),
 ("R06","クラフト（さくら薬局）","大手全国","さくら薬局グループ","中",["クラフト","さくら薬局"]),  # ※「さくら調剤薬局」は別:要確認
 ("R07","HYUGA PRIMARY CARE（きらり薬局）","在宅特化・上場","東証GRT・分析対象本命","高",["ＨＹＵＧＡ","HYUGA","きらり薬局"]),
 ("R08","ウエルシアHD（イオン系）","ドラッグ大手","調剤併設","高",["ウエルシア","イオンウエルシア"]),
 ("R09","ツルハHD","ドラッグ大手","調剤併設","高",["ツルハ","くすりの福太郎","ウェルネス"]),
 ("R10","マツキヨココカラ&Co","ドラッグ大手","調剤併設","高",["ココカラファイン","マツモトキヨシ","マツキヨ"]),
 ("R11","スギHD（スギ薬局）","ドラッグ大手","調剤併設","高",["スギ薬局","スギホールディングス"]),
 # --- 九州・福岡 地場大手/中堅 ---
 ("R20","大賀薬局","地場大手(福岡)","福岡地盤の有力チェーン","高",["大賀薬局"]),
 ("R21","新生堂薬局","地場大手(福岡)","福岡・ドラッグ＋調剤","高",["新生堂"]),
 ("R22","サンキュードラッグ","地場(北九州)","北九州地盤ドラッグ","高",["サンキュードラッグ"]),
 ("R23","タカラ薬局（大蔵商事/タカラ）","地場(福岡)","福岡地盤","中",["タカラ薬局"]),
 ("R24","溝上薬局","地場大手(福岡)","福岡・九州で展開","高",["溝上薬局"]),
 ("R25","タケシタ調剤薬局","地場(福岡・熊本)","","高",["タケシタ"]),
 ("R26","コスモス薬品","ドラッグ(九州発)","ディスカウントドラッグコスモス","高",["コスモス薬品","ディスカウントドラッグコスモス"]),
 ("R27","ドラッグストアモリ","ドラッグ(九州)","","高",["ドラッグストアモリ","ｄｒｕｇ　ｓｔｏｒｅ　ｍｏｒｉ"]),
 ("R28","ドラッグイレブン","ドラッグ(九州)","JR九州系","中",["ドラッグイレブン"]),
 ("R29","大信薬局","地場(福岡)","在宅・施設に注力","中",["大信薬局"]),
 ("R30","ニック（ニック調剤薬局）","地場(九州)","","中",["ニック"]),
 ("R31","アガペ","地場(福岡)","","中",["アガペ"]),
 ("R32","ミズ","地場(福岡)","","中",["株式会社ミズ","　ミズ　","ミズ薬局"]),
 ("R33","タカサキ（高崎）薬局","地場","","中",["タカサキ","高崎薬局"]),
 # --- 名前が一般的で複数事業者混在の恐れ → 中/要確認で個別ルール化 ---
 ("R40","すこやか薬局［要確認：同名複数の可能性］","要確認","汎用名・複数法人混在の恐れ","要確認",["すこやか薬局"]),
 ("R41","さくら調剤薬局［要確認：クラフトと別］","要確認","R06と別法人の可能性","要確認",["さくら調剤"]),
]

def clean_corp(kaisetsu):
    """福岡: 開設者列から法人名コア（代表者名を除去）"""
    if not kaisetsu: return ""
    s=str(kaisetsu)
    s=re.split(r'(代表取締役|代表者|取締役|理事長|理事|院長|管理者|会長|社長|代表社員|業務執行社員|開設者)',s)[0]
    return s.replace('　',' ').strip()

def base_yago(name):
    """九州: 屋号から事業者コア（…薬局/…ファーマシー まで）"""
    if not name: return ""
    s=str(name).replace('　','').replace(' ','')
    for t in ['株式会社','有限会社','合同会社','合資会社','医療法人社団','医療法人','一般社団法人','社会福祉法人','公益財団法人']:
        s=s.replace(t,'')
    m=re.search(r'^(.*?(薬局|ファーマシー|ドラッグ|薬店|調剤|ファーマ))',s)
    core=m.group(1) if m else s[:10]
    if core in ('薬局','調剤','保険薬局') or len(core)<2:
        core=s[:10]  # 正規化失敗→原文先頭
    return core

def match_rule(raw_for_match):
    """原文(開設者 or 屋号)に対してルールを評価。戻り: (ルールID,事業者,タイプ,信頼度) or None"""
    s=str(raw_for_match).replace('　','').replace(' ','')
    for rid,name,typ,memo,conf,aliases in RULES:
        for a in aliases:
            if a.replace('　','').replace(' ','') in s:
                return rid,name,typ,conf
    return None

# ================= スタイル =================
HEAD = Font(bold=True,color="FFFFFF",size=11,name="IPAGothic")
HFILL= PatternFill("solid",fgColor="2E6DA4")
GFILL= PatternFill("solid",fgColor="E3F2E9")
YFILL= PatternFill("solid",fgColor="FFF2CC")
TITLE= Font(bold=True,size=14,name="IPAGothic")
NORM = Font(size=10,name="IPAGothic")
BOLD = Font(bold=True,size=10,name="IPAGothic")
thin = Side(style="thin",color="BBBBBB")
BORD = Border(left=thin,right=thin,top=thin,bottom=thin)
WRAP = Alignment(wrap_text=True,vertical="top")
CEN  = Alignment(horizontal="center",vertical="center")

def style_header(ws,row,ncol):
    for c in range(1,ncol+1):
        cell=ws.cell(row=row,column=c); cell.font=HEAD; cell.fill=HFILL
        cell.alignment=CEN; cell.border=BORD

def link_font():
    return Font(size=10,name="IPAGothic",color="0563C1",underline="single")

# ================= ファイル定義 =================
FUKUOKA_FILE = ("福岡県","c253f318-r8_06_fukuoka_yakkyoku_02.xlsx")
SHISETSU = OrderedDict([
 ("佐賀県","3b941e75-r8_05_shisetsu_saga_yakkyoku_02.xlsx"),
 ("長崎県","e0bd4eb7-r8_05_shisetsu_nagasaki_yakkyoku_02.xlsx"),
 ("熊本県","a0c45bc5-r8_05_shisetsu_kumamoto_yakkyoku_02.xlsx"),
 ("大分県","19405bbf-r8_05_shisetsu_ooita_yakkyoku_02.xlsx"),
 ("宮崎県","da1c0ae4-r8_05_shisetsu_miyazaki_yakkyoku_02.xlsx"),
 ("鹿児島県","4ad91e46-r8_05_shisetsu_kagoshima_yakkyoku_02.xlsx"),
 ("沖縄県","0655d75f-r8_05_shisetsu_okinawa_yakkyoku_02.xlsx"),
])
SRC_URL = "https://kouseikyoku.mhlw.go.jp/kyushu/gyomu/gyomu/hoken_kikan/index_00004.html"
TAISEI={"在宅薬学総合体制加算１","在宅薬学総合体制加算２"}
HOMEV ="在宅患者訪問薬剤管理指導料"

# ================= パース =================
def parse_fukuoka():
    """基本薬局一覧: 1薬局1行。code,name,addr,開設者raw を返す"""
    import os
    wb=openpyxl.load_workbook(os.path.join(SRC,FUKUOKA_FILE[1]),read_only=True)
    ws=wb["Sheet1"]; rows=[]
    for row in ws.iter_rows(values_only=True):
        c1=row[0]
        if c1 is None or not str(c1).strip().isdigit(): continue
        c2=row[1]
        if c2 is None or not re.match(r'^\d{2,3},\d{3},\d$',str(c2).strip()): continue
        code=str(c2).strip(); name=row[2]; addr=row[3]; kaisetsu=row[5]
        rows.append(dict(pref="福岡県",code=code,name=name,addr=addr,raw=kaisetsu))
    wb.close(); return rows

def parse_shisetsu():
    """施設基準届出: 1薬局複数行→集約。code,name,addr,pref,flags"""
    import os
    ph=OrderedDict()
    for pref,f in SHISETSU.items():
        wb=openpyxl.load_workbook(os.path.join(SRC,f),read_only=True); ws=wb["Sheet1"]
        for row in ws.iter_rows(min_row=5,values_only=True):
            if row[0] is None or not str(row[0]).strip().isdigit(): continue
            code=row[4]
            if code is None: continue
            key=(pref,str(code))
            if key not in ph:
                ph[key]=dict(pref=pref,code=str(code),name=row[7],addr=row[9],flags=set())
            if row[13]: ph[key]["flags"].add(row[13])
        wb.close()
    return list(ph.values())

# ================= 名寄せ適用 =================
def normalize(rows, match_field):
    """match_field: 'raw'(福岡開設者) or 'name'(九州屋号). 各rowに正規化結果を付与"""
    hist=OrderedDict()  # 原キー -> dict
    for r in rows:
        if match_field=="raw":
            rawkey=clean_corp(r["raw"]); fallback=rawkey
        else:
            rawkey=str(r["name"] or ""); fallback=base_yago(r["name"])
        m=match_rule(rawkey if match_field=="raw" else r["name"])
        if m:
            rid,corp,typ,conf=m
        else:
            rid,corp,typ,conf="-",fallback,"中小・独立","自動"
        r["corp"]=corp; r["rid"]=rid; r["type"]=typ; r["conf"]=conf
        hk=(rawkey, rid, corp)
        if hk not in hist:
            hist[hk]=dict(raw=rawkey,rid=rid,corp=corp,type=typ,conf=conf,cnt=0)
        hist[hk]["cnt"]+=1
    return list(hist.values())

print("logic loaded")

# ================= 共通シート =================
def add_dict_sheet(wb):
    ws=wb.create_sheet("01_名寄せ辞書")
    ws["A1"]="名寄せ辞書（大手・主要チェーン）"; ws["A1"].font=TITLE
    ws["A2"]="上から順に評価し最初に一致したルールを採用（first match wins）。未一致は『中小・独立』として原名称コアを事業者キーに採用。"
    ws["A2"].font=NORM
    heads=["ルールID","正規化事業者名","タイプ","親会社・メモ","信頼度","マッチ語（部分一致）"]
    for j,h in enumerate(heads,1): ws.cell(row=4,column=j,value=h)
    style_header(ws,4,len(heads))
    r=5
    for rid,name,typ,memo,conf,aliases in RULES:
        ws.cell(row=r,column=1,value=rid)
        ws.cell(row=r,column=2,value=name)
        ws.cell(row=r,column=3,value=typ)
        ws.cell(row=r,column=4,value=memo)
        ws.cell(row=r,column=5,value=conf)
        ws.cell(row=r,column=6,value="；".join(aliases))
        for c in range(1,7):
            ws.cell(row=r,column=c).font=NORM; ws.cell(row=r,column=c).border=BORD
            ws.cell(row=r,column=c).alignment=WRAP
        if conf=="要確認":
            for c in range(1,7): ws.cell(row=r,column=c).fill=YFILL
        r+=1
    widths=[10,34,16,30,10,46]
    for j,w in enumerate(widths,1): ws.column_dimensions[get_column_letter(j)].width=w
    return ws

def add_history_sheet(wb, hist, raw_label):
    ws=wb.create_sheet("02_名寄せ履歴")
    ws["A1"]="名寄せ履歴（監査用：原名称→適用ルール→正規化事業者）"; ws["A1"].font=TITLE
    ws["A2"]=f"原データの『{raw_label}』を一意化し、どのルールで・どの事業者へ寄せたかを全件表示。件数の合計はローデータ行数と一致。"
    ws["A2"].font=NORM
    heads=[f"原・{raw_label}（クリーニング後）","適用ルールID","正規化事業者名","タイプ","信頼度","該当店舗数"]
    for j,h in enumerate(heads,1): ws.cell(row=4,column=j,value=h)
    style_header(ws,4,len(heads))
    # sort: 大手(ルール有)を上、件数降順
    hist_sorted=sorted(hist,key=lambda x:(x["rid"]=="-", -x["cnt"]))
    r=5
    for h in hist_sorted:
        ws.cell(row=r,column=1,value=h["raw"])
        ws.cell(row=r,column=2,value=h["rid"])
        ws.cell(row=r,column=3,value=h["corp"])
        ws.cell(row=r,column=4,value=h["type"])
        ws.cell(row=r,column=5,value=h["conf"])
        ws.cell(row=r,column=6,value=h["cnt"])
        for c in range(1,7):
            ws.cell(row=r,column=c).font=NORM; ws.cell(row=r,column=c).border=BORD
        if h["conf"]=="要確認":
            for c in range(1,7): ws.cell(row=r,column=c).fill=YFILL
        elif h["rid"]!="-":
            for c in range(1,7): ws.cell(row=r,column=c).fill=GFILL
        r+=1
    ws.cell(row=r,column=5,value="合計").font=BOLD
    ws.cell(row=r,column=6,value=f"=SUM(F5:F{r-1})").font=BOLD
    widths=[34,12,34,16,10,12]
    for j,w in enumerate(widths,1): ws.column_dimensions[get_column_letter(j)].width=w
    ws.freeze_panes="A5"
    return ws

# ================= 福岡 ワークブック =================
def build_fukuoka():
    rows=parse_fukuoka()
    hist=normalize(rows,"raw")
    wb=openpyxl.Workbook(); wb.remove(wb.active)
    # README
    ws=wb.create_sheet("00_README")
    L=[("福岡県 保険薬局 事業者集計（A案）",TITLE),
       ("",NORM),
       ("【元データ】令和8年6月1日現在 保険薬局一覧（コード内容別医療機関一覧表）／九州厚生局",BOLD),
       (f"  出典URL: {SRC_URL}",link_font()),
       ("  ファイル: r8_06_fukuoka_yakkyoku_02.xlsx",NORM),
       ("",NORM),
       ("【重要・データ種別の注意】",BOLD),
       ("  本ファイルは『基本薬局一覧』で、開設者（法人名）は全件あるが、在宅薬学総合体制加算などの",NORM),
       ("  施設基準（=在宅対応フラグ）が含まれない。よって本集計は『全保険薬局の事業者別店舗数』であり、",NORM),
       ("  在宅対応の絞り込みは不可。在宅で絞るには福岡県の“施設基準届出状況”ファイル（shisetsu版）が必要。",NORM),
       ("  → そのファイルを頂ければ、九州他県と同じ在宅ベースに再集計します。",NORM),
       ("",NORM),
       ("【在宅の定義（九州他県ファイルで採用）】",BOLD),
       ("  主定義 = 在宅薬学総合体制加算1 or 2 の届出（=本気の在宅体制）",NORM),
       ("  広義  = 在宅患者訪問薬剤管理指導料 の届出（届出のみ・休眠含む）",NORM),
       ("",NORM),
       ("【名寄せ方針】",BOLD),
       ("  福岡=開設者（法人名）をキーに名寄せ。大手/主要チェーンは辞書(01)で正規化、未一致は法人名コアで独立計上。",NORM),
       ("  全マッピングは 02_名寄せ履歴 で監査可能（原名称→ルールID→正規化事業者→件数）。",NORM),
       ("",NORM),
       ("【シート構成】",BOLD),
       ("  01_名寄せ辞書 / 02_名寄せ履歴 / 10_ローデータ / 20_集計_事業者別（COUNTIFS）",NORM),
       ("  集計値はローデータをCOUNTIFS参照する数式。ローデータを修正すると集計も自動更新。",NORM),
       ]
    for i,(t,f) in enumerate(L,1):
        c=ws.cell(row=i,column=1,value=t); c.font=f
    ws.column_dimensions["A"].width=110

    add_dict_sheet(wb)
    add_history_sheet(wb,hist,"開設者(法人名)")

    # 10_ローデータ
    wd=wb.create_sheet("10_ローデータ")
    heads=["No","都道府県","保険薬局コード","薬局名称","所在地","開設者（原文）","開設者コア","適用ルールID","正規化事業者名","事業者タイプ","信頼度"]
    for j,h in enumerate(heads,1): wd.cell(row=1,column=j,value=h)
    style_header(wd,1,len(heads))
    for i,r in enumerate(rows,1):
        vals=[i,r["pref"],r["code"],r["name"],r["addr"],r["raw"],clean_corp(r["raw"]),r["rid"],r["corp"],r["type"],r["conf"]]
        for j,v in enumerate(vals,1):
            cell=wd.cell(row=i+1,column=j,value=v); cell.font=NORM; cell.border=BORD
    wd.freeze_panes="A2"
    widths=[5,9,14,30,40,34,24,10,30,14,9]
    for j,w in enumerate(widths,1): wd.column_dimensions[get_column_letter(j)].width=w
    nrow=len(rows)+1
    corp_col="I"  # 正規化事業者名

    # 20_集計_事業者別（COUNTIFS）
    ws=wb.create_sheet("20_集計_事業者別")
    ws["A1"]="福岡県 事業者別 店舗数（全保険薬局・在宅フィルタ無し）"; ws["A1"].font=TITLE
    ws["A2"]=f"店舗数列は =COUNTIF('10_ローデータ'!{corp_col}:{corp_col}, 事業者名) の数式。並びは大手優先＋店舗数降順。"
    ws["A2"].font=NORM
    heads=["順位","正規化事業者名","事業者タイプ","信頼度","店舗数","シェア"]
    for j,h in enumerate(heads,1): ws.cell(row=4,column=j,value=h)
    style_header(ws,4,len(heads))
    # 事業者一覧（タイプ・信頼度付き、件数降順）
    agg=OrderedDict()
    for r in rows:
        k=r["corp"]
        if k not in agg: agg[k]=dict(type=r["type"],conf=r["conf"],cnt=0,rid=r["rid"])
        agg[k]["cnt"]+=1
    order=sorted(agg.items(),key=lambda kv:(kv[1]["rid"]=="-",-kv[1]["cnt"]))
    r=5
    for rank,(corp,info) in enumerate(order,1):
        ws.cell(row=r,column=1,value=rank)
        ws.cell(row=r,column=2,value=corp)
        ws.cell(row=r,column=3,value=info["type"])
        ws.cell(row=r,column=4,value=info["conf"])
        ws.cell(row=r,column=5,value=f"=COUNTIF('10_ローデータ'!{corp_col}:{corp_col},B{r})")
        ws.cell(row=r,column=6,value=f'=E{r}/$E${5+len(order)+1}')
        ws.cell(row=r,column=6).number_format='0.0%'
        for c in range(1,7):
            ws.cell(row=r,column=c).font=NORM; ws.cell(row=r,column=c).border=BORD
        if info["rid"]!="-" and info["conf"]!="要確認":
            for c in range(1,7): ws.cell(row=r,column=c).fill=GFILL
        if info["conf"]=="要確認":
            for c in range(1,7): ws.cell(row=r,column=c).fill=YFILL
        r+=1
    ws.cell(row=r,column=2,value="合計（検算）").font=BOLD
    ws.cell(row=r,column=5,value=f"=SUM(E5:E{r-1})").font=BOLD
    ws.cell(row=r+1,column=2,value="ローデータ行数").font=NORM
    ws.cell(row=r+1,column=5,value=f"=COUNTA('10_ローデータ'!C2:C{nrow})").font=NORM
    widths=[6,30,16,10,10,10]
    for j,w in enumerate(widths,1): ws.column_dimensions[get_column_letter(j)].width=w
    ws.freeze_panes="A5"

    path=f"{OUT}/福岡県_事業者集計.xlsx"; wb.save(path)
    return path,len(rows),len(order)

# ================= 九州在宅 ワークブック =================
def build_kyushu():
    rows=parse_shisetsu()
    # フラグ展開
    for r in rows:
        fl=r["flags"]
        r["t1"]=1 if "在宅薬学総合体制加算１" in fl else 0
        r["t2"]=1 if "在宅薬学総合体制加算２" in fl else 0
        r["taisei"]=1 if (r["t1"] or r["t2"]) else 0
        r["homev"]=1 if HOMEV in fl else 0
    # 在宅薬局のみ（主定義: 在薬総1/2）に名寄せ＆集計、ただしローデータは全件保持しフラグで判別
    hist=normalize(rows,"name")
    wb=openpyxl.Workbook(); wb.remove(wb.active)

    pref_list=list(SHISETSU.keys())
    ws=wb.create_sheet("00_README")
    L=[("九州（福岡除く7県）在宅薬局 事業者集計（A案）",TITLE),
       ("",NORM),
       ("【元データ】令和8年5月1日現在 施設基準届出状況（保険薬局）／九州厚生局",BOLD),
       (f"  出典URL: {SRC_URL}",link_font()),
       ("  対象県: 佐賀・長崎・熊本・大分・宮崎・鹿児島・沖縄（7県）",NORM),
       ("  ※福岡県は別種別ファイル（基本薬局一覧）のため在宅フラグ無し→本集計から除外。別途『福岡県_事業者集計.xlsx』参照。",NORM),
       ("  ファイル: r8_05_shisetsu_(各県)_yakkyoku_02.xlsx",NORM),
       ("",NORM),
       ("【在宅の定義】",BOLD),
       ("  ◎主定義 = 在宅薬学総合体制加算1 or 2 の届出（本気の在宅体制。集計20/21の既定軸）",NORM),
       ("  ○広義   = 在宅患者訪問薬剤管理指導料 の届出（届出のみ・休眠含む。参考列）",NORM),
       ("  在総1=15→30点(R8改定)／在総2=個人宅100点・施設50点。詳細はローデータの在総1・在総2列で判別可。",NORM),
       ("",NORM),
       ("【名寄せ方針】",BOLD),
       ("  施設基準ファイルには開設者(法人)列が無いため、屋号（医療機関名称）をキーに名寄せ。",NORM),
       ("  大手/主要チェーンは辞書(01)で正規化、未一致は屋号コア（…薬局まで）で独立計上。",NORM),
       ("  全マッピングは 02_名寄せ履歴 で監査可能。汎用名（すこやか薬局/さくら調剤等）は黄=要確認で明示。",NORM),
       ("",NORM),
       ("【シート構成】",BOLD),
       ("  01_名寄せ辞書 / 02_名寄せ履歴 / 10_ローデータ(全薬局・在宅フラグ付) /",NORM),
       ("  20_集計_事業者別(在総ベース) / 21_集計_県別サマリ。集計は全てCOUNTIFS数式でローデータ参照。",NORM),
       ]
    for i,(t,f) in enumerate(L,1):
        ws.cell(row=i,column=1,value=t).font=f
    ws.column_dimensions["A"].width=112

    add_dict_sheet(wb)
    add_history_sheet(wb,hist,"屋号(医療機関名称)")

    # 10_ローデータ（全薬局・フラグ付）
    wd=wb.create_sheet("10_ローデータ")
    heads=["No","都道府県","保険薬局コード","薬局名称","所在地","在総1","在総2","在総(1or2)","在薬(訪問届出)","屋号コア","適用ルールID","正規化事業者名","事業者タイプ","信頼度"]
    for j,h in enumerate(heads,1): wd.cell(row=1,column=j,value=h)
    style_header(wd,1,len(heads))
    for i,r in enumerate(rows,1):
        vals=[i,r["pref"],r["code"],r["name"],r["addr"],r["t1"],r["t2"],r["taisei"],r["homev"],
              base_yago(r["name"]),r["rid"],r["corp"],r["type"],r["conf"]]
        for j,v in enumerate(vals,1):
            cell=wd.cell(row=i+1,column=j,value=v); cell.font=NORM; cell.border=BORD
    wd.freeze_panes="A2"
    widths=[5,9,13,30,38,7,7,10,13,22,10,30,14,9]
    for j,w in enumerate(widths,1): wd.column_dimensions[get_column_letter(j)].width=w
    nrow=len(rows)+1
    corp_col="L"; pref_col="B"; taisei_col="H"; homev_col="I"

    # 20_集計_事業者別（在総ベース, COUNTIFS）
    ws=wb.create_sheet("20_集計_事業者別")
    ws["A1"]="九州7県 在宅薬局 事業者別 店舗数（在宅薬学総合体制加算ベース）"; ws["A1"].font=TITLE
    ws["A2"]=(f"在総店舗数 = COUNTIFS(事業者一致, かつ 在総=1)。広義(在薬)列も併記。並びは大手優先＋在総数降順。")
    ws["A2"].font=NORM
    heads=["順位","正規化事業者名","事業者タイプ","信頼度","在総薬局数","広義(在薬)薬局数","在総シェア"]
    for j,h in enumerate(heads,1): ws.cell(row=4,column=j,value=h)
    style_header(ws,4,len(heads))
    agg=OrderedDict()
    for r in rows:
        if not r["taisei"]:
            # 集計の行ラベルは在総保有事業者を基本にするが、広義のみの事業者も拾う
            pass
        k=r["corp"]
        if k not in agg: agg[k]=dict(type=r["type"],conf=r["conf"],rid=r["rid"],taisei=0,homev=0)
        agg[k]["taisei"]+=r["taisei"]; agg[k]["homev"]+=r["homev"]
    # 在総>0 の事業者のみ表示（在宅薬局事業者の内訳）
    order=[(k,v) for k,v in agg.items() if v["taisei"]>0]
    order=sorted(order,key=lambda kv:(kv[1]["rid"]=="-",-kv[1]["taisei"]))
    r=5
    for rank,(corp,info) in enumerate(order,1):
        ws.cell(row=r,column=1,value=rank)
        ws.cell(row=r,column=2,value=corp)
        ws.cell(row=r,column=3,value=info["type"])
        ws.cell(row=r,column=4,value=info["conf"])
        ws.cell(row=r,column=5,value=f"=COUNTIFS('10_ローデータ'!{corp_col}:{corp_col},B{r},'10_ローデータ'!{taisei_col}:{taisei_col},1)")
        ws.cell(row=r,column=6,value=f"=COUNTIFS('10_ローデータ'!{corp_col}:{corp_col},B{r},'10_ローデータ'!{homev_col}:{homev_col},1)")
        ws.cell(row=r,column=7,value=f'=E{r}/$E${5+len(order)+1}'); ws.cell(row=r,column=7).number_format='0.0%'
        for c in range(1,8):
            ws.cell(row=r,column=c).font=NORM; ws.cell(row=r,column=c).border=BORD
        if info["rid"]!="-" and info["conf"]!="要確認":
            for c in range(1,8): ws.cell(row=r,column=c).fill=GFILL
        if info["conf"]=="要確認":
            for c in range(1,8): ws.cell(row=r,column=c).fill=YFILL
        r+=1
    ws.cell(row=r,column=2,value="合計（検算）").font=BOLD
    ws.cell(row=r,column=5,value=f"=SUM(E5:E{r-1})").font=BOLD
    ws.cell(row=r,column=6,value=f"=SUM(F5:F{r-1})").font=BOLD
    ws.cell(row=r+1,column=2,value="在総薬局 総数（ローデータ）").font=NORM
    ws.cell(row=r+1,column=5,value=f"=COUNTIF('10_ローデータ'!{taisei_col}2:{taisei_col}{nrow},1)").font=NORM
    widths=[6,32,16,10,12,15,11]
    for j,w in enumerate(widths,1): ws.column_dimensions[get_column_letter(j)].width=w
    ws.freeze_panes="A5"

    # 21_集計_県別サマリ
    ws=wb.create_sheet("21_集計_県別サマリ")
    ws["A1"]="県別サマリ（在宅薬局数・在宅実施率）"; ws["A1"].font=TITLE
    ws["A2"]="各数値はローデータをCOUNTIFS参照。実施率=在総薬局数/全薬局数。"; ws["A2"].font=NORM
    heads=["都道府県","全薬局数","在総薬局数(本気)","在総実施率","広義(在薬)薬局数","広義実施率"]
    for j,h in enumerate(heads,1): ws.cell(row=4,column=j,value=h)
    style_header(ws,4,len(heads))
    r=5
    for pref in pref_list:
        ws.cell(row=r,column=1,value=pref)
        ws.cell(row=r,column=2,value=f"=COUNTIF('10_ローデータ'!{pref_col}:{pref_col},A{r})")
        ws.cell(row=r,column=3,value=f"=COUNTIFS('10_ローデータ'!{pref_col}:{pref_col},A{r},'10_ローデータ'!{taisei_col}:{taisei_col},1)")
        ws.cell(row=r,column=4,value=f'=C{r}/B{r}'); ws.cell(row=r,column=4).number_format='0.0%'
        ws.cell(row=r,column=5,value=f"=COUNTIFS('10_ローデータ'!{pref_col}:{pref_col},A{r},'10_ローデータ'!{homev_col}:{homev_col},1)")
        ws.cell(row=r,column=6,value=f'=E{r}/B{r}'); ws.cell(row=r,column=6).number_format='0.0%'
        for c in range(1,7):
            ws.cell(row=r,column=c).font=NORM; ws.cell(row=r,column=c).border=BORD
        r+=1
    ws.cell(row=r,column=1,value="7県合計").font=BOLD
    for col,L in [(2,"B"),(3,"C"),(5,"E")]:
        ws.cell(row=r,column=col,value=f"=SUM({L}5:{L}{r-1})").font=BOLD
    ws.cell(row=r,column=4,value=f"=C{r}/B{r}").font=BOLD; ws.cell(row=r,column=4).number_format='0.0%'
    ws.cell(row=r,column=6,value=f"=E{r}/B{r}").font=BOLD; ws.cell(row=r,column=6).number_format='0.0%'
    widths=[12,12,16,12,16,12]
    for j,w in enumerate(widths,1): ws.column_dimensions[get_column_letter(j)].width=w

    path=f"{OUT}/九州7県_在宅薬局_事業者集計.xlsx"; wb.save(path)
    return path,len(rows),len(order)

if __name__=="__main__":
    p1,n1,o1=build_fukuoka()
    print(f"福岡: {p1}  薬局{n1}件 事業者{o1}")
    p2,n2,o2=build_kyushu()
    print(f"九州: {p2}  薬局{n2}件 在宅事業者{o2}")
