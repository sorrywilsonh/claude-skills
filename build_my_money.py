#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""產生「my錢錢」資產追蹤 .xlsx：上傳 Google Drive 後用 Google Sheets 開啟即可。
分頁：儀表板 / 設定 / 帳戶總覽 / 持股明細 / 交易明細 / 資產快照 / 現金流。
標色規則：✍️黃底=需手動填寫；🔒灰底=自動計算（勿改）。
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import PieChart, LineChart, BarChart, Reference
from openpyxl.utils import get_column_letter

OUT = "/Users/ucpc/Google Drive/我的雲端硬碟/my錢錢.xlsx"

# ---- 樣式 ----
TITLE   = Font(name="Arial", size=16, bold=True, color="1F3864")
HEAD_M  = Font(name="Arial", size=10, bold=True, color="7F6000")   # 手動欄位標題字
HEAD_A  = Font(name="Arial", size=10, bold=True, color="3F3F3F")   # 自動欄位標題字
HEAD_AW = Font(name="Arial", size=10, bold=True, color="FFFFFF")
NOTE    = Font(name="Arial", size=9, italic=True, color="808080")
BOLDB   = Font(name="Arial", size=11, bold=True, color="1F3864")

FILL_MANHEAD = PatternFill("solid", fgColor="FFE599")  # 手動標題 深黃
FILL_AUTHEAD = PatternFill("solid", fgColor="808080")   # 自動標題 灰
FILL_MAN     = PatternFill("solid", fgColor="FFF9E6")   # 手動資料 淡黃
FILL_AUT     = PatternFill("solid", fgColor="F0F0F0")   # 自動資料 淡灰
FILL_KPI     = PatternFill("solid", fgColor="DDEBF7")
FILL_LEGEND  = PatternFill("solid", fgColor="FCE4D6")

THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CTR = Alignment(horizontal="center", vertical="center")
LFT = Alignment(horizontal="left", vertical="center", wrap_text=True)

wb = openpyxl.Workbook()

def legend(ws, row=2):
    ws.cell(row=row, column=1, value="✍️ 黃底 = 需手動填寫／更新").fill = FILL_MAN
    ws.cell(row=row, column=1).font = NOTE
    ws.cell(row=row, column=3, value="🔒 灰底 = 自動計算（請勿手動修改）").fill = FILL_AUT
    ws.cell(row=row, column=3).font = NOTE

def headers(ws, row, cols):
    """cols = list of (title, 'M'|'A', width)"""
    for i, (title, kind, width) in enumerate(cols, start=1):
        c = ws.cell(row=row, column=i)
        if kind == "M":
            c.value = "✍️ " + title; c.fill = FILL_MANHEAD; c.font = HEAD_M
        else:
            c.value = "🔒 " + title; c.fill = FILL_AUTHEAD; c.font = HEAD_AW
        c.alignment = CTR; c.border = BORDER
        ws.column_dimensions[get_column_letter(i)].width = width

def fill_col(ws, col, r1, r2, kind, numfmt=None):
    f = FILL_MAN if kind == "M" else FILL_AUT
    for r in range(r1, r2 + 1):
        c = ws.cell(row=r, column=col)
        c.fill = f; c.border = BORDER
        if numfmt: c.number_format = numfmt

NT = '#,##0'        # 台幣整數
NT2 = '#,##0.00'    # 兩位
PCT = '0.0%'
RATE_RANGE = "'設定'!$A$5:$B$8"   # 幣別→台幣匯率 查表範圍

# ============================================================ 設定
s = wb.active; s.title = "設定"
s["A1"] = "⚙️ 設定（請先在這裡建立你的帳戶清單）"; s["A1"].font = TITLE
legend(s)
# 匯率表
s["A4"] = "幣別"; s["B4"] = "對台幣匯率（自動，需網路）"
for c in ("A4", "B4"): s[c].fill = FILL_AUTHEAD; s[c].font = HEAD_AW; s[c].alignment = CTR
rates = [("TWD", 1), ("USD", '=GOOGLEFINANCE("CURRENCY:USDTWD")'),
         ("JPY", '=GOOGLEFINANCE("CURRENCY:JPYTWD")')]
for i, (cur, val) in enumerate(rates, start=5):
    s.cell(row=i, column=1, value=cur).fill = FILL_AUT
    b = s.cell(row=i, column=2, value=val); b.fill = FILL_AUT; b.number_format = NT2
s["A8"].fill = FILL_AUT; s["B8"].fill = FILL_AUT  # 預留一列
s["A10"] = "💡 想新增幣別：在上表加一列並填 GOOGLEFINANCE 匯率公式"; s["A10"].font = NOTE
s["A11"] = "💡 同一帳戶有多種幣別時，請分多列填（如富邦銀行 TWD 一列、USD 一列）"; s["A11"].font = NOTE

# 帳戶清單（手動）
s["D4"] = "✍️ 帳戶名稱"; s["E4"] = "✍️ 類型"; s["F4"] = "✍️ 幣別"
for c in ("D4", "E4", "F4"): s[c].fill = FILL_MANHEAD; s[c].font = HEAD_M; s[c].alignment = CTR
example_accts = [("台新銀行", "銀行", "TWD"), ("富邦銀行", "銀行", "TWD"),
                 ("富邦銀行", "銀行", "USD"), ("國泰證券", "證券", "TWD"),
                 ("Firstrade", "Firstrade", "USD"), ("基富通基金", "基金", "TWD")]
for i, (n, t, cur) in enumerate(example_accts, start=5):
    s.cell(row=i, column=4, value=n); s.cell(row=i, column=5, value=t); s.cell(row=i, column=6, value=cur)
fill_col(s, 4, 5, 54, "M"); fill_col(s, 5, 5, 54, "M"); fill_col(s, 6, 5, 54, "M")
for w, col in [(16, "D"), (12, "E"), (8, "F")]: s.column_dimensions[col].width = w

# 各種清單（給下拉選單用）
lists = {
    "H": ("類型清單", ["銀行", "證券", "基金", "Firstrade", "其他"]),
    "J": ("幣別清單", ["TWD", "USD", "JPY"]),
    "L": ("資產類別", ["現金", "台股", "美股", "ETF", "基金", "其他"]),
    "N": ("交易類型", ["買", "賣", "股息", "入金", "出金", "手續費", "稅"]),
}
for col, (title, items) in lists.items():
    s[f"{col}4"] = title; s[f"{col}4"].fill = FILL_AUTHEAD; s[f"{col}4"].font = HEAD_AW
    for i, it in enumerate(items, start=5):
        s[f"{col}{i}"] = it
    s.column_dimensions[col].width = 11

ACCT_LIST = "'設定'!$D$5:$D$54"
TYPE_LIST = "'設定'!$H$5:$H$9"
CUR_LIST  = "'設定'!$J$5:$J$7"
TXN_LIST  = "'設定'!$N$5:$N$11"

# ============================================================ 交易明細
t = wb.create_sheet("交易明細")
t["A1"] = "🧾 交易明細（核心流水帳：每筆買/賣/股息/入金/出金都記一列）"; t["A1"].font = TITLE
legend(t)
t["A3"] = "💡 入金/出金/股息：把金額填在「價格」、股數填 1 即可。"; t["A3"].font = NOTE
txn_cols = [("日期","M",12),("帳戶","M",14),("類型","M",10),("代號","M",12),("名稱","M",14),
            ("股數","M",10),("價格","M",10),("幣別","M",8),("手續費","M",10),("稅","M",9),
            ("金額小計(原幣)","A",13),("台幣金額","A",13),("手續費稅(台幣)","A",14),
            ("年月","A",9),("備註","M",20)]
headers(t, 4, txn_cols)
T_R1, T_R2 = 5, 504
for r in range(T_R1, T_R2 + 1):
    t.cell(row=r, column=11, value=f"=IF($F{r}=\"\",\"\",$F{r}*$G{r})")   # K 金額小計原幣
    t.cell(row=r, column=12, value=f"=IF($K{r}=\"\",\"\",$K{r}*VLOOKUP($H{r},{RATE_RANGE},2,FALSE))")  # L 台幣金額
    t.cell(row=r, column=13, value=f"=IF(AND($I{r}=\"\",$J{r}=\"\"),\"\",(N($I{r})+N($J{r}))*VLOOKUP($H{r},{RATE_RANGE},2,FALSE))")  # M 手續費稅台幣
    t.cell(row=r, column=14, value=f"=IF($A{r}=\"\",\"\",TEXT($A{r},\"yyyy-mm\"))")  # N 年月
for col in (1,2,3,4,5,6,7,8,9,10,15): fill_col(t, col, T_R1, T_R2, "M")
for col in (11,12,13,14): fill_col(t, col, T_R1, T_R2, "A")
for col,fmt in [(1,'yyyy-mm-dd'),(6,NT2),(7,NT2),(9,NT2),(10,NT2),(11,NT2),(12,NT),(13,NT)]:
    for r in range(T_R1, T_R2+1): t.cell(row=r, column=col).number_format = fmt
# 範例列
ex = ["2026-01-05","國泰證券","買","TPE:2330","台積電",10,600,"TWD",20,0]
for i,v in enumerate(ex, start=1): t.cell(row=5, column=i, value=v)
ex2 = ["2026-02-10","Firstrade","買","VOO","Vanguard S&P500",5,480,"USD",0,0]
for i,v in enumerate(ex2, start=1): t.cell(row=6, column=i, value=v)
ex3 = ["2026-03-01","台新銀行","入金","","薪轉",1,50000,"TWD","",""]
for i,v in enumerate(ex3, start=1): t.cell(row=7, column=i, value=v)
# 下拉
t.add_data_validation((dv := DataValidation(type="list", formula1=ACCT_LIST, allow_blank=True))); dv.add(f"B{T_R1}:B{T_R2}")
t.add_data_validation((dv := DataValidation(type="list", formula1=TXN_LIST,  allow_blank=True))); dv.add(f"C{T_R1}:C{T_R2}")
t.add_data_validation((dv := DataValidation(type="list", formula1=CUR_LIST,  allow_blank=True))); dv.add(f"H{T_R1}:H{T_R2}")
t.freeze_panes = "A5"

# ============================================================ 持股明細
h = wb.create_sheet("持股明細")
h["A1"] = "📈 持股明細（股數與平均成本由交易明細自動加權計算）"; h["A1"].font = TITLE
legend(h)
h["A3"] = "💡 新增持股時，只需手動填：帳戶 / 代號 / 名稱 / 幣別。其餘自動。"; h["A3"].font = NOTE
hold_cols = [("帳戶","M",14),("代號","M",12),("名稱","M",16),("幣別","M",8),
             ("股數","A",10),("平均成本","A",11),("即時股價","A",11),("市值(原幣)","A",13),
             ("台幣市值","A",13),("投入成本(原幣)","A",14),("未實現損益(原幣)","A",15),
             ("報酬率%","A",10),("台幣投入成本","A",14),("台幣未實現損益","A",15)]
headers(h, 4, hold_cols)
H_R1, H_R2 = 5, 54
TB = f"'交易明細'!$B${T_R1}:$B${T_R2}"   # 帳戶
TD = f"'交易明細'!$D${T_R1}:$D${T_R2}"   # 代號
TC = f"'交易明細'!$C${T_R1}:$C${T_R2}"   # 類型
TF = f"'交易明細'!$F${T_R1}:$F${T_R2}"   # 股數
TK = f"'交易明細'!$K${T_R1}:$K${T_R2}"   # 金額小計原幣(=股數*價格)
TI = f"'交易明細'!$I${T_R1}:$I${T_R2}"   # 手續費
TJ = f"'交易明細'!$J${T_R1}:$J${T_R2}"   # 稅
for r in range(H_R1, H_R2 + 1):
    A, B = f"$A{r}", f"$B{r}"
    buyshares = f"SUMIFS({TF},{TB},{A},{TD},{B},{TC},\"買\")"
    sellshares = f"SUMIFS({TF},{TB},{A},{TD},{B},{TC},\"賣\")"
    # 買入總成本 = 買入金額小計 + 買入手續費 + 買入稅（全部只算「買」）
    buycost = (f"SUMIFS({TK},{TB},{A},{TD},{B},{TC},\"買\")"
               f"+SUMIFS({TI},{TB},{A},{TD},{B},{TC},\"買\")+SUMIFS({TJ},{TB},{A},{TD},{B},{TC},\"買\")")
    h.cell(row=r, column=5,  value=f"=IF($B{r}=\"\",\"\",{buyshares}-{sellshares})")             # 股數
    h.cell(row=r, column=6,  value=f"=IF($E{r}=\"\",\"\",IFERROR(({buycost})/{buyshares},0))")    # 平均成本
    h.cell(row=r, column=7,  value=f"=IF($B{r}=\"\",\"\",IFERROR(GOOGLEFINANCE($B{r}),0))")        # 即時股價
    h.cell(row=r, column=8,  value=f"=IF($E{r}=\"\",\"\",$E{r}*$G{r})")                            # 市值原幣
    h.cell(row=r, column=9,  value=f"=IF($H{r}=\"\",\"\",$H{r}*VLOOKUP($D{r},{RATE_RANGE},2,FALSE))")  # 台幣市值
    h.cell(row=r, column=10, value=f"=IF($E{r}=\"\",\"\",$E{r}*$F{r})")                            # 投入成本原幣
    h.cell(row=r, column=11, value=f"=IF($H{r}=\"\",\"\",$H{r}-$J{r})")                            # 未實現損益原幣
    h.cell(row=r, column=12, value=f"=IF($J{r}=\"\",\"\",IFERROR($K{r}/$J{r},0))")                 # 報酬率%
    h.cell(row=r, column=13, value=f"=IF($J{r}=\"\",\"\",$J{r}*VLOOKUP($D{r},{RATE_RANGE},2,FALSE))")  # 台幣投入成本
    h.cell(row=r, column=14, value=f"=IF($I{r}=\"\",\"\",$I{r}-$M{r})")                            # 台幣未實現損益
for col in (1,2,3,4): fill_col(h, col, H_R1, H_R2, "M")
for col in range(5,15): fill_col(h, col, H_R1, H_R2, "A")
for col,fmt in [(5,NT2),(6,NT2),(7,NT2),(8,NT2),(9,NT),(10,NT2),(11,NT2),(12,PCT),(13,NT),(14,NT)]:
    for r in range(H_R1, H_R2+1): h.cell(row=r, column=col).number_format = fmt
# 範例持股
h.cell(row=5, column=1, value="國泰證券"); h.cell(row=5, column=2, value="TPE:2330"); h.cell(row=5, column=3, value="台積電"); h.cell(row=5, column=4, value="TWD")
h.cell(row=6, column=1, value="Firstrade"); h.cell(row=6, column=2, value="VOO"); h.cell(row=6, column=3, value="Vanguard S&P500"); h.cell(row=6, column=4, value="USD")
h.add_data_validation((dv := DataValidation(type="list", formula1=ACCT_LIST, allow_blank=True))); dv.add(f"A{H_R1}:A{H_R2}")
h.add_data_validation((dv := DataValidation(type="list", formula1=CUR_LIST,  allow_blank=True))); dv.add(f"D{H_R1}:D{H_R2}")
h.freeze_panes = "A5"

# ============================================================ 帳戶總覽
a = wb.create_sheet("帳戶總覽")
a["A1"] = "🏦 帳戶總覽（同一帳戶若有多種幣別，請分多列：帳戶相同、幣別不同）"; a["A1"].font = TITLE
legend(a)
a["A3"] = "💡 類型自動帶入；幣別請手動選；持股市值依『帳戶＋幣別』分別加總。"; a["A3"].font = NOTE
acc_cols = [("帳戶名稱","M",16),("類型","A",10),("幣別","M",8),("現金餘額(原幣)","M",14),
            ("持股市值(台幣)","A",14),("台幣總值","A",14),("佔比%","A",10)]
headers(a, 4, acc_cols)
A_R1, A_R2 = 5, 24
TOTAL_ROW = 26
for r in range(A_R1, A_R2 + 1):
    a.cell(row=r, column=2, value=f"=IFERROR(VLOOKUP($A{r},'設定'!$D$5:$F$54,2,FALSE),\"\")")  # 類型自動
    # 持股市值：依『帳戶＋幣別』加總（同帳戶不同幣別不會互相混到）
    a.cell(row=r, column=5, value=f"=IF($A{r}=\"\",\"\",SUMIFS('持股明細'!$I${H_R1}:$I${H_R2},'持股明細'!$A${H_R1}:$A${H_R2},$A{r},'持股明細'!$D${H_R1}:$D${H_R2},$C{r}))")
    a.cell(row=r, column=6, value=f"=IF($A{r}=\"\",\"\",N($D{r})*VLOOKUP($C{r},{RATE_RANGE},2,FALSE)+N($E{r}))")
    a.cell(row=r, column=7, value=f"=IFERROR($F{r}/$F${TOTAL_ROW},\"\")")
for col in (1,3,4): fill_col(a, col, A_R1, A_R2, "M")   # 帳戶/幣別/現金餘額 手動
for col in (2,5,6,7): fill_col(a, col, A_R1, A_R2, "A")  # 類型/持股市值/總值/佔比 自動
for col,fmt in [(4,NT),(5,NT),(6,NT),(7,PCT)]:
    for r in range(A_R1, A_R2+1): a.cell(row=r, column=col).number_format = fmt
a.cell(row=TOTAL_ROW, column=1, value="總計").font = BOLDB
tot = a.cell(row=TOTAL_ROW, column=6, value=f"=SUM(F{A_R1}:F{A_R2})"); tot.font = BOLDB; tot.number_format = NT; tot.fill = FILL_KPI
# 範例：富邦銀行同時有 TWD 與 USD，分兩列
examples = [("台新銀行","TWD",100000),("富邦銀行","TWD",50000),("富邦銀行","USD",3000),
            ("國泰證券","TWD",0),("Firstrade","USD",0)]
for i,(nm,cur,bal) in enumerate(examples):
    a.cell(row=A_R1+i, column=1, value=nm); a.cell(row=A_R1+i, column=3, value=cur); a.cell(row=A_R1+i, column=4, value=bal)
a.add_data_validation((dv := DataValidation(type="list", formula1=ACCT_LIST, allow_blank=True))); dv.add(f"A{A_R1}:A{A_R2}")
a.add_data_validation((dv := DataValidation(type="list", formula1=CUR_LIST,  allow_blank=True))); dv.add(f"C{A_R1}:C{A_R2}")
a.freeze_panes = "A5"

# ============================================================ 資產快照
sn = wb.create_sheet("資產快照")
sn["A1"] = "📅 資產快照（每月把儀表板的總淨資產『貼數值』記一列 → 畫趨勢線）"; sn["A1"].font = TITLE
legend(sn)
sn["A3"] = "💡 請貼『數值』不要貼公式，否則歷史會跟著變動。"; sn["A3"].font = NOTE
snap_cols = [("日期","M",14),("淨資產合計(台幣)","M",18),("備註","M",24)]
headers(sn, 4, snap_cols)
S_R1, S_R2 = 5, 64
fill_col(sn, 1, S_R1, S_R2, "M", 'yyyy-mm-dd')
fill_col(sn, 2, S_R1, S_R2, "M", NT)
fill_col(sn, 3, S_R1, S_R2, "M")
sn.cell(row=5, column=1, value="2026-01-31"); sn.cell(row=5, column=2, value=1200000)
sn.cell(row=6, column=1, value="2026-02-28"); sn.cell(row=6, column=2, value=1255000)
sn.cell(row=7, column=1, value="2026-03-31"); sn.cell(row=7, column=2, value=1310000)
sn.freeze_panes = "A5"

# ============================================================ 現金流
cf = wb.create_sheet("現金流")
cf["A1"] = "💰 現金流（金額自動從交易明細彙整成台幣 → 畫長條圖）"; cf["A1"].font = TITLE
legend(cf)
cf["A3"] = "💡 只需手動填『年月』(如 2026-06)，其餘自動。"; cf["A3"].font = NOTE
cf_cols = [("年月","M",10),("入金","A",12),("股息","A",12),("出金/提領","A",12),
           ("手續費+稅","A",12),("淨現金流","A",13)]
headers(cf, 4, cf_cols)
C_R1, C_R2 = 5, 40
TL = f"'交易明細'!$N${T_R1}:$N${T_R2}"   # 年月
TLamt = f"'交易明細'!$L${T_R1}:$L${T_R2}" # 台幣金額
TM = f"'交易明細'!$M${T_R1}:$M${T_R2}"   # 手續費稅台幣
for r in range(C_R1, C_R2 + 1):
    A = f"$A{r}"
    cf.cell(row=r, column=2, value=f"=IF($A{r}=\"\",\"\",SUMIFS({TLamt},{TL},{A},{TC},\"入金\"))")
    cf.cell(row=r, column=3, value=f"=IF($A{r}=\"\",\"\",SUMIFS({TLamt},{TL},{A},{TC},\"股息\"))")
    cf.cell(row=r, column=4, value=f"=IF($A{r}=\"\",\"\",SUMIFS({TLamt},{TL},{A},{TC},\"出金\"))")
    cf.cell(row=r, column=5, value=f"=IF($A{r}=\"\",\"\",SUMIFS({TM},{TL},{A}))")
    cf.cell(row=r, column=6, value=f"=IF($A{r}=\"\",\"\",N($B{r})+N($C{r})-N($D{r})-N($E{r}))")
fill_col(cf, 1, C_R1, C_R2, "M")
for col in (2,3,4,5,6): fill_col(cf, col, C_R1, C_R2, "A", NT)
cf.cell(row=5, column=1, value="2026-01"); cf.cell(row=6, column=1, value="2026-02"); cf.cell(row=7, column=1, value="2026-03")
cf.freeze_panes = "A5"

# ============================================================ 儀表板
d = wb.create_sheet("儀表板")
wb.move_sheet("儀表板", -(len(wb.sheetnames)-1))  # 移到最前
d["A1"] = "📊 my錢錢 — 資產儀表板"; d["A1"].font = Font(name="Arial", size=20, bold=True, color="1F3864")
d["A2"] = "上傳 Google Drive → 右鍵『用 Google Sheets 開啟』，即時股價/匯率會自動生效。"; d["A2"].font = NOTE
# KPI
kpis = [("總淨資產 (台幣)", "='帳戶總覽'!$F$26", NT),
        ("總未實現損益 (台幣)", f"=SUM('持股明細'!$N${H_R1}:$N${H_R2})", NT),
        ("總報酬率", f"=IFERROR(SUM('持股明細'!$N${H_R1}:$N${H_R2})/SUM('持股明細'!$M${H_R1}:$M${H_R2}),0)", PCT),
        ("最近月份淨現金流", "='現金流'!$F$5", NT)]
for i, (label, formula, fmt) in enumerate(kpis):
    col = 1 + i * 3
    lc = d.cell(row=4, column=col, value=label); lc.font = HEAD_AW; lc.fill = FILL_AUTHEAD; lc.alignment = CTR
    d.merge_cells(start_row=4, start_column=col, end_row=4, end_column=col+1)
    vc = d.cell(row=5, column=col, value=formula); vc.font = Font(name="Arial", size=16, bold=True, color="1F3864")
    vc.fill = FILL_KPI; vc.number_format = fmt; vc.alignment = CTR
    d.merge_cells(start_row=5, start_column=col, end_row=6, end_column=col+1)
    d.column_dimensions[get_column_letter(col)].width = 16

# 手動維護清單
d["A9"] = "🔧 需要你『手動更新』的項目（黃底欄位）"; d["A9"].font = BOLDB; d["A9"].fill = FILL_LEGEND
manual_items = [
    "① 設定分頁：建立『帳戶清單』（帳戶名稱／類型／幣別）— 一次性",
    "② 帳戶總覽：各帳戶的『現金餘額(原幣)』— 不定期更新",
    "③ 持股明細：新增持股時填『帳戶／代號／名稱／幣別』(股數與成本自動)",
    "④ 交易明細：每筆買/賣/股息/入金/出金記一列 — 平時主要工作",
    "⑤ 資產快照：每月把上方『總淨資產』貼成數值一列",
    "⑥ 現金流：填入『年月』(如 2026-06)，金額自動算",
]
for i, txt in enumerate(manual_items, start=10):
    d.cell(row=i, column=1, value=txt).font = Font(name="Arial", size=10, color="7F6000")
    d.cell(row=i, column=1).fill = FILL_MAN
    d.merge_cells(start_row=i, start_column=1, end_row=i, end_column=6)

# ---- 圖表 ----
# 1 資產配置圓餅
pie = PieChart(); pie.title = "資產配置（各帳戶台幣總值）"; pie.height = 7.5; pie.width = 11
data = Reference(a, min_col=6, min_row=4, max_row=A_R2)
cats = Reference(a, min_col=1, min_row=A_R1, max_row=A_R2)
pie.add_data(data, titles_from_data=True); pie.set_categories(cats)
d.add_chart(pie, "A18")
# 2 淨資產趨勢折線
line = LineChart(); line.title = "淨資產趨勢"; line.height = 7.5; line.width = 11
line.y_axis.numFmt = NT
ld = Reference(sn, min_col=2, min_row=4, max_row=S_R2)
lc = Reference(sn, min_col=1, min_row=S_R1, max_row=S_R2)
line.add_data(ld, titles_from_data=True); line.set_categories(lc)
d.add_chart(line, "H18")
# 3 個股報酬率長條
bar = BarChart(); bar.title = "個股報酬率%"; bar.height = 7.5; bar.width = 11; bar.type = "col"
bar.y_axis.numFmt = PCT
bd = Reference(h, min_col=12, min_row=4, max_row=H_R2)
bc = Reference(h, min_col=2, min_row=H_R1, max_row=H_R2)
bar.add_data(bd, titles_from_data=True); bar.set_categories(bc)
d.add_chart(bar, "A34")
# 4 現金流長條
bar2 = BarChart(); bar2.title = "每月淨現金流"; bar2.height = 7.5; bar2.width = 11; bar2.type = "col"
bar2.y_axis.numFmt = NT
b2d = Reference(cf, min_col=6, min_row=4, max_row=C_R2)
b2c = Reference(cf, min_col=1, min_row=C_R1, max_row=C_R2)
bar2.add_data(b2d, titles_from_data=True); bar2.set_categories(b2c)
d.add_chart(bar2, "H34")

wb.save(OUT)
print("✅ 已產生:", OUT)
print("分頁:", wb.sheetnames)
