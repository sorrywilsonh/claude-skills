#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""產生「my錢錢」資產追蹤 .xlsx：上傳 Google Drive 後用 Google Sheets 開啟即可。
分頁：儀表板 / 設定 / 交易明細 / 持股明細 / 帳戶總覽 / 資產快照 / 現金流。

子帳戶設計：在『設定』把每個子帳戶定義一次（帳戶／類型／子分類），
系統自動產生『幣別』與『識別碼』。其餘分頁只需選一個『識別碼』下拉，幣別自動帶入。
標色：✍️黃底=需手動填寫；🔒灰底=自動計算（勿改）。
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import PieChart, LineChart, BarChart, Reference
from openpyxl.utils import get_column_letter

OUT = "/Users/ucpc/Google Drive/我的雲端硬碟/my錢錢.xlsx"

# ---- 樣式 ----
TITLE   = Font(name="Arial", size=16, bold=True, color="1F3864")
HEAD_M  = Font(name="Arial", size=10, bold=True, color="7F6000")
HEAD_AW = Font(name="Arial", size=10, bold=True, color="FFFFFF")
NOTE    = Font(name="Arial", size=9, italic=True, color="808080")
BOLDB   = Font(name="Arial", size=11, bold=True, color="1F3864")

FILL_MANHEAD = PatternFill("solid", fgColor="FFE599")
FILL_AUTHEAD = PatternFill("solid", fgColor="808080")
FILL_MAN     = PatternFill("solid", fgColor="FFF9E6")
FILL_AUT     = PatternFill("solid", fgColor="F0F0F0")
FILL_KPI     = PatternFill("solid", fgColor="DDEBF7")
FILL_LEGEND  = PatternFill("solid", fgColor="FCE4D6")

THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CTR = Alignment(horizontal="center", vertical="center")

wb = openpyxl.Workbook()

def legend(ws, row=2):
    ws.cell(row=row, column=1, value="✍️ 黃底 = 需手動填寫／更新").fill = FILL_MAN
    ws.cell(row=row, column=1).font = NOTE
    ws.cell(row=row, column=3, value="🔒 灰底 = 自動計算（請勿手動修改）").fill = FILL_AUT
    ws.cell(row=row, column=3).font = NOTE

def headers(ws, row, cols):
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

NT = '#,##0'
NT2 = '#,##0.00'
PCT = '0.0%'

# ---- 跨表參照常數 ----
RATE_RANGE = "'設定'!$A$5:$B$9"            # 幣別→台幣匯率
ACCT_KEY   = "'設定'!$H$5:$H$54"           # 識別碼清單（下拉用）
KEYCOL     = "'設定'!$H$5:$H$54"           # 識別碼欄
TYPECOL    = "'設定'!$E$5:$E$54"           # 類型欄
CURCOL     = "'設定'!$G$5:$G$54"           # 幣別欄
SUBMAP     = "'設定'!$J$5:$K$14"           # 子分類→幣別 對應表
SUBCAT_LIST= "'設定'!$J$5:$J$14"
TYPE_LIST  = "'設定'!$M$5:$M$9"
CUR_LIST   = "'設定'!$N$5:$N$7"
TXN_LIST   = "'設定'!$P$5:$P$11"

def look_type(keycell): return f"IFERROR(INDEX({TYPECOL},MATCH({keycell},{KEYCOL},0)),\"\")"
def look_cur(keycell):  return f"IFERROR(INDEX({CURCOL},MATCH({keycell},{KEYCOL},0)),\"\")"

# ============================================================ 設定
s = wb.active; s.title = "設定"
s["A1"] = "⚙️ 設定（請先在這裡建立『子帳戶清單』）"; s["A1"].font = TITLE
legend(s)
# 匯率表
s["A4"] = "🔒 幣別"; s["B4"] = "🔒 對台幣匯率（自動，需網路）"
for c in ("A4", "B4"): s[c].fill = FILL_AUTHEAD; s[c].font = HEAD_AW; s[c].alignment = CTR
for i, (cur, val) in enumerate([("TWD", 1), ("USD", '=GOOGLEFINANCE("CURRENCY:USDTWD")'),
                                ("JPY", '=GOOGLEFINANCE("CURRENCY:JPYTWD")')], start=5):
    s.cell(row=i, column=1, value=cur).fill = FILL_AUT
    b = s.cell(row=i, column=2, value=val); b.fill = FILL_AUT; b.number_format = NT2
for i in (8, 9):
    s.cell(row=i, column=1).fill = FILL_AUT; s.cell(row=i, column=2).fill = FILL_AUT
s["A11"] = "💡 子帳戶：每個『帳戶＋子分類』填一列，幣別與識別碼會自動產生。"; s["A11"].font = NOTE
s["A12"] = "💡 銀行的子分類用 台幣/美金/日圓；證券用 台股/複委託/美股。"; s["A12"].font = NOTE

# 子帳戶清單（手動填 帳戶/類型/子分類；幣別、識別碼自動）
acc_hdr = [("D", "帳戶名稱", "M"), ("E", "類型", "M"), ("F", "子分類", "M"),
           ("G", "幣別(自動)", "A"), ("H", "識別碼(自動)", "A")]
for col, title, kind in acc_hdr:
    c = s[f"{col}4"]
    c.value = ("✍️ " if kind == "M" else "🔒 ") + title
    c.fill = FILL_MANHEAD if kind == "M" else FILL_AUTHEAD
    c.font = HEAD_M if kind == "M" else HEAD_AW
    c.alignment = CTR
AC_R1, AC_R2 = 5, 54
for r in range(AC_R1, AC_R2 + 1):
    s.cell(row=r, column=7, value=f"=IFERROR(VLOOKUP($F{r},{SUBMAP},2,FALSE),\"\")")        # 幣別自動
    s.cell(row=r, column=8, value=f"=IF($D{r}=\"\",\"\",$D{r}&\"-\"&$F{r})")                # 識別碼自動
for col in (4, 5, 6): fill_col(s, col, AC_R1, AC_R2, "M")
for col in (7, 8): fill_col(s, col, AC_R1, AC_R2, "A")
for w, col in [(14, "D"), (10, "E"), (10, "F"), (9, "G"), (18, "H")]:
    s.column_dimensions[col].width = w
example_accts = [("富邦銀行", "銀行", "台幣"), ("富邦銀行", "銀行", "美金"),
                 ("國泰證券", "證券", "台股"), ("國泰證券", "證券", "複委託"),
                 ("Firstrade", "證券", "美股"), ("基富通", "基金", "台幣")]
for i, (n, t, sub) in enumerate(example_accts, start=AC_R1):
    s.cell(row=i, column=4, value=n); s.cell(row=i, column=5, value=t); s.cell(row=i, column=6, value=sub)

# 子分類→幣別 對應表（可自行增減）
s["J4"] = "✍️ 子分類"; s["K4"] = "✍️ 對應幣別"
for c in ("J4", "K4"): s[c].fill = FILL_MANHEAD; s[c].font = HEAD_M; s[c].alignment = CTR
submap = [("台幣", "TWD"), ("美金", "USD"), ("日圓", "JPY"),
          ("台股", "TWD"), ("美股", "USD"), ("複委託", "USD")]
for i, (sub, cur) in enumerate(submap, start=5):
    s.cell(row=i, column=10, value=sub); s.cell(row=i, column=11, value=cur)
fill_col(s, 10, 5, 14, "M"); fill_col(s, 11, 5, 14, "M")
s.column_dimensions["J"].width = 11; s.column_dimensions["K"].width = 11

# 其他清單
for col, title, items in [("M", "類型清單", ["銀行", "證券", "基金", "保險", "其他"]),
                          ("N", "幣別清單", ["TWD", "USD", "JPY"]),
                          ("P", "交易類型", ["買", "賣", "股息", "入金", "出金", "手續費", "稅"])]:
    s[f"{col}4"] = "🔒 " + title; s[f"{col}4"].fill = FILL_AUTHEAD; s[f"{col}4"].font = HEAD_AW
    for i, it in enumerate(items, start=5): s[f"{col}{i}"] = it
    s.column_dimensions[col].width = 11

s.add_data_validation((dv := DataValidation(type="list", formula1=TYPE_LIST, allow_blank=True))); dv.add(f"E{AC_R1}:E{AC_R2}")
s.add_data_validation((dv := DataValidation(type="list", formula1=SUBCAT_LIST, allow_blank=True))); dv.add(f"F{AC_R1}:F{AC_R2}")
s.add_data_validation((dv := DataValidation(type="list", formula1=CUR_LIST, allow_blank=True))); dv.add("K5:K14")

# ============================================================ 交易明細
t = wb.create_sheet("交易明細")
t["A1"] = "🧾 交易明細（核心流水帳：每筆買/賣/股息/入金/出金記一列）"; t["A1"].font = TITLE
legend(t)
t["A3"] = "💡 帳戶選『識別碼』，幣別自動帶入。入金/出金/股息：金額填『價格』、股數填 1。"; t["A3"].font = NOTE
txn_cols = [("日期","M",12),("帳戶(子帳戶)","M",16),("類型","M",10),("代號","M",12),("名稱","M",14),
            ("股數","M",10),("價格","M",10),("幣別(自動)","A",10),("手續費","M",10),("稅","M",9),
            ("金額小計(原幣)","A",13),("台幣金額","A",13),("手續費稅(台幣)","A",14),
            ("年月","A",9),("備註","M",20)]
headers(t, 4, txn_cols)
T_R1, T_R2 = 5, 504
for r in range(T_R1, T_R2 + 1):
    t.cell(row=r, column=8,  value=f"=IF($B{r}=\"\",\"\",{look_cur(f'$B{r}')})")              # H 幣別自動
    t.cell(row=r, column=11, value=f"=IF($F{r}=\"\",\"\",$F{r}*$G{r})")                        # K 金額小計原幣
    t.cell(row=r, column=12, value=f"=IF($K{r}=\"\",\"\",$K{r}*VLOOKUP($H{r},{RATE_RANGE},2,FALSE))")  # L 台幣金額
    t.cell(row=r, column=13, value=f"=IF(AND($I{r}=\"\",$J{r}=\"\"),\"\",(N($I{r})+N($J{r}))*VLOOKUP($H{r},{RATE_RANGE},2,FALSE))")  # M 手續費稅台幣
    t.cell(row=r, column=14, value=f"=IF($A{r}=\"\",\"\",TEXT($A{r},\"yyyy-mm\"))")            # N 年月
for col in (1,2,3,4,5,6,7,9,10,15): fill_col(t, col, T_R1, T_R2, "M")
for col in (8,11,12,13,14): fill_col(t, col, T_R1, T_R2, "A")
for col,fmt in [(1,'yyyy-mm-dd'),(6,NT2),(7,NT2),(9,NT2),(10,NT2),(11,NT2),(12,NT),(13,NT)]:
    for r in range(T_R1, T_R2+1): t.cell(row=r, column=col).number_format = fmt
# 範例（幣別欄 8 不填，自動）
def putrow(ws, r, vals):
    for col, v in vals.items(): ws.cell(row=r, column=col, value=v)
putrow(t, 5, {1:"2026-01-05",2:"國泰證券-台股",3:"買",4:"TPE:2330",5:"台積電",6:10,7:600,9:20,10:0})
putrow(t, 6, {1:"2026-02-10",2:"Firstrade-美股",3:"買",4:"VOO",5:"Vanguard S&P500",6:5,7:480,9:0,10:0})
putrow(t, 7, {1:"2026-03-01",2:"富邦銀行-台幣",3:"入金",6:1,7:50000,15:"薪轉"})
t.add_data_validation((dv := DataValidation(type="list", formula1=ACCT_KEY, allow_blank=True))); dv.add(f"B{T_R1}:B{T_R2}")
t.add_data_validation((dv := DataValidation(type="list", formula1=TXN_LIST, allow_blank=True))); dv.add(f"C{T_R1}:C{T_R2}")
t.freeze_panes = "A5"

# ============================================================ 持股明細
h = wb.create_sheet("持股明細")
h["A1"] = "📈 持股明細（股數、平均成本由交易明細自動加權計算）"; h["A1"].font = TITLE
legend(h)
h["A3"] = "💡 新增持股只需填：帳戶(子帳戶) / 代號 / 名稱。幣別與其餘全自動。"; h["A3"].font = NOTE
hold_cols = [("帳戶(子帳戶)","M",16),("代號","M",12),("名稱","M",16),("幣別(自動)","A",10),
             ("股數","A",10),("平均成本","A",11),("即時股價","A",11),("市值(原幣)","A",13),
             ("台幣市值","A",13),("投入成本(原幣)","A",14),("未實現損益(原幣)","A",15),
             ("報酬率%","A",10),("台幣投入成本","A",14),("台幣未實現損益","A",15)]
headers(h, 4, hold_cols)
H_R1, H_R2 = 5, 54
TB = f"'交易明細'!$B${T_R1}:$B${T_R2}"; TD = f"'交易明細'!$D${T_R1}:$D${T_R2}"
TC = f"'交易明細'!$C${T_R1}:$C${T_R2}"; TF = f"'交易明細'!$F${T_R1}:$F${T_R2}"
TK = f"'交易明細'!$K${T_R1}:$K${T_R2}"; TI = f"'交易明細'!$I${T_R1}:$I${T_R2}"
TJ = f"'交易明細'!$J${T_R1}:$J${T_R2}"
for r in range(H_R1, H_R2 + 1):
    A, B = f"$A{r}", f"$B{r}"
    buyshares = f"SUMIFS({TF},{TB},{A},{TD},{B},{TC},\"買\")"
    sellshares = f"SUMIFS({TF},{TB},{A},{TD},{B},{TC},\"賣\")"
    buycost = (f"SUMIFS({TK},{TB},{A},{TD},{B},{TC},\"買\")"
               f"+SUMIFS({TI},{TB},{A},{TD},{B},{TC},\"買\")+SUMIFS({TJ},{TB},{A},{TD},{B},{TC},\"買\")")
    h.cell(row=r, column=4,  value=f"=IF($A{r}=\"\",\"\",{look_cur(f'$A{r}')})")                  # 幣別自動
    h.cell(row=r, column=5,  value=f"=IF($B{r}=\"\",\"\",{buyshares}-{sellshares})")               # 股數
    h.cell(row=r, column=6,  value=f"=IF($E{r}=\"\",\"\",IFERROR(({buycost})/{buyshares},0))")      # 平均成本
    h.cell(row=r, column=7,  value=f"=IF($B{r}=\"\",\"\",IFERROR(GOOGLEFINANCE($B{r}),0))")          # 即時股價
    h.cell(row=r, column=8,  value=f"=IF($E{r}=\"\",\"\",$E{r}*$G{r})")                              # 市值原幣
    h.cell(row=r, column=9,  value=f"=IF($H{r}=\"\",\"\",$H{r}*VLOOKUP($D{r},{RATE_RANGE},2,FALSE))")# 台幣市值
    h.cell(row=r, column=10, value=f"=IF($E{r}=\"\",\"\",$E{r}*$F{r})")                              # 投入成本原幣
    h.cell(row=r, column=11, value=f"=IF($H{r}=\"\",\"\",$H{r}-$J{r})")                              # 未實現損益原幣
    h.cell(row=r, column=12, value=f"=IF($J{r}=\"\",\"\",IFERROR($K{r}/$J{r},0))")                   # 報酬率%
    h.cell(row=r, column=13, value=f"=IF($J{r}=\"\",\"\",$J{r}*VLOOKUP($D{r},{RATE_RANGE},2,FALSE))")# 台幣投入成本
    h.cell(row=r, column=14, value=f"=IF($I{r}=\"\",\"\",$I{r}-$M{r})")                              # 台幣未實現損益
for col in (1,2,3): fill_col(h, col, H_R1, H_R2, "M")
for col in range(4,15): fill_col(h, col, H_R1, H_R2, "A")
for col,fmt in [(5,NT2),(6,NT2),(7,NT2),(8,NT2),(9,NT),(10,NT2),(11,NT2),(12,PCT),(13,NT),(14,NT)]:
    for r in range(H_R1, H_R2+1): h.cell(row=r, column=col).number_format = fmt
putrow(h, 5, {1:"國泰證券-台股",2:"TPE:2330",3:"台積電"})
putrow(h, 6, {1:"Firstrade-美股",2:"VOO",3:"Vanguard S&P500"})
h.add_data_validation((dv := DataValidation(type="list", formula1=ACCT_KEY, allow_blank=True))); dv.add(f"A{H_R1}:A{H_R2}")
h.freeze_panes = "A5"

# ============================================================ 帳戶總覽
a = wb.create_sheet("帳戶總覽")
a["A1"] = "🏦 帳戶總覽（選一個子帳戶即可，類型/幣別自動帶入）"; a["A1"].font = TITLE
legend(a)
a["A3"] = "💡 富邦銀行的台幣與美金是兩個不同子帳戶，各選一個即可。"; a["A3"].font = NOTE
acc_cols = [("帳戶(子帳戶)","M",18),("類型","A",10),("幣別","A",8),("現金餘額(原幣)","M",14),
            ("持股市值(台幣)","A",14),("台幣總值","A",14),("佔比%","A",10)]
headers(a, 4, acc_cols)
A_R1, A_R2 = 5, 24
TOTAL_ROW = 26
for r in range(A_R1, A_R2 + 1):
    a.cell(row=r, column=2, value=f"=IF($A{r}=\"\",\"\",{look_type(f'$A{r}')})")               # 類型自動
    a.cell(row=r, column=3, value=f"=IF($A{r}=\"\",\"\",{look_cur(f'$A{r}')})")                # 幣別自動
    a.cell(row=r, column=5, value=f"=IF($A{r}=\"\",\"\",SUMIFS('持股明細'!$I${H_R1}:$I${H_R2},'持股明細'!$A${H_R1}:$A${H_R2},$A{r}))")  # 持股市值依識別碼
    a.cell(row=r, column=6, value=f"=IF($A{r}=\"\",\"\",N($D{r})*VLOOKUP($C{r},{RATE_RANGE},2,FALSE)+N($E{r}))")
    a.cell(row=r, column=7, value=f"=IFERROR($F{r}/$F${TOTAL_ROW},\"\")")
for col in (1,4): fill_col(a, col, A_R1, A_R2, "M")
for col in (2,3,5,6,7): fill_col(a, col, A_R1, A_R2, "A")
for col,fmt in [(4,NT),(5,NT),(6,NT),(7,PCT)]:
    for r in range(A_R1, A_R2+1): a.cell(row=r, column=col).number_format = fmt
a.cell(row=TOTAL_ROW, column=1, value="總計").font = BOLDB
tot = a.cell(row=TOTAL_ROW, column=6, value=f"=SUM(F{A_R1}:F{A_R2})"); tot.font = BOLDB; tot.number_format = NT; tot.fill = FILL_KPI
for i,(key,bal) in enumerate([("富邦銀行-台幣",50000),("富邦銀行-美金",3000),
                              ("國泰證券-台股",0),("國泰證券-複委託",0),("Firstrade-美股",0)]):
    a.cell(row=A_R1+i, column=1, value=key); a.cell(row=A_R1+i, column=4, value=bal)
a.add_data_validation((dv := DataValidation(type="list", formula1=ACCT_KEY, allow_blank=True))); dv.add(f"A{A_R1}:A{A_R2}")
a.freeze_panes = "A5"

# ============================================================ 資產快照
sn = wb.create_sheet("資產快照")
sn["A1"] = "📅 資產快照（每月把儀表板的總淨資產『貼數值』記一列 → 畫趨勢線）"; sn["A1"].font = TITLE
legend(sn)
sn["A3"] = "💡 請貼『數值』不要貼公式，否則歷史會跟著變動。"; sn["A3"].font = NOTE
headers(sn, 4, [("日期","M",14),("淨資產合計(台幣)","M",18),("備註","M",24)])
S_R1, S_R2 = 5, 64
fill_col(sn, 1, S_R1, S_R2, "M", 'yyyy-mm-dd'); fill_col(sn, 2, S_R1, S_R2, "M", NT); fill_col(sn, 3, S_R1, S_R2, "M")
for i,(d_,v) in enumerate([("2026-01-31",1200000),("2026-02-28",1255000),("2026-03-31",1310000)]):
    sn.cell(row=S_R1+i, column=1, value=d_); sn.cell(row=S_R1+i, column=2, value=v)
sn.freeze_panes = "A5"

# ============================================================ 現金流
cf = wb.create_sheet("現金流")
cf["A1"] = "💰 現金流（金額自動從交易明細彙整成台幣 → 畫長條圖）"; cf["A1"].font = TITLE
legend(cf)
cf["A3"] = "💡 只需手動填『年月』(如 2026-06)，其餘自動。"; cf["A3"].font = NOTE
headers(cf, 4, [("年月","M",10),("入金","A",12),("股息","A",12),("出金/提領","A",12),("手續費+稅","A",12),("淨現金流","A",13)])
C_R1, C_R2 = 5, 40
TL = f"'交易明細'!$N${T_R1}:$N${T_R2}"; TLamt = f"'交易明細'!$L${T_R1}:$L${T_R2}"; TM = f"'交易明細'!$M${T_R1}:$M${T_R2}"
for r in range(C_R1, C_R2 + 1):
    A = f"$A{r}"
    cf.cell(row=r, column=2, value=f"=IF($A{r}=\"\",\"\",SUMIFS({TLamt},{TL},{A},{TC},\"入金\"))")
    cf.cell(row=r, column=3, value=f"=IF($A{r}=\"\",\"\",SUMIFS({TLamt},{TL},{A},{TC},\"股息\"))")
    cf.cell(row=r, column=4, value=f"=IF($A{r}=\"\",\"\",SUMIFS({TLamt},{TL},{A},{TC},\"出金\"))")
    cf.cell(row=r, column=5, value=f"=IF($A{r}=\"\",\"\",SUMIFS({TM},{TL},{A}))")
    cf.cell(row=r, column=6, value=f"=IF($A{r}=\"\",\"\",N($B{r})+N($C{r})-N($D{r})-N($E{r}))")
fill_col(cf, 1, C_R1, C_R2, "M")
for col in (2,3,4,5,6): fill_col(cf, col, C_R1, C_R2, "A", NT)
for i,m in enumerate(["2026-01","2026-02","2026-03"]): cf.cell(row=C_R1+i, column=1, value=m)
cf.freeze_panes = "A5"

# ============================================================ 儀表板
d = wb.create_sheet("儀表板")
wb.move_sheet("儀表板", -(len(wb.sheetnames)-1))
d["A1"] = "📊 my錢錢 — 資產儀表板"; d["A1"].font = Font(name="Arial", size=20, bold=True, color="1F3864")
d["A2"] = "上傳 Google Drive → 右鍵『用 Google Sheets 開啟』，即時股價/匯率會自動生效。"; d["A2"].font = NOTE
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
d["A9"] = "🔧 需要你『手動更新』的項目（黃底欄位）"; d["A9"].font = BOLDB; d["A9"].fill = FILL_LEGEND
manual_items = [
    "① 設定：建立『子帳戶清單』(帳戶／類型／子分類) — 幣別與識別碼自動，一次性",
    "② 帳戶總覽：選子帳戶 + 填『現金餘額(原幣)』— 不定期",
    "③ 持股明細：新增持股填『帳戶／代號／名稱』(幣別與成本自動)",
    "④ 交易明細：每筆買/賣/股息/入金/出金記一列 — 平時主力",
    "⑤ 資產快照：每月把上方『總淨資產』貼成數值一列",
    "⑥ 現金流：填入『年月』(如 2026-06)，金額自動算",
]
for i, txt in enumerate(manual_items, start=10):
    c = d.cell(row=i, column=1, value=txt); c.font = Font(name="Arial", size=10, color="7F6000"); c.fill = FILL_MAN
    d.merge_cells(start_row=i, start_column=1, end_row=i, end_column=6)
# 圖表
pie = PieChart(); pie.title = "資產配置（各子帳戶台幣總值）"; pie.height = 7.5; pie.width = 11
pie.add_data(Reference(a, min_col=6, min_row=4, max_row=A_R2), titles_from_data=True)
pie.set_categories(Reference(a, min_col=1, min_row=A_R1, max_row=A_R2)); d.add_chart(pie, "A18")
line = LineChart(); line.title = "淨資產趨勢"; line.height = 7.5; line.width = 11; line.y_axis.numFmt = NT
line.add_data(Reference(sn, min_col=2, min_row=4, max_row=S_R2), titles_from_data=True)
line.set_categories(Reference(sn, min_col=1, min_row=S_R1, max_row=S_R2)); d.add_chart(line, "H18")
bar = BarChart(); bar.title = "個股報酬率%"; bar.height = 7.5; bar.width = 11; bar.type = "col"; bar.y_axis.numFmt = PCT
bar.add_data(Reference(h, min_col=12, min_row=4, max_row=H_R2), titles_from_data=True)
bar.set_categories(Reference(h, min_col=2, min_row=H_R1, max_row=H_R2)); d.add_chart(bar, "A34")
bar2 = BarChart(); bar2.title = "每月淨現金流"; bar2.height = 7.5; bar2.width = 11; bar2.type = "col"; bar2.y_axis.numFmt = NT
bar2.add_data(Reference(cf, min_col=6, min_row=4, max_row=C_R2), titles_from_data=True)
bar2.set_categories(Reference(cf, min_col=1, min_row=C_R1, max_row=C_R2)); d.add_chart(bar2, "H34")

wb.save(OUT)
print("✅ 已產生:", OUT)
print("分頁:", wb.sheetnames)
