#!/usr/bin/env python3
"""《工艺参数DOE实验设计工具包.xlsx》：5 Sheet，全部关键数字用公式联动。
示例案例为封装键合三因子：因子水平与 8 个角点数据见 Sheet1/Sheet2，效应/系数
全部由角点数据公式现算。本表示例全部数值为构造的演示数据，不是实测；中心点为 42.3/41.8/42.1 g
与 188/196/192 ppm，正式使用请替换为真实数据。
"""
import itertools
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUT = "/home/hatch/workspace/your_files/doe-tool/工艺参数DOE实验设计工具包.xlsx"

TITLE = Font(bold=True, size=14, color="0D2B45")
H = Font(bold=True, color="FFFFFF")
HFILL = PatternFill("solid", fgColor="0B6BCB")
SUBFILL = PatternFill("solid", fgColor="E8F2FC")
INPUTFILL = PatternFill("solid", fgColor="FFFDF0")
NOTE = Font(size=10, color="5B6B7C")
thin = Side(style="thin", color="DDE4EC")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

def header_row(ws, row, values, widths=None):
    for j, v in enumerate(values, start=1):
        c = ws.cell(row=row, column=j, value=v)
        c.font = H; c.fill = HFILL; c.border = BORDER
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    if widths:
        for j, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(j)].width = w

def put(ws, row, col, value, *, bold=False, fill=None, fmt=None, border=True, wrap=False):
    c = ws.cell(row=row, column=col, value=value)
    if bold: c.font = Font(bold=True)
    if fill: c.fill = fill
    if fmt: c.number_format = fmt
    if border: c.border = BORDER
    if wrap: c.alignment = Alignment(wrap_text=True, vertical="top")
    return c

wb = Workbook()

# ---------------- Sheet1 因子与水平 ----------------
ws = wb.active
ws.title = "1-因子与水平"
ws["A1"] = "因子与水平（案例：封装键合参数优化）"; ws["A1"].font = TITLE
ws["A2"] = "用法：先在本表定因子与高低水平，后面各 Sheet 的实际值、效应、预测全部公式联动。黄色格为输入。"
ws["A2"].font = NOTE
header_row(ws, 4, ["因子", "符号", "低水平 (−1)", "高水平 (+1)", "中心 (0) = (低+高)/2", "单位", "备注"],
           widths=[16, 8, 14, 14, 20, 8, 42])
factors = [
    ("超声功率", "A", 55, 75, "mW", "设备能力与历史窗口内取值"),
    ("键合压力", "B", 35, 55, "cN", ""),
    ("键合时间", "C", 12, 22, "ms", ""),
]
for i, (name, sym, lo, hi, unit, remark) in enumerate(factors):
    r = 5 + i
    put(ws, r, 1, name); put(ws, r, 2, sym)
    put(ws, r, 3, lo, fill=INPUTFILL); put(ws, r, 4, hi, fill=INPUTFILL)
    put(ws, r, 5, f"=(C{r}+D{r})/2", fmt="0.0")
    put(ws, r, 6, unit); put(ws, r, 7, remark)
put(ws, 9, 1, "响应指标", bold=True)
header_row(ws, 10, ["响应", "名称", "目标", "规格/说明", "", "", ""])
for i, (sym, name, goal, spec) in enumerate([
        ("Y1", "球剪切力", "越大越好", "规格 ≥35 g，单位 g"),
        ("Y2", "颈部开裂", "越小越好", "单位 ppm，无硬规格，越低越好")]):
    r = 11 + i
    put(ws, r, 1, sym); put(ws, r, 2, name); put(ws, r, 3, goal); put(ws, r, 4, spec)
put(ws, 14, 1, "检查清单（分析前逐项勾）：因子和水平已定 / 响应指标已明确 / 运行序已随机 / 中心点已安排（本表 3 个）/ 每次实验的设备与环境记录已备好。漏掉随机化，后面所有效应都可能被时间趋势带偏。", wrap=True)
ws.merge_cells("A14:G16")

# ---------------- Sheet2 实验方案与结果 ----------------
ws = wb.create_sheet("2-实验方案与结果")
ws["A1"] = "2³ 全因子 + 3 中心点实验方案与结果（11 次）"; ws["A1"].font = TITLE
ws["A2"] = "标准序只用于记账，实做必须按「运行序」列的随机顺序执行。实际值列由 Sheet1 公式换算；黄色格填实测响应。本表数值为构造的演示数据，不是实测；换成你的真实实验数据即可，公式会自动重算。"
ws["A2"].font = NOTE
header_row(ws, 4, ["标准序", "运行序(随机)", "A 编码", "B 编码", "C 编码",
                   "A 实际", "B 实际", "C 实际",
                   "Y1 球剪切力 (g)", "Y2 颈部开裂 (ppm)",
                   "AB", "AC", "BC", "类型"],
           widths=[9, 13, 9, 9, 9, 10, 10, 10, 16, 17, 7, 7, 7, 9])
corners = [(-1,-1,-1),(1,-1,-1),(-1,1,-1),(1,1,-1),(-1,-1,1),(1,-1,1),(-1,1,1),(1,1,1)]
corner_y = [(37.6,138),(41.3,238),(37.9,98),(47.4,297),(40.0,177),(42.1,309),(40.8,93),(48.7,330)]
center_y = [(42.3,188),(41.8,196),(42.1,192)]
run_order = [7, 3, 11, 1, 9, 5, 2, 10, 4, 8, 6]  # 标准序 -> 随机运行序位置的一组示例排列
rows_data = [(c, y, "角点") for c, y in zip(corners, corner_y)] + [((0,0,0), y, "中心点") for y in center_y]
for i, ((a,b,c), (y1,y2), typ) in enumerate(rows_data):
    r = 5 + i
    put(ws, r, 1, i+1)
    put(ws, r, 2, run_order[i])
    put(ws, r, 3, a); put(ws, r, 4, b); put(ws, r, 5, c)
    for j, col_letter in ((6, "C"), (7, "D"), (8, "E")):
        frow = 5 + (j - 6)
        put(ws, r, j, f"='1-因子与水平'!$E${frow}+{col_letter}{r}*('1-因子与水平'!$D${frow}-'1-因子与水平'!$C${frow})/2", fmt="0.0")
    put(ws, r, 9, y1, fill=INPUTFILL, fmt="0.0")
    put(ws, r, 10, y2, fill=INPUTFILL, fmt="0")
    put(ws, r, 11, f"=C{r}*D{r}"); put(ws, r, 12, f"=C{r}*E{r}"); put(ws, r, 13, f"=D{r}*E{r}")
    put(ws, r, 14, typ)
S2 = "'2-实验方案与结果'"
COR = "5:12"  # 角点行 5..12，中心点行 13..15

# ---------------- Sheet3 效应分析 ----------------
ws = wb.create_sheet("3-效应分析")
ws["A1"] = "效应分析（仅用 8 个角点计算；效应 = +1 均值 − −1 均值；系数 = 效应/2）"; ws["A1"].font = TITLE
ws["A2"] = "全部数字为公式现算。判读：|效应| 大于约 2× 中心点纯误差 SD 才值得认真对待（工程快速判据）；正式结论看 ANOVA 的 P 值。"
ws["A2"].font = NOTE
header_row(ws, 4, ["项", "Y1 效应", "Y1 系数", "Y2 效应", "Y2 系数", "Y1 |效应| 条形（相对）", "含义"],
           widths=[10, 12, 12, 12, 12, 30, 46])
terms = [("A", "C", "超声功率主效应"), ("B", "D", "键合压力主效应"), ("C", "E", "键合时间主效应"),
         ("A×B", "K", "功率×压力交互"), ("A×C", "L", "功率×时间交互"), ("B×C", "M", "压力×时间交互")]
for i, (name, col, meaning) in enumerate(terms):
    r = 5 + i
    put(ws, r, 1, name, bold=True)
    # 效应公式：AVERAGEIF(项列,1,Y列)-AVERAGEIF(项列,-1,Y列)，角点行 5:12
    put(ws, r, 2, f"=AVERAGEIF({S2}!${col}$5:${col}$12,1,{S2}!$I$5:$I$12)-AVERAGEIF({S2}!${col}$5:${col}$12,-1,{S2}!$I$5:$I$12)", fmt="0.00")
    put(ws, r, 3, f"=B{r}/2", fmt="0.00")
    put(ws, r, 4, f"=AVERAGEIF({S2}!${col}$5:${col}$12,1,{S2}!$J$5:$J$12)-AVERAGEIF({S2}!${col}$5:${col}$12,-1,{S2}!$J$5:$J$12)", fmt="0.00")
    put(ws, r, 5, f"=D{r}/2", fmt="0.00")
    put(ws, r, 6, f'=REPT("█",ROUND(ABS(B{r})/MAX($B$5:$B$10)*24,0))')
    put(ws, r, 7, meaning)
put(ws, 11, 1, "截距（角点均值）", bold=True)
put(ws, 11, 2, f"=AVERAGE({S2}!I5:I12)", fmt="0.00")
put(ws, 11, 4, f"=AVERAGE({S2}!J5:J12)", fmt="0.00")
put(ws, 12, 1, "中心点均值", bold=True)
put(ws, 12, 2, f"=AVERAGE({S2}!I13:I15)", fmt="0.00")
put(ws, 12, 4, f"=AVERAGE({S2}!J13:J15)", fmt="0.00")
put(ws, 13, 1, "中心点纯误差 SD", bold=True)
put(ws, 13, 2, f"=STDEV({S2}!I13:I15)", fmt="0.00")
put(ws, 13, 4, f"=STDEV({S2}!J13:J15)", fmt="0.00")
put(ws, 14, 1, "曲率 = 中心均值 − 截距", bold=True)
put(ws, 14, 2, f"=B12-B11", fmt="0.00")
put(ws, 14, 4, f"=D12-D11", fmt="0.00")
put(ws, 16, 1, "回归方程（公式自动拼装，编码值 −1/0/+1）：", bold=True)
put(ws, 17, 1, f'="Y1 球剪切力 = "&TEXT(B11,"0.00")&" "&TEXT(C5,"+0.00;-0.00")&"·A "&TEXT(C6,"+0.00;-0.00")&"·B "&TEXT(C7,"+0.00;-0.00")&"·C "&TEXT(C8,"+0.00;-0.00")&"·AB "&TEXT(C9,"+0.00;-0.00")&"·AC "&TEXT(C10,"+0.00;-0.00")&"·BC"', wrap=True)
ws.merge_cells("A17:G17")
put(ws, 18, 1, f'="Y2 颈部开裂 = "&TEXT(D11,"0.00")&" "&TEXT(E5,"+0.00;-0.00")&"·A "&TEXT(E6,"+0.00;-0.00")&"·B "&TEXT(E7,"+0.00;-0.00")&"·C "&TEXT(E8,"+0.00;-0.00")&"·AB "&TEXT(E9,"+0.00;-0.00")&"·AC "&TEXT(E10,"+0.00;-0.00")&"·BC"', wrap=True)
ws.merge_cells("A18:G18")
put(ws, 20, 1, "判读要点：Y1 的 A 最大、B 次之、A×B 第三且大于 C 的主效应，不可忽略；Y2 的第一驱动是 A 本身（+167 ppm，远大于 A×B 的 +51），只压交互不压功率，开裂下不来。交互图必须看：单看主效应图分不清「功率高开裂高」和「高功率配高压力才爆」差多少。", wrap=True)
ws.merge_cells("A20:G22")

# ---------------- Sheet4 响应优化对照 ----------------
ws = wb.create_sheet("4-响应优化对照")
ws["A1"] = "响应优化对照与网格搜索（预测值全部由 Sheet3 系数公式计算）"; ws["A1"].font = TITLE
ws["A2"] = "上表：三个对照方案（实测值引自 Sheet2 对应角点/中心点）。下表：{−1,0,+1}³ 共 27 个候选点的双响应预测与等权综合满意度（Y1 越大越好、Y2 越小越好，各自按预测范围归一化）。优化完必须做 3–5 批确认实验。"
ws["A2"].font = NOTE
header_row(ws, 4, ["方案", "A", "B", "C", "预测 Y1 (g)", "预测 Y2 (ppm)", "实测 Y1 (g)", "实测 Y2 (ppm)", "说明"],
           widths=[18, 7, 7, 7, 13, 14, 12, 13, 40])
S3 = "'3-效应分析'"
def pred_formula(row, coef_col, inter_col):
    # 预测 = 截距 + Σ 系数*项值；截距在 Sheet3 的 B11(Y1)/D11(Y2)，系数在 C 列(Y1)/E 列(Y2)
    # 项值引用本表 B/C/D 列编码
    return (f"={S3}!${inter_col}$11"
            f"+{S3}!${coef_col}$5*$B{row}+{S3}!${coef_col}$6*$C{row}+{S3}!${coef_col}$7*$D{row}"
            f"+{S3}!${coef_col}$8*$B{row}*$C{row}+{S3}!${coef_col}$9*$B{row}*$D{row}+{S3}!${coef_col}$10*$C{row}*$D{row}")
schemes = [("基线（中心点）", 0, 0, 0, "重复误差与曲率基准"),
           ("推荐（A+B+C−）", 1, 1, -1, "推荐方案：短时间换开裂余量"),
           ("OFAT 全高（A+B+C+）", 1, 1, 1, "只盯剪切力的下场：开裂最高")]
for i, (name, a, b, c, note) in enumerate(schemes):
    r = 5 + i
    put(ws, r, 1, name, bold=True)
    put(ws, r, 2, a); put(ws, r, 3, b); put(ws, r, 4, c)
    put(ws, r, 5, pred_formula(r, "C", "B"), fmt="0.0")
    put(ws, r, 6, pred_formula(r, "E", "D"), fmt="0.0")
    # 实测：角点在 Sheet2 的行序 = 标准序；(+,+,-) 为第 4 个角点行 8，(+,+,+) 为第 8 个角点行 12，中心取 Sheet3 中心均值
    if name.startswith("基线"):
        put(ws, r, 7, f"={S3}!B12", fmt="0.0"); put(ws, r, 8, f"={S3}!D12", fmt="0.0")
    elif name.startswith("推荐"):
        put(ws, r, 7, f"={S2}!I8", fmt="0.0"); put(ws, r, 8, f"={S2}!J8", fmt="0")
    else:
        put(ws, r, 7, f"={S2}!I12", fmt="0.0"); put(ws, r, 8, f"={S2}!J12", fmt="0")
    put(ws, r, 9, note)
put(ws, 9, 1, "网格搜索（27 个候选点，等权满意度）", bold=True)
header_row(ws, 10, ["A", "B", "C", "预测 Y1 (g)", "预测 Y2 (ppm)", "综合满意度", "排名", "", "", ""])
grid_rows = list(itertools.product([-1, 0, 1], repeat=3))
for i, (a, b, c) in enumerate(grid_rows):
    r = 11 + i
    put(ws, r, 1, a); put(ws, r, 2, b); put(ws, r, 3, c)
    # 注意：本表网格区编码在 A/B/C 列，与上面对照区（B/C/D 列）不同，单独写公式
    put(ws, r, 4, f"={S3}!$B$11+{S3}!$C$5*$A{r}+{S3}!$C$6*$B{r}+{S3}!$C$7*$C{r}+{S3}!$C$8*$A{r}*$B{r}+{S3}!$C$9*$A{r}*$C{r}+{S3}!$C$10*$B{r}*$C{r}", fmt="0.0")
    put(ws, r, 5, f"={S3}!$D$11+{S3}!$E$5*$A{r}+{S3}!$E$6*$B{r}+{S3}!$E$7*$C{r}+{S3}!$E$8*$A{r}*$B{r}+{S3}!$E$9*$A{r}*$C{r}+{S3}!$E$10*$B{r}*$C{r}", fmt="0.0")
    put(ws, r, 6, f"=((D{r}-MIN($D$11:$D$37))/(MAX($D$11:$D$37)-MIN($D$11:$D$37)) + (1-(E{r}-MIN($E$11:$E$37))/(MAX($E$11:$E$37)-MIN($E$11:$E$37))))/2", fmt="0.000")
    put(ws, r, 7, f"=RANK(F{r},$F$11:$F$37)")
put(ws, 39, 1, "读表：满意度第 1 名即等权口径下的兼顾点。注意它和「推荐（A+B+C−）」不同不奇怪——那个推荐相当于把剪切力权重拉高、并接受约 300 ppm 为上限；目标函数不同，最优点就不同。先定取舍，再看排名。", wrap=True)
ws.merge_cells("A39:J41")

# ---------------- Sheet5 操作对照与核对说明 ----------------
ws = wb.create_sheet("5-操作对照与说明")
ws["A1"] = "Minitab 与 Excel 操作对照 + 数据说明"; ws["A1"].font = TITLE
header_row(ws, 3, ["步骤", "Minitab 菜单路径", "Excel / 本表做法"],
           widths=[14, 52, 64])
steps = [
    ("创建设计", "Stat → DOE → Factorial → Create Factorial Design（2-Level，3 因子，加 3 中心点）",
     "Sheet2：标准序 8 角点 + 3 中心点，运行序随机；换因子/水平时改 Sheet1，实际值自动换算"),
    ("分析效应", "Stat → DOE → Factorial → Analyze Factorial Design（两个响应分别分析，看 Pareto 与交互图）",
     "Sheet3：AVERAGEIF 按编码/乘积列分组求效应，系数 = 效应/2；或用「数据 → 数据分析 → 回归」，Y 选响应、X 选 A/B/C/AB/AC/BC 六列，结果与本表一致"),
    ("响应优化", "Stat → DOE → Factorial → Response Optimizer（设两个响应的目标/界限）",
     "Sheet4：27 点网格预测 + 满意度排序；权重/界限不同，最优点不同，先说清取舍"),
    ("确认实验", "优化点转成实际参数后跑 3–5 批确认实验，实测落在预测区间内才进量产",
     "同左：把确认批数据追加到 Sheet2 响应列，效应与预测会自动更新，可直接比对"),
]
for i, (s, m, e) in enumerate(steps):
    r = 4 + i
    put(ws, r, 1, s, bold=True); put(ws, r, 2, m, wrap=True); put(ws, r, 3, e, wrap=True)
    ws.row_dimensions[r].height = 46
put(ws, 9, 1, "数据说明（本表示例数据的口径）", bold=True)
notes = [
    "1) 效应值：本表全部效应由 Sheet2 的 8 个角点数据公式现算（Sheet3 结果）：Y1 为 A=5.80、B=3.45、C=1.85、A×B=2.90、A×C=−0.80、B×C=0.25；Y2 为 A=+167、B=−11、C=+34.5、A×B=+51、A×C=+17.5、B×C=−20.5（ppm）。引用时以这组实算值为准。",
    "2) 回归方程：由 Sheet3 系数自动拼装，系数 = 效应/2。注意 Y1 的 A×C 系数为 −0.40（负号），B×C 仅 +0.13，接近零——写方程时不要把交互项的符号和量级凭经验凑整。",
    "3) 开裂驱动：Y2 的第一驱动是 A 主效应（+167 ppm），A×B（+51）次之。只看交互图容易高估交互、低估功率本身；结论方向是高功率配高压力最差，功率本身也必须单独管控。",
    "4) 分辨率：全因子设计无混淆，分辨率是部分因子设计的概念；用部分因子设计查交互时，需选分辨率 ≥V 的设计，主效应才不与二因子交互混淆。",
    "5) 中心点：本表中心点数值为演示用示意重复值，仅用于演示纯误差/曲率计算，正式使用请替换为真实数据。",
]
for i, t in enumerate(notes):
    r = 10 + i * 2
    put(ws, r, 1, t, wrap=True, border=False)
    ws.merge_cells(start_row=r, start_column=1, end_row=r + 1, end_column=3)
    ws.row_dimensions[r].height = 28

wb.save(OUT)
print("saved", OUT)
