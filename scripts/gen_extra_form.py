# -*- coding: utf-8 -*-
"""重新生成《表外数据收集表.xlsx / .csv》到 examples/ 目录。
用法：python scripts/gen_extra_form.py
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
OUT_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "examples")

ITEMS = [
    ("", "应收票据贴现利息支出", "票据贴现产生的利息（如无贴现可留空）", "选填"),
    ("", "支付给职工的工资", "全年实发工资总额（含奖金津贴）★直接影响经营净额", "必填"),
    ("", "支付给职工的四金", "社保+公积金等（不填则按工资×26.6%自动推算）", "选填"),
    ("", "支付给职工的其他福利费", "其他职工福利（不填则按工资×1.06%自动推算）", "选填"),
    ("", "销项税额", "全年销项税额（不填则按收入×6%推算，请按实际税率改）", "选填"),
    ("", "进项税额", "全年进项税额（不填则按成本×6%推算）", "选填"),
    ("", "应交增值税", "销项−进项（不填自动计算）", "选填"),
    ("", "其他各项税", "除增值税外各税种（不填取利润表营业税金及附加）", "选填"),
    ("", "所得税", "全年所得税（不填取利润表所得税费用）", "选填"),
    ("", "管理费用中列支的税金", "计入管理费用的税金（如房产税、印花税）", "选填"),
    ("", "在其他业务支出中列支的税金", "其他业务负担的税金", "选填"),
    ("", "分配股利所支付的现金", "全年向股东分配股利支付的现金", "选填"),
    ("", "分配利润所支付的现金", "全年分配利润支付的现金", "选填"),
    ("", "利息支出", "利息支出总额（不抵减利息收入；即利润表财务费用中的利息部分）", "选填"),
    ("", "计提坏账准备的比率", "坏账计提比例（默认0.5%，按公司实际改）", "选填"),
    ("", "计提的坏账准备", "当年计提坏账（不填为0，与模板口径一致）", "选填"),
    ("", "无形资产摊销", "当年无形资产摊销额", "选填"),
    ("", "长期待摊费用摊销", "当年长期待摊费用摊销额", "选填"),
    ("", "固定资产报废损失", "当年报废损失（不计入处置现金净额）", "选填"),
    ("", "收到投资分红或利润", "当年收到的对外投资分红/利润", "选填"),
    ("", "处置固定资产收回的现金净额", "处置固定资产收到的现金净额", "选填"),
    ("", "处置无形资产收回的现金净额", "处置无形资产收到的现金净额", "选填"),
    ("", "处置其他长期资产收回的现金净额", "处置其他长期资产收到的现金净额", "选填"),
    ("", "收到的其他与投资活动有关的现金", "投资活动其他现金流入", "选填"),
    ("", "支付的其他与投资活动有关的现金", "投资活动其他现金流出", "选填"),
    ("", "收到的其他与筹资活动有关的现金", "筹资活动其他现金流入", "选填"),
    ("", "偿还债务所支付的现金", "偿还债务本金支付的现金（不填则按借款变动自动倒挤）", "选填"),
    ("", "支付的其他与筹资活动有关的现金", "筹资活动其他现金流出", "选填"),
    ("", "汇率变动对现金的影响", "外币折算对现金的影响（无外币业务填0）", "选填"),
    ("", "现金等价物期末余额", "如有现金等价物（一般企业可留空）", "选填"),
    ("", "现金等价物期初余额", "如有现金等价物（一般企业可留空）", "选填"),
    ("", "支付的其他与经营活动有关的现金", "经营活动其他现金流出（不填则由平衡项自动倒挤）", "选填"),
]

# ===== 高精度模式（可选）：直接从现金流量表主表抄真实值，精度最高 =====
# 填了这些项后，引擎用真实值替代公式推算，误差可压到 ±1% 以内（前提：各真实值口径一致）
HIGH_PRECISION_ITEMS = [
    ("", "销售商品提供劳务收到的现金", '现金流量表主表"销售商品、提供劳务收到的现金"（高精度模式，选填）', "选填"),
    ("", "收到的税费返还", '现金流量表主表"收到的税费返还"（高精度模式，选填）', "选填"),
    ("", "购买商品接受劳务支付的现金", '现金流量表主表"购买商品、接受劳务支付的现金"（高精度模式，选填）', "选填"),
    ("", "支付给职工以及为职工支付的现金", '现金流量表主表"支付给职工以及为职工支付的现金"（高精度模式，选填）', "选填"),
    ("", "支付的各项税费", '现金流量表主表"支付的各项税费"（高精度模式，选填）', "选填"),
    ("", "收到的其他与经营活动有关的现金", '现金流量表主表"收到的其他与经营活动有关的现金"（高精度模式，选填）★填此项后经营净额不再倒挤，按各项真实值求和', "选填"),
    ("", "支付的其他与经营活动有关的现金", '现金流量表主表"支付的其他与经营活动有关的现金"（高精度模式，选填）', "选填"),
    ("", "购建固定资产无形资产和其他长期资产支付的现金", '现金流量表主表"购建固定资产、无形资产和其他长期资产支付的现金"（高精度模式，选填）', "选填"),
    ("", "投资所支付的现金", '现金流量表主表"投资所支付的现金"（高精度模式，选填）', "选填"),
    ("", "吸收投资所收到的现金", '现金流量表主表"吸收投资收到的现金"（高精度模式，选填）', "选填"),
    ("", "借款所收到的现金", '现金流量表主表"借款收到的现金"（高精度模式，选填）', "选填"),
]

NOTES = [
    "★ 必填项只有：支付给职工的工资。其余均可留空由引擎自动推算/置0。",
    "★ 金额单位：元。直接填数字，不要带千分位逗号。",
    "★ 填完后保存，运行时作为 --extra 参数传入：python cash_flow_generator.py --bs 资产负债表 --pl 利润表 --extra 表外数据收集表.xlsx",
    "★ 补齐本表后，'经营活动产生的现金流量净额'精度可由 ±20% 提升至 ±5% 以内。",
    "★ 数据来源建议：工资表/社保申报表、增值税及所得税申报表、固定资产折旧明细表、银行对账单。",
]


def gen_xlsx(path):
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

    wb = Workbook()
    ws = wb.active
    ws.title = "表外数据收集表"

    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    title_font = Font(name="宋体", size=14, bold=True)
    head_font = Font(name="宋体", size=11, bold=True)
    normal_font = Font(name="宋体", size=11)
    must_fill = PatternFill("solid", fgColor="FFF2CC")
    head_fill = PatternFill("solid", fgColor="BDD7EE")
    note_fill = PatternFill("solid", fgColor="F2F2F2")

    ws.merge_cells("A1:D1")
    ws["A1"] = "表 外 数 据 收 集 表"
    ws["A1"].font = title_font
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:D2")
    ws["A2"] = "编制现金流量表所需表外数据 —— 在白色'金额'栏填写，其余列不要改动"
    ws["A2"].font = normal_font

    hdr = ["序号", "项  目", "金  额", "说明"]
    for ci, h in enumerate(hdr, 1):
        cell = ws.cell(row=3, column=ci, value=h)
        cell.font = head_font
        cell.fill = head_fill
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center")

    r = 4
    for i, (seq, name, note, req) in enumerate(ITEMS, 1):
        ws.cell(row=r, column=1, value=i).border = border
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="center")
        cname = ws.cell(row=r, column=2, value=name)
        cname.border = border
        cname.font = normal_font
        cval = ws.cell(row=r, column=3)
        cval.border = border
        cval.alignment = Alignment(horizontal="right")
        cnote = ws.cell(row=r, column=4, value=note)
        cnote.border = border
        cnote.font = Font(name="宋体", size=9, color="808080")
        cnote.fill = note_fill
        if req == "必填":
            cname.fill = must_fill
            cval.fill = must_fill
        r += 1

    # ===== 高精度模式区域 =====
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    hcell = ws.cell(row=r, column=1,
                    value="【高精度模式·可选】以下从现金流量表主表直接抄真实值，填了即用真实值替代公式推算，误差可压到 ±1% 以内")
    hcell.font = Font(name="宋体", size=10, bold=True, color="006100")
    hcell.fill = PatternFill("solid", fgColor="E2EFDA")
    r += 1
    for ci, h in enumerate(hdr, 1):
        cell = ws.cell(row=r, column=ci, value=h)
        cell.font = head_font
        cell.fill = PatternFill("solid", fgColor="C6EFCE")
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center")
    r += 1
    for i, (seq, name, note, req) in enumerate(HIGH_PRECISION_ITEMS, 1):
        ws.cell(row=r, column=1, value=i).border = border
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="center")
        cname = ws.cell(row=r, column=2, value=name)
        cname.border = border
        cname.font = normal_font
        cval = ws.cell(row=r, column=3)
        cval.border = border
        cval.alignment = Alignment(horizontal="right")
        cnote = ws.cell(row=r, column=4, value=note)
        cnote.border = border
        cnote.font = Font(name="宋体", size=9, color="808080")
        cnote.fill = note_fill
        r += 1

    r += 1
    for note in NOTES:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        cell = ws.cell(row=r, column=1, value="※ " + note)
        cell.font = Font(name="宋体", size=9, color="C00000")
        r += 1

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 52
    wb.save(path)
    print("已生成:", path)


def gen_csv(path):
    import csv
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["序号", "项目", "金额"])
        for i, (_, name, _, _) in enumerate(ITEMS, 1):
            w.writerow([i, name, ""])
        w.writerow([])
        w.writerow(["高精度模式（可选）：从现金流量表主表直接抄真实值，填了即用真实值替代公式推算"])
        for i, (_, name, _, _) in enumerate(HIGH_PRECISION_ITEMS, 1):
            w.writerow([i, name, ""])
    print("已生成:", path)


if __name__ == "__main__":
    os.makedirs(OUT_DIR, exist_ok=True)
    gen_xlsx(os.path.join(OUT_DIR, "表外数据收集表.xlsx"))
    gen_csv(os.path.join(OUT_DIR, "表外数据收集表.csv"))
