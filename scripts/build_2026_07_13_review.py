from __future__ import annotations

import csv
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


ROOT = Path(__file__).resolve().parents[1]
INPUT = ROOT / "data" / "holding-intake.csv"
OUTPUT = ROOT / "reports" / "2026-07-13-基金仓位复述与市场复盘.docx"

BLUE = "2E74B5"
DARK_BLUE = "1F4D78"
INK = "1F2937"
MUTED = "667085"
LIGHT = "F2F4F7"
PALE_BLUE = "E8EEF5"
RED = "9B1C1C"
GOLD = "7A5A00"
GREEN = "256D4A"
WHITE = "FFFFFF"


def set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_margins(cell, top=80, start=120, bottom=80, end=120) -> None:
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for margin, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{margin}"))
        if node is None:
            node = OxmlElement(f"w:{margin}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_repeat_table_header(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def set_table_borders(table, color="D0D5DD", size="4") -> None:
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = borders.find(qn(f"w:{edge}"))
        if tag is None:
            tag = OxmlElement(f"w:{edge}")
            borders.append(tag)
        tag.set(qn("w:val"), "single")
        tag.set(qn("w:sz"), size)
        tag.set(qn("w:color"), color)


def set_table_geometry(table, widths_dxa: list[int], indent_dxa=120) -> None:
    total = sum(widths_dxa)
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    tbl_pr = table._tbl.tblPr
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:w"), str(total))
    tbl_w.set(qn("w:type"), "dxa")
    tbl_ind = tbl_pr.find(qn("w:tblInd"))
    if tbl_ind is None:
        tbl_ind = OxmlElement("w:tblInd")
        tbl_pr.append(tbl_ind)
    tbl_ind.set(qn("w:w"), str(indent_dxa))
    tbl_ind.set(qn("w:type"), "dxa")
    grid = table._tbl.tblGrid
    for child in list(grid):
        grid.remove(child)
    for width in widths_dxa:
        col = OxmlElement("w:gridCol")
        col.set(qn("w:w"), str(width))
        grid.append(col)
    for row in table.rows:
        for idx, cell in enumerate(row.cells):
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_w = tc_pr.find(qn("w:tcW"))
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                tc_pr.append(tc_w)
            tc_w.set(qn("w:w"), str(widths_dxa[idx]))
            tc_w.set(qn("w:type"), "dxa")
            cell.width = Inches(widths_dxa[idx] / 1440)
            set_cell_margins(cell)


def set_font(run, size=11, bold=False, color=INK, italic=False) -> None:
    run.font.name = "Calibri"
    run._element.get_or_add_rPr().rFonts.set(qn("w:ascii"), "Calibri")
    run._element.get_or_add_rPr().rFonts.set(qn("w:hAnsi"), "Calibri")
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "PingFang SC")
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = RGBColor.from_string(color)


def format_paragraph(paragraph, before=0, after=6, line=1.10, align=None) -> None:
    pf = paragraph.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line
    if align is not None:
        paragraph.alignment = align


def add_rich_paragraph(doc, parts, before=0, after=6, line=1.10, align=None):
    p = doc.add_paragraph()
    format_paragraph(p, before, after, line, align)
    for text, options in parts:
        run = p.add_run(text)
        set_font(run, **options)
    return p


def add_heading(doc, text: str, level=1):
    p = doc.add_paragraph(style=f"Heading {level}")
    p.add_run(text)
    return p


def add_callout(doc, label: str, text: str, tone="blue") -> None:
    fills = {"blue": PALE_BLUE, "red": "FCE8E6", "gold": "FFF4CE", "green": "E6F4EA"}
    colors = {"blue": DARK_BLUE, "red": RED, "gold": GOLD, "green": GREEN}
    table = doc.add_table(rows=1, cols=1)
    set_table_geometry(table, [9360], 120)
    set_table_borders(table, fills[tone], "2")
    cell = table.cell(0, 0)
    set_cell_shading(cell, fills[tone])
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    p = cell.paragraphs[0]
    format_paragraph(p, 1, 1, 1.10)
    r = p.add_run(f"{label}  ")
    set_font(r, 11, True, colors[tone])
    r = p.add_run(text)
    set_font(r, 11, False, INK)
    spacer = doc.add_paragraph()
    format_paragraph(spacer, 0, 2, 1.0)


def add_hyperlink(paragraph, text: str, url: str):
    part = paragraph.part
    rid = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), rid)
    run = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), BLUE)
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    rfonts = OxmlElement("w:rFonts")
    rfonts.set(qn("w:ascii"), "Calibri")
    rfonts.set(qn("w:hAnsi"), "Calibri")
    rfonts.set(qn("w:eastAsia"), "PingFang SC")
    size = OxmlElement("w:sz")
    size.set(qn("w:val"), "20")
    r_pr.extend([rfonts, color, underline, size])
    run.append(r_pr)
    text_node = OxmlElement("w:t")
    text_node.text = text
    run.append(text_node)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def add_page_number(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = paragraph.add_run("第 ")
    set_font(r, 9, False, MUTED)
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    r._r.extend([begin, instr, separate, end])
    r2 = paragraph.add_run(" 页")
    set_font(r2, 9, False, MUTED)


def build_document() -> None:
    rows = list(csv.DictReader(INPUT.open(encoding="utf-8")))
    active = [r for r in rows if not r["status"].startswith("sold_")]
    active_total = sum(float(r["market_value"]) for r in active)
    active_cost = sum(float(r["estimated_cost"]) for r in active)
    active_gain = sum(float(r["holding_gain"]) for r in active)
    all_total = sum(float(r["market_value"]) for r in rows)
    qdii_total = sum(float(r["market_value"]) for r in active if "qdii" in r["cluster"])
    tech_total = sum(
        float(r["market_value"])
        for r in active
        if any(x in r["cluster"] for x in ("semiconductor", "tech", "communication", "cloud"))
    )

    doc = Document()
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(1.0)
    section.bottom_margin = Inches(1.0)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)
    section.header_distance = Inches(0.492)
    section.footer_distance = Inches(0.492)

    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = "Calibri"
    normal._element.rPr.rFonts.set(qn("w:ascii"), "Calibri")
    normal._element.rPr.rFonts.set(qn("w:hAnsi"), "Calibri")
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "PingFang SC")
    normal.font.size = Pt(11)
    normal.font.color.rgb = RGBColor.from_string(INK)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.10
    for name, size, color, before, after in (
        ("Heading 1", 16, BLUE, 16, 8),
        ("Heading 2", 13, BLUE, 12, 6),
        ("Heading 3", 12, DARK_BLUE, 8, 4),
    ):
        style = styles[name]
        style.font.name = "Calibri"
        style._element.rPr.rFonts.set(qn("w:ascii"), "Calibri")
        style._element.rPr.rFonts.set(qn("w:hAnsi"), "Calibri")
        style._element.rPr.rFonts.set(qn("w:eastAsia"), "PingFang SC")
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor.from_string(color)
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(after)
        style.paragraph_format.keep_with_next = True

    header = section.header.paragraphs[0]
    format_paragraph(header, 0, 0, 1.0)
    r = header.add_run("基金账户复盘  ·  研究记录")
    set_font(r, 9, True, MUTED)
    footer = section.footer.paragraphs[0]
    add_page_number(footer)

    p = doc.add_paragraph()
    format_paragraph(p, 8, 4, 1.0)
    r = p.add_run("基金仓位复述与市场复盘")
    set_font(r, 24, True, "000000")
    p = doc.add_paragraph()
    format_paragraph(p, 0, 16, 1.0)
    r = p.add_run("2026年7月13日收盘后｜仓位底稿截至7月12日截图口径")
    set_font(r, 12, False, MUTED)

    metadata = [
        ("账户定位", "中国大陆场外公募基金；人工确认交易"),
        ("报告范围", "持仓迁移核对、今日信息面、技术面与风险审计"),
        ("数据时点", "A股/港股截至7月13日收盘；美股仅截至7月10日收盘"),
        ("重要限制", "当日基金正式净值、份额、确认日期、费率与现金余额不完整"),
    ]
    for label, value in metadata:
        add_rich_paragraph(doc, [(f"{label}：", {"size": 10.5, "bold": True, "color": "000000"}), (value, {"size": 10.5, "bold": False, "color": "000000"})], after=2)

    add_heading(doc, "一、先复述你的仓位", 1)
    add_callout(
        doc,
        "迁移结论",
        f"底稿共19条。剔除“创新药已卖出、待确认”后，当前暂按18只在持，市值 {active_total:,.2f} 元，估算成本 {active_cost:,.2f} 元，截图口径浮亏 {abs(active_gain):,.2f} 元（{active_gain/active_cost:.2%}）。若把待确认卖出的1,421.58元也计入，截图总额为 {all_total:,.2f} 元。",
        "blue",
    )
    add_rich_paragraph(doc, [("口头版复述：", {"size": 11, "bold": True, "color": DARK_BLUE}), ("你以沪深300作为最大单一底仓，同时叠加半导体材料设备、电子信息、通信设备、云计算和多只科技主动基金；海外部分由全球成长、全球产业升级、高端制造、纳指100及若干科技QDII组成。仓位数量多、科技因子重叠明显。", {"size": 11, "bold": False, "color": INK})], after=8)

    holdings = doc.add_table(rows=1, cols=6)
    headers = ["代码", "基金", "市值（元）", "占在持", "持有收益", "状态"]
    widths = [950, 3860, 1300, 1050, 1100, 1100]
    for i, text in enumerate(headers):
        cell = holdings.rows[0].cells[i]
        set_cell_shading(cell, LIGHT)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        format_paragraph(p, 0, 0, 1.0)
        run = p.add_run(text)
        set_font(run, 9.2, True, "000000")
    set_repeat_table_header(holdings.rows[0])
    for rdata in rows:
        sold = rdata["status"].startswith("sold_")
        vals = [
            rdata["fund_code"],
            rdata["fund_name"],
            f"{float(rdata['market_value']):,.2f}",
            "—" if sold else f"{float(rdata['market_value'])/active_total:.2%}",
            f"{float(rdata['holding_gain']):+,.2f} / {float(rdata['holding_return']):+.2%}",
            "已卖待确认" if sold else "暂按在持",
        ]
        row = holdings.add_row()
        for i, text in enumerate(vals):
            cell = row.cells[i]
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if i == 1 else WD_ALIGN_PARAGRAPH.CENTER
            format_paragraph(p, 0, 0, 1.0)
            run = p.add_run(text)
            color = MUTED if sold else (RED if i == 4 and float(rdata["holding_gain"]) < 0 else INK)
            set_font(run, 8.7, False, color)
    set_table_geometry(holdings, widths, 120)
    set_table_borders(holdings)

    add_heading(doc, "二、仓位结构与迁移质量", 1)
    structure = doc.add_table(rows=1, cols=4)
    for i, text in enumerate(["检查项", "当前口径", "项目约束", "结论"]):
        set_cell_shading(structure.rows[0].cells[i], PALE_BLUE)
        p = structure.rows[0].cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        format_paragraph(p, 0, 0, 1.0)
        set_font(p.add_run(text), 9.5, True, "000000")
    checks = [
        ("持仓数量", "18只在持 + 1只待确认卖出", "最多3只", "明显不一致"),
        ("最大单只", "沪深300 25.30%", "不高于25%", "轻微超限"),
        ("显性科技主题", f"至少 {tech_total/active_total:.2%}", "主题不高于40%", "已超；实际或更高"),
        ("QDII", f"{qdii_total:,.2f}元 / {qdii_total/active_total:.2%}", "不高于25%", "口径内未超"),
        ("资金规模", f"在持市值 {active_total:,.2f}元", "配置本金30,000元", "配置可能已过期"),
        ("现金与回撤", "现金未知；无组合历史", "现金≥25%；回撤8%降仓", "无法核验"),
    ]
    for values in checks:
        row = structure.add_row()
        for i, text in enumerate(values):
            cell = row.cells[i]
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if i in (0, 1) else WD_ALIGN_PARAGRAPH.CENTER
            format_paragraph(p, 0, 0, 1.0)
            color = RED if i == 3 and text not in ("口径内未超",) else INK
            set_font(p.add_run(text), 9.2, i == 0, color)
    set_table_geometry(structure, [1700, 3300, 1900, 2460], 120)
    set_table_borders(structure)
    add_callout(doc, "迁移缺口", "当前只有截图估值，没有基金份额、申购/确认日期、已确认净值、交易流水、现金余额和适用费率。因而不能准确计算7月13日账户收益、持有期赎回费或账户从高点回撤。", "gold")

    add_heading(doc, "三、今天的信息面复盘", 1)
    add_heading(doc, "1. 外部冲击：地缘风险抬升油价与利率压力", 2)
    add_rich_paragraph(doc, [("事实。", {"size": 11, "bold": True, "color": DARK_BLUE}), ("周末中东冲突再度升级，7月13日全球风险资产承压。AP报道亚洲交易时段韩国KOSPI下跌8.9%；美国10年期国债收益率升至4.59%，高于上周五的4.56%。", {"size": 11, "bold": False, "color": INK})])
    add_rich_paragraph(doc, [("因果链。", {"size": 11, "bold": True, "color": DARK_BLUE}), ("冲突升级 → 原油与通胀风险上升 → 降息预期受压、长端利率抬升 → 高估值成长资产折现率上升。你的半导体、通信、电子信息和全球成长仓位对这一链条更敏感。", {"size": 11, "bold": False, "color": INK})])
    add_rich_paragraph(doc, [("反方解释。", {"size": 11, "bold": True, "color": DARK_BLUE}), ("今天并非单纯由地缘事件造成。前期AI硬件拥挤、估值偏高和中报盈利验证压力，可能才是内部主因；外部冲击更像触发器和放大器。", {"size": 11, "bold": False, "color": INK})])

    add_heading(doc, "2. A股内部：科技拥挤交易进入盈利验证", 2)
    add_rich_paragraph(doc, [("事实。", {"size": 11, "bold": True, "color": DARK_BLUE}), ("上周股票ETF净流入约828亿元，半导体材料设备相关ETF是主要吸金方向；但今天存储、PCB、先进封装、光纤光通信等高位方向集中回撤。7月又进入中报业绩预告密集期，资金从“叙事定价”转向订单、利润和估值匹配。", {"size": 11, "bold": False, "color": INK})])
    add_rich_paragraph(doc, [("市场验证。", {"size": 11, "bold": True, "color": DARK_BLUE}), ("银行、石油天然气、煤炭、燃气等防御与资源方向相对占优，而电子元器件、半导体、电脑硬件和光通信明显承压，符合“高拥挤成长降温、资金向防御再平衡”的解释。", {"size": 11, "bold": False, "color": INK})])
    add_rich_paragraph(doc, [("明日事件。", {"size": 11, "bold": True, "color": DARK_BLUE}), ("国家统计局日程显示7月14日9:30将发布国民经济运行相关数据。若数据或政策预期明显偏离市场预期，可能改变宽基与成长风格的相对表现。", {"size": 11, "bold": False, "color": INK})])

    add_heading(doc, "四、今天的技术面复盘", 1)
    market = doc.add_table(rows=1, cols=4)
    for i, text in enumerate(["观察对象", "7月13日表现", "技术含义", "对应持仓"]):
        set_cell_shading(market.rows[0].cells[i], LIGHT)
        p = market.rows[0].cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        format_paragraph(p, 0, 0, 1.0)
        set_font(p.add_run(text), 9.5, True, "000000")
    observations = [
        ("沪深300", "-1.79%", "跌幅小于成长指数，权重价值相对抗跌", "沪深300联接"),
        ("上证指数", "-2.06%，3913.79", "早盘-0.75%后继续走弱，日内承接不足", "宽基/主动A股"),
        ("深证成指", "-3.48%，14522.85", "成长与科技权重拖累明显", "电子、半导体、通信"),
        ("创业板指", "-3.10%，3723.52", "高弹性成长弱于沪深300约1.31个百分点", "主动成长"),
        ("科创综指", "-4.36%", "高估值硬科技是主要风险释放区", "半导体材料设备"),
        ("市场宽度", "超4600股下跌；逾170股跌停", "不是少数权重拖累，而是广泛风险收缩", "几乎全部A股仓位"),
        ("两市成交", "2.8178万亿元；缩量5708亿元", "跌价缩量说明抛压与接盘同时退潮，尚非放量恐慌出清", "观察次日修复质量"),
        ("恒生指数", "+0.16%；成交3095亿港元", "大盘横住，但内部风格分化", "沪港深与港股QDII"),
        ("恒生科技", "-0.96%", "科技弱于恒指1.12个百分点", "云计算、全球成长"),
        ("美股上周五", "纳指+0.3%，周+1.7%", "海外趋势截至周五仍偏强；周一未收盘", "纳指及海外科技QDII"),
    ]
    for values in observations:
        row = market.add_row()
        for i, text in enumerate(values):
            cell = row.cells[i]
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if i in (2, 3) else WD_ALIGN_PARAGRAPH.CENTER
            format_paragraph(p, 0, 0, 1.0)
            color = RED if i == 1 and text.startswith("-") else INK
            set_font(p.add_run(text), 9.1, i == 0, color)
    set_table_geometry(market, [1500, 2200, 3500, 2160], 120)
    set_table_borders(market)

    add_rich_paragraph(doc, [("技术结论。", {"size": 11, "bold": True, "color": DARK_BLUE}), ("今天是“低开—弱反弹—继续走低”的广泛风险收缩日，且科技成长显著弱于宽基。缩量大跌不等于止跌：它说明买盘退缩，但尚未出现足以确认筹码充分交换的放量企稳。短线需要等待价格、成交和市场宽度同时改善。", {"size": 11, "bold": False, "color": INK})], before=8)
    add_callout(doc, "不能做的技术判断", "底稿没有任何基金净值历史，无法合规计算20日/60日均线、RSI、最大回撤或相对强弱序列。因此本报告只使用当日指数、成交与市场宽度，不伪造精确买卖点。", "gold")

    add_heading(doc, "五、情景、行动与失效条件", 1)
    scenarios = doc.add_table(rows=1, cols=4)
    for i, text in enumerate(["情景", "主观概率", "验证信号", "当前动作"]):
        set_cell_shading(scenarios.rows[0].cells[i], PALE_BLUE)
        p = scenarios.rows[0].cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        format_paragraph(p, 0, 0, 1.0)
        set_font(p.add_run(text), 9.5, True, "000000")
    data = [
        ("基准：高位科技继续震荡消化", "55%", "科技反弹但成交与宽度不足，指数反复", "暂不追涨；先补齐持仓数据"),
        ("偏强：外部风险缓和、业绩兑现", "25%", "科技放量修复，涨跌家数显著转正，强于沪深300", "仅在费率/持有期核验后讨论分批调整"),
        ("偏弱：冲突与利率再升级", "20%", "油价与长债收益率再升，科技继续放量破位", "优先审计科技重叠和超限仓位"),
    ]
    for values in data:
        row = scenarios.add_row()
        for i, text in enumerate(values):
            cell = row.cells[i]
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i == 1 else WD_ALIGN_PARAGRAPH.LEFT
            format_paragraph(p, 0, 0, 1.0)
            set_font(p.add_run(text), 9.2, i == 0, INK)
    set_table_geometry(scenarios, [2550, 1200, 3250, 2360], 120)
    set_table_borders(scenarios)
    add_callout(doc, "风险官结论：暂不行动", "在份额、申购日期、确认净值、赎回费率、现金余额和今日正式净值补齐前，不生成申购或赎回金额。今天的市场下跌本身不足以证明长期逻辑失效，也不足以证明已经见底。", "red")
    add_rich_paragraph(doc, [("推翻“继续震荡”判断的证据：", {"size": 11, "bold": True, "color": DARK_BLUE}), ("连续交易日出现科技相对沪深300转强、成交额回升、上涨家数显著占优，并由中报订单与利润兑现配合。", {"size": 11, "bold": False, "color": INK})])
    add_rich_paragraph(doc, [("推翻“只是技术性调整”的证据：", {"size": 11, "bold": True, "color": DARK_BLUE}), ("科技主题持续放量下跌、核心持仓盈利预期下修、AI资本开支或半导体设备订单出现实质性下调，同时账户真实回撤触发8%纪律线。", {"size": 11, "bold": False, "color": INK})])

    add_heading(doc, "六、下一步需要补齐的信息", 1)
    required = [
        ("每只基金", "当前份额、申购日期、确认日期、申购金额、已确认最新净值"),
        ("交易状态", "创新药赎回申请时间、确认份额/净值、到账金额与费用"),
        ("账户层面", "当前现金余额、账户历史高点、是否还有未迁移基金或在途交易"),
        ("费用与限制", "销售平台当前申购/赎回状态、持有期赎回费、QDII限购与净值时滞"),
    ]
    req_table = doc.add_table(rows=0, cols=2)
    for label, value in required:
        row = req_table.add_row()
        set_cell_shading(row.cells[0], LIGHT)
        for i, text in enumerate((label, value)):
            p = row.cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            format_paragraph(p, 0, 0, 1.0)
            set_font(p.add_run(text), 9.5, i == 0, INK)
            row.cells[i].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    set_table_geometry(req_table, [1900, 7460], 120)
    set_table_borders(req_table)

    add_heading(doc, "七、来源与口径", 1)
    sources = [
        ("本地持仓底稿", "data/holding-intake.csv（截图迁移，时点2026-07-12）", None),
        ("A股收盘与成交", "每日经济新闻｜7月13日A股收盘", "https://www.nbd.com.cn/articles/2026-07-13/4469407.html"),
        ("宽基指数表现", "金融界｜7月13日宽基指数", "https://finance.jrj.com.cn/2026/07/13151657779197.shtml"),
        ("A股跌幅与原因", "东方财富｜A股大跌，发生了什么", "https://finance.eastmoney.com/a/202607133804134796.html"),
        ("ETF资金流", "东方财富｜股票ETF单周净流入近830亿元", "https://finance.eastmoney.com/a/202607133803694914.html"),
        ("港股收盘", "新浪财经/智通｜7月13日港股收盘", "https://finance.sina.com.cn/stock/hkstock/hkstocknews/2026-07-13/doc-inihscma8936694.shtml"),
        ("全球市场与利率", "AP｜7月13日油价、亚洲股市与美债", "https://apnews.com/article/2d6744b09c68b5473d0bc8584b89e60e"),
        ("美股上周五", "AP｜7月10日美国主要指数", "https://apnews.com/article/d3c5b8171e25e98da62831181de9a666"),
        ("统计数据日程", "国家统计局｜2026年主要统计信息发布日程", "https://www.stats.gov.cn/xw/tjxw/tzgg/202512/t20251224_1962137.html"),
    ]
    for label, text, url in sources:
        p = doc.add_paragraph()
        format_paragraph(p, 0, 4, 1.08)
        set_font(p.add_run(f"{label}："), 10, True, INK)
        if url:
            add_hyperlink(p, text, url)
        else:
            set_font(p.add_run(text), 10, False, INK)

    add_rich_paragraph(doc, [("免责声明：", {"size": 9.5, "bold": True, "color": MUTED}), ("本报告用于个人研究与仓位核对，不承诺收益，不替代基金销售平台的正式净值、费率与交易确认。场外基金按未知价申赎，QDII还存在时区与净值确认时滞；任何交易需由你本人核对并手工确认。", {"size": 9.5, "bold": False, "color": MUTED})], before=10, after=0, line=1.05)

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    doc.save(OUTPUT)
    print(OUTPUT)


if __name__ == "__main__":
    build_document()
