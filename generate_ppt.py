"""
generate_ppt.py
基于 report_template.pptx 模板，读取 CSV 数据并生成最终 PPT 报告。

模板包含 3 张预设幻灯片（索引 0=封面, 1=分隔页, 2=数据页），
本脚本克隆模板幻灯片，填入文字和表格数据后删除原始模板页。
"""
import csv
import os
import copy
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn

# ── 配置 ──────────────────────────────────────────────
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(SCRIPT_DIR, "report_template.pptx")
OUTPUT = os.path.join(SCRIPT_DIR, "report.pptx")

PRISM_BLUE = RGBColor(0x00, 0x70, 0xC0)
DARK_GRAY = RGBColor(0x33, 0x33, 0x33)
MEDIUM_GRAY = RGBColor(0x66, 0x66, 0x66)
TABLE_HEADER_BG = RGBColor(0x00, 0x70, 0xC0)
TABLE_HEADER_FG = RGBColor(0xFF, 0xFF, 0xFF)
TABLE_ROW_ALT = RGBColor(0xE8, 0xF0, 0xFE)

csv_files = [
    ("employees.csv", "Employee Information & Payroll",
     "Employee personal details, salary structure and deductions"),
    ("products.csv", "Product Inventory Management",
     "Product catalog, pricing, stock levels and warehouse locations"),
    ("orders.csv", "Order Records",
     "Customer orders with payment, shipping and fulfillment status"),
]


# ── helpers ────────────────────────────────────────────
def _find_shape(slide, name):
    """Return shape by name, or None."""
    for sp in slide.shapes:
        if sp.name == name:
            return sp
    return None


def _set_text(shape, text, size=None, bold=None, color=None, align=None):
    """Replace text while preserving template formatting when no overrides given."""
    tf = shape.text_frame
    tf.word_wrap = True

    if tf.paragraphs and tf.paragraphs[0].runs:
        # shape already has a run with template formatting — keep it
        run = tf.paragraphs[0].runs[0]
        run.text = text
        if size is not None:
            run.font.size = Pt(size)
        if bold is not None:
            run.font.bold = bold
        if color is not None:
            run.font.color.rgb = color
        if align is not None:
            tf.paragraphs[0].alignment = align
    else:
        # no existing run, create fresh
        tf.clear()
        p = tf.paragraphs[0]
        if align is not None:
            p.alignment = align
        r = p.add_run()
        r.text = text
        if size:
            r.font.size = Pt(size)
        if bold is not None:
            r.font.bold = bold
        if color:
            r.font.color.rgb = color
        r.font.name = "Calibri"


def _set_cell(cell, text, font_size=9, bold=False, color=DARK_GRAY, align=PP_ALIGN.LEFT):
    cell.text = ""
    p = cell.text_frame.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = str(text)
    r.font.size = Pt(font_size)
    r.font.bold = bold
    r.font.color.rgb = color
    r.font.name = "Calibri"
    cell.vertical_anchor = 1  # middle


def _fill_cell(cell, color):
    cell.fill.solid()
    cell.fill.fore_color.rgb = color


def _clone_slide(prs, template_slide, exclude_names=None):
    """Deep-copy a template slide into the presentation and return the new slide.
    exclude_names: list of shape names to skip (e.g. 'SampleTable').
    """
    exclude = set(exclude_names or [])
    slide_layout = prs.slide_layouts[6]  # blank
    new_slide = prs.slides.add_slide(slide_layout)

    # remove default elements from new slide's spTree
    sp_tree = new_slide.shapes._spTree
    for elem in list(sp_tree):
        if elem.tag.endswith("}sp") or elem.tag.endswith("}graphicFrame"):
            sp_tree.remove(elem)

    # copy all shape elements from template, skipping excluded ones
    src_tree = template_slide.shapes._spTree
    for elem in src_tree:
        if elem.tag.endswith("}sp") or elem.tag.endswith("}graphicFrame") or elem.tag.endswith("}pic"):
            # check if this element's name is in the exclude list
            nvSpPr = elem.find(qn("p:nvSpPr"))
            if nvSpPr is not None:
                nvPr = nvSpPr.find(qn("p:cNvPr"))
                if nvPr is not None and nvPr.get("name") in exclude:
                    continue
            # for graphicFrame (tables), check nvGrpSpPr/cNvPr
            nvGrpFr = elem.find(qn("p:nvGraphicFramePr"))
            if nvGrpFr is not None:
                nvPr = nvGrpFr.find(qn("p:cNvPr"))
                if nvPr is not None and nvPr.get("name") in exclude:
                    continue
            sp_tree.append(copy.deepcopy(elem))

    return new_slide


def read_csv(filepath):
    with open(filepath, newline="", encoding="utf-8") as f:
        reader = csv.reader(f)
        headers = next(reader)
        data = [row for row in reader]
    return headers, data


# ── slide builders ────────────────────────────────────
def build_cover(prs, template_slide):
    slide = _clone_slide(prs, template_slide)

    sh = _find_shape(slide, "Title")
    if sh:
        _set_text(sh, "CSV Data Report")

    sh = _find_shape(slide, "Subtitle")
    if sh:
        _set_text(sh, "Overview of three datasets from CSV files")

    sh = _find_shape(slide, "ContentList")
    if sh:
        tf = sh.text_frame
        tf.word_wrap = True
        # preserve the template run's font as base style
        base_font = None
        if tf.paragraphs and tf.paragraphs[0].runs:
            base_font = tf.paragraphs[0].runs[0].font
        tf.clear()

        for i, (_, title, desc) in enumerate(csv_files):
            if i > 0:
                tf.add_paragraph().space_before = Pt(6)

            p = tf.add_paragraph()
            p.alignment = PP_ALIGN.LEFT
            rn = p.add_run()
            rn.text = f"{i + 1}. "
            rn.font.size = Pt(18)
            rn.font.bold = True
            if base_font and base_font.color and base_font.color.rgb:
                rn.font.color.rgb = base_font.color.rgb
            else:
                rn.font.color.rgb = PRISM_BLUE
            rn.font.name = base_font.name if base_font and base_font.name else "Calibri"

            rt = p.add_run()
            rt.text = title
            rt.font.size = Pt(18)
            rt.font.bold = True
            if base_font and base_font.color and base_font.color.rgb:
                # slightly darker than accent for title text
                rt.font.color.rgb = DARK_GRAY
            else:
                rt.font.color.rgb = DARK_GRAY
            rt.font.name = base_font.name if base_font and base_font.name else "Calibri"

            pd = tf.add_paragraph()
            pd.alignment = PP_ALIGN.LEFT
            rd = pd.add_run()
            rd.text = f"     {desc}"
            rd.font.size = Pt(14)
            if base_font and base_font.color and base_font.color.rgb:
                rd.font.color.rgb = MEDIUM_GRAY
            else:
                rd.font.color.rgb = MEDIUM_GRAY
            rd.font.name = base_font.name if base_font and base_font.name else "Calibri"


def build_section(prs, template_slide, title, rows, cols):
    slide = _clone_slide(prs, template_slide)

    sh = _find_shape(slide, "SectionTitle")
    if sh:
        _set_text(sh, title, align=PP_ALIGN.CENTER)

    sh = _find_shape(slide, "DimensionInfo")
    if sh:
        _set_text(sh, f"{rows} rows  x  {cols} columns", align=PP_ALIGN.CENTER)


def build_table(prs, template_slide, headers, data_chunk,
                page_num, total_pages, section_title):
    num_cols = len(headers)
    num_rows = len(data_chunk) + 1

    # clone slide but skip the sample table
    slide = _clone_slide(prs, template_slide, exclude_names=["SampleTable"])

    sh = _find_shape(slide, "SectionTitle")
    if sh:
        _set_text(sh, section_title)

    sh = _find_shape(slide, "PageInfo")
    if sh:
        _set_text(sh, f"Page {page_num} / {total_pages}", align=PP_ALIGN.RIGHT)

    # table height: top at 0.85", bottom at 7.2" → max ~6.35"
    TABLE_TOP = Inches(0.85)
    TABLE_BOTTOM = Inches(7.2)
    table_height = TABLE_BOTTOM - TABLE_TOP

    # add data table
    tbl_shape = slide.shapes.add_table(
        num_rows, num_cols,
        Inches(0.5), TABLE_TOP,
        Inches(12.3), table_height,
    )
    tbl = tbl_shape.table
    col_w = int(Inches(12.3) / num_cols)
    for i in range(num_cols):
        tbl.columns[i].width = col_w

    # header
    for ci, h in enumerate(headers):
        c = tbl.cell(0, ci)
        _set_cell(c, h, font_size=10, bold=True,
                  color=TABLE_HEADER_FG, align=PP_ALIGN.CENTER)
        _fill_cell(c, TABLE_HEADER_BG)

    # data
    for ri, row in enumerate(data_chunk):
        for ci, val in enumerate(row):
            c = tbl.cell(ri + 1, ci)
            _set_cell(c, val, font_size=9)
            if ri % 2 == 1:
                _fill_cell(c, TABLE_ROW_ALT)


# ── main ───────────────────────────────────────────────
def _read_rows_per_page(template_slide):
    """Read the data rows count from the sample table in DATA_TABLE template slide.
    Returns (total_rows - 1) as the number of data rows per page.
    """
    for sp in template_slide.shapes:
        if sp.name == "SampleTable" and sp.has_table:
            total_rows = len(sp.table.rows)
            data_rows = total_rows - 1  # exclude header row
            print(f"  SampleTable detected: {total_rows} rows total -> {data_rows} data rows per page")
            return data_rows
    # fallback
    print("  Warning: SampleTable not found in template, defaulting to 5 rows per page")
    return 5


def main():
    prs = Presentation(TEMPLATE)

    # grab the 3 template slides (will be deleted later)
    tmpl_cover   = prs.slides[0]
    tmpl_section = prs.slides[1]
    tmpl_table   = prs.slides[2]

    # read rows-per-page from template's sample table
    rows_per_page = _read_rows_per_page(tmpl_table)

    # ── build cover ───────────────────────────────────
    build_cover(prs, tmpl_cover)

    # ── build each CSV section ────────────────────────
    for filename, title, _ in csv_files:
        filepath = os.path.join(SCRIPT_DIR, filename)
        headers, data = read_csv(filepath)
        total_rows = len(data)
        total_pages = (total_rows + rows_per_page - 1) // rows_per_page

        build_section(prs, tmpl_section, title, total_rows, len(headers))

        for page in range(total_pages):
            start = page * rows_per_page
            end = min(start + rows_per_page, total_rows)
            build_table(prs, tmpl_table, headers, data[start:end],
                        page + 1, total_pages, title)

    # ── remove the 3 original template slides ─────────
    # delete in reverse order to keep indices stable
    for idx in [2, 1, 0]:
        rId = prs.slides._sldIdLst[idx].get(qn("r:id"))
        prs.part.drop_rel(rId)
        prs.slides._sldIdLst.remove(prs.slides._sldIdLst[idx])

    prs.save(OUTPUT)
    print(f"Done -> {OUTPUT}")
    print(f"Total slides: {len(prs.slides)}")


if __name__ == "__main__":
    main()
