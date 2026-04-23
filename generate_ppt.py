import csv
import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# ── 配置 ──────────────────────────────────────────────
CSV_DIR = r"D:\workspace\csv_handle"
OUTPUT = os.path.join(CSV_DIR, "report.pptx")
ROWS_PER_PAGE = 5

PRISM_BLUE = RGBColor(0x00, 0x70, 0xC0)
DARK_GRAY = RGBColor(0x33, 0x33, 0x33)
MEDIUM_GRAY = RGBColor(0x66, 0x66, 0x66)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GRAY = RGBColor(0xF2, 0xF2, 0xF2)
TABLE_HEADER_BG = RGBColor(0x00, 0x70, 0xC0)
TABLE_HEADER_FG = RGBColor(0xFF, 0xFF, 0xFF)
TABLE_ROW_ALT = RGBColor(0xE8, 0xF0, 0xFE)

SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)

csv_files = [
    ("employees.csv", "Employee Information & Payroll", "Employee personal details, salary structure and deductions"),
    ("products.csv", "Product Inventory Management", "Product catalog, pricing, stock levels and warehouse locations"),
    ("orders.csv", "Order Records", "Customer orders with payment, shipping and fulfillment status"),
]


def set_cell_text(cell, text, font_size=10, bold=False, color=DARK_GRAY, alignment=PP_ALIGN.LEFT):
    cell.text = ""
    p = cell.text_frame.paragraphs[0]
    p.alignment = alignment
    run = p.add_run()
    run.text = str(text)
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = "Calibri"
    cell.vertical_anchor = 1  # middle


def set_cell_fill(cell, color):
    cell.fill.solid()
    cell.fill.fore_color.rgb = color


def add_title_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank

    # left accent bar
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.35), SLIDE_HEIGHT)
    bar.fill.solid()
    bar.fill.fore_color.rgb = PRISM_BLUE
    bar.line.fill.background()

    # title
    txBox = slide.shapes.add_textbox(Inches(1.2), Inches(1.8), Inches(10), Inches(1.2))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = "CSV Data Report"
    run.font.size = Pt(44)
    run.font.bold = True
    run.font.color.rgb = PRISM_BLUE
    run.font.name = "Calibri"

    # subtitle
    txBox2 = slide.shapes.add_textbox(Inches(1.2), Inches(3.1), Inches(10), Inches(0.6))
    tf2 = txBox2.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.alignment = PP_ALIGN.LEFT
    run2 = p2.add_run()
    run2.text = "Overview of three datasets from CSV files"
    run2.font.size = Pt(22)
    run2.font.color.rgb = MEDIUM_GRAY
    run2.font.name = "Calibri"

    # divider
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.2), Inches(3.9), Inches(3), Pt(3))
    line.fill.solid()
    line.fill.fore_color.rgb = PRISM_BLUE
    line.line.fill.background()

    # content list
    txBox3 = slide.shapes.add_textbox(Inches(1.2), Inches(4.3), Inches(10), Inches(2.5))
    tf3 = txBox3.text_frame
    tf3.word_wrap = True
    for i, (_, title, desc) in enumerate(csv_files):
        if i > 0:
            p_empty = tf3.add_paragraph()
            p_empty.space_before = Pt(6)

        p = tf3.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        run_num = p.add_run()
        run_num.text = f"{i + 1}. "
        run_num.font.size = Pt(18)
        run_num.font.bold = True
        run_num.font.color.rgb = PRISM_BLUE
        run_num.font.name = "Calibri"

        run_title = p.add_run()
        run_title.text = f"{title}"
        run_title.font.size = Pt(18)
        run_title.font.bold = True
        run_title.font.color.rgb = DARK_GRAY
        run_title.font.name = "Calibri"

        p_desc = tf3.add_paragraph()
        p_desc.alignment = PP_ALIGN.LEFT
        p_desc.level = 1
        run_desc = p_desc.add_run()
        run_desc.text = f"     {desc}"
        run_desc.font.size = Pt(14)
        run_desc.font.color.rgb = MEDIUM_GRAY
        run_desc.font.name = "Calibri"


def add_section_slide(prs, title, row_count, col_count):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_WIDTH, Inches(0.15))
    bar.fill.solid()
    bar.fill.fore_color.rgb = PRISM_BLUE
    bar.line.fill.background()

    txBox = slide.shapes.add_textbox(Inches(1), Inches(2.4), Inches(11), Inches(1.0))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = title
    run.font.size = Pt(36)
    run.font.bold = True
    run.font.color.rgb = PRISM_BLUE
    run.font.name = "Calibri"

    txBox2 = slide.shapes.add_textbox(Inches(1), Inches(3.6), Inches(11), Inches(0.6))
    tf2 = txBox2.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.alignment = PP_ALIGN.CENTER
    run2 = p2.add_run()
    run2.text = f"{row_count} rows  x  {col_count} columns"
    run2.font.size = Pt(20)
    run2.font.color.rgb = MEDIUM_GRAY
    run2.font.name = "Calibri"


def add_table_slide(prs, headers, data_chunk, page_num, total_pages, section_title):
    num_cols = len(headers)
    num_rows = len(data_chunk) + 1  # +1 header

    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # top bar
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_WIDTH, Inches(0.08))
    bar.fill.solid()
    bar.fill.fore_color.rgb = PRISM_BLUE
    bar.line.fill.background()

    # section title (small)
    txTitle = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(8), Inches(0.5))
    tf_title = txTitle.text_frame
    p_title = tf_title.paragraphs[0]
    run_t = p_title.add_run()
    run_t.text = section_title
    run_t.font.size = Pt(14)
    run_t.font.bold = True
    run_t.font.color.rgb = PRISM_BLUE
    run_t.font.name = "Calibri"

    # page indicator
    txPage = slide.shapes.add_textbox(Inches(9.5), Inches(0.2), Inches(3.5), Inches(0.5))
    tf_page = txPage.text_frame
    p_page = tf_page.paragraphs[0]
    p_page.alignment = PP_ALIGN.RIGHT
    run_pg = p_page.add_run()
    run_pg.text = f"Page {page_num} / {total_pages}"
    run_pg.font.size = Pt(12)
    run_pg.font.color.rgb = MEDIUM_GRAY
    run_pg.font.name = "Calibri"

    # table
    table_width = Inches(12.3)
    left = Inches(0.5)
    top = Inches(0.85)
    table_height = Inches(5.8)

    tbl_shape = slide.shapes.add_table(num_rows, num_cols, left, top, table_width, table_height)
    tbl = tbl_shape.table

    # set column widths evenly
    col_w = int(table_width / num_cols)
    for i in range(num_cols):
        tbl.columns[i].width = col_w

    # header row
    for ci, h in enumerate(headers):
        cell = tbl.cell(0, ci)
        set_cell_text(cell, h, font_size=10, bold=True, color=TABLE_HEADER_FG, alignment=PP_ALIGN.CENTER)
        set_cell_fill(cell, TABLE_HEADER_BG)

    # data rows
    for ri, row in enumerate(data_chunk):
        for ci, val in enumerate(row):
            cell = tbl.cell(ri + 1, ci)
            set_cell_text(cell, val, font_size=9)
            if ri % 2 == 1:
                set_cell_fill(cell, TABLE_ROW_ALT)


def read_csv(filepath):
    with open(filepath, newline="", encoding="utf-8") as f:
        reader = csv.reader(f)
        headers = next(reader)
        data = [row for row in reader]
    return headers, data


def main():
    prs = Presentation()
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT

    # 1) Cover slide
    add_title_slide(prs)

    # 2) For each CSV: section divider + table pages
    for filename, title, _ in csv_files:
        filepath = os.path.join(CSV_DIR, filename)
        headers, data = read_csv(filepath)
        total_rows = len(data)
        total_pages = (total_rows + ROWS_PER_PAGE - 1) // ROWS_PER_PAGE

        # section divider
        add_section_slide(prs, title, total_rows, len(headers))

        # table pages
        for page in range(total_pages):
            start = page * ROWS_PER_PAGE
            end = min(start + ROWS_PER_PAGE, total_rows)
            chunk = data[start:end]
            add_table_slide(prs, headers, chunk, page + 1, total_pages, title)

    prs.save(OUTPUT)
    print(f"Done -> {OUTPUT}")
    print(f"Total slides: {len(prs.slides)}")


if __name__ == "__main__":
    main()
