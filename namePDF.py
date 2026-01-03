
"""
produce a pdf
"""

import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor

in_name = "DataSheet.xlsx"
df = pd.read_excel(in_name)
df['sport'] = df['sport'].fillna('N/A')

output_path = "NamePDF.pdf"
c = canvas.Canvas(output_path, pagesize=letter)

PAGE_WIDTH, PAGE_HEIGHT = letter
TAGS_PER_ROW = 2
TAGS_PER_COL = 4
TAG_WIDTH = (PAGE_WIDTH - 2 * 15 * mm - (TAGS_PER_ROW - 1) * 10 * mm) / TAGS_PER_ROW
TAG_WIDTH = TAG_WIDTH * 0.95
TAG_HEIGHT = (PAGE_HEIGHT - 2 * 20 * mm - (TAGS_PER_COL - 1) * 10 * mm) / TAGS_PER_COL
MARGIN_X = 15 * mm
MARGIN_Y = 20 * mm
GAP_X = 10 * mm
GAP_Y = 10 * mm

STAR_IMAGE_PATH = "star.png"
STAR_WIDTH = 40
STAR_HEIGHT = 40

def draw_name_tag(c, x, y, name, sport, room1, room2, food):
    try:
        r1 = str(int(room1)) if pd.notna(room1) else "outdoor"
        r2 = str(int(room2)) if pd.notna(room2) else "outdoor"
    except:
        r1, r2 = "?", "?"

    label = f"{sport}-Lunch--{r1}-{r2}"

    c.setStrokeColor(HexColor("#0066cc"))
    c.setLineWidth(4)
    c.rect(x, y, TAG_WIDTH, TAG_HEIGHT)

    c.setFont("Helvetica", 18)
    c.setFillColor("red")
    c.drawCentredString(x + TAG_WIDTH / 2, y + TAG_HEIGHT - 23, label)

    c.setFont("Helvetica-Bold", 28)
    c.setFillColor("black")
    c.drawCentredString(x + TAG_WIDTH / 2, y + TAG_HEIGHT / 2, name)

    if str(food).strip().upper() == "Y":
        padding = 6
        star_x = x + TAG_WIDTH - STAR_WIDTH - padding
        star_y = y + padding
        c.drawImage("star.png", star_x, star_y, width=STAR_WIDTH, height=STAR_HEIGHT, mask='auto')

row, col = 0, 0
for _, row_data in df.iterrows():
    x = MARGIN_X + col * (TAG_WIDTH + GAP_X)
    y = PAGE_HEIGHT - MARGIN_Y - (row + 1) * TAG_HEIGHT - row * GAP_Y
    draw_name_tag(
        c,
        x, y,
        row_data["name"],
        row_data["sport"],
        row_data["room1"],
        row_data["room2"],
        row_data["food"]
    )
    col += 1
    if col >= TAGS_PER_ROW:
        col = 0
        row += 1
    if row >= TAGS_PER_COL:
        c.showPage()
        row, col = 0, 0

c.save()
print("successful")


