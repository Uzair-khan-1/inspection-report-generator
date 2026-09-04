"""
excel_processor.py
-------------------
Fills the inspection-report Excel template with data and photos.

Template layout (sheet "Report"):
  Rows 1-7   : fixed report header (title, date, report no, column titles row 7)
  Rows 8-13  : THREE ready-made "item blocks" already in the template, each block
               is 2 rows tall:
                   header row  -> "Location / Building:", "Findings", "Corrective
                                   Action", "Completed On", "Status"
                   data row    -> the actual values / photos go here (row height
                                   is already set tall to fit a photo)
  Row 14+    : if the user adds a 4th, 5th, ... item, a new 2-row block is
               inserted and styled to match the existing ones exactly (same
               fonts, borders, fills, row heights) so the template formatting
               is fully preserved no matter how many items are added.

Column mapping (as requested):
  C -> Location / Building   -> AI-cleaned location text, then the AI-cleaned
                                 finding text on the line below, in the same cell.
  D -> Findings               -> the BEFORE photo, stretched to fill the entire cell.
  E -> Corrective Action      -> the AFTER photo, stretched to fill the entire cell.
  F -> Completed On           -> date.
  G -> Status                 -> text, e.g. "Open" / "Closed".

Photos are resized locally (Pillow) and stretched to exactly cover their
target cell (width AND height), so every photo -- whatever size/shape it
was uploaded as -- ends up the same size on the sheet, filling the cell
edge to edge. Column E (Corrective Action) is widened to match column D
(Findings) so the Before/After photos are always the same size as each
other. The user never has to resize or move anything by hand.
"""

import copy
import io
from datetime import date

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils import get_column_letter
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill
from PIL import Image as PILImage

SHEET_NAME = "Report"
FIRST_HEADER_ROW = 8      # row of the very first "Location / Building:" header
BLOCK_SIZE = 2            # header row + data row
EXISTING_BLOCKS = 3       # the template ships with 3 ready-made blocks (rows 8-13)

FINDINGS_COL = 4          # D - photo (before)
CORRECTIVE_COL = 5        # E - photo (after)

HEADER_TEXT = {
    3: "Location / Building: ",
    4: "Findings ",
    5: "Corrective Action",
    6: "Completed On",
    7: "Status",
}


# --------------------------------------------------------------------------- #
# small geometry helpers
# --------------------------------------------------------------------------- #
def _col_width_px(width):
    """Approximate Excel column width (characters) -> pixels (Calibri 11)."""
    if not width:
        width = 8.43
    return int(round(width * 7 + 5))


def _row_height_px(height):
    """Excel row height (points) -> pixels."""
    if not height:
        height = 15
    return int(round(height * 4 / 3))


def _stretch_to_fill(image_bytes, target_w, target_h):
    """Resize (in memory) to EXACTLY target_w x target_h, ignoring the
    original aspect ratio, so the photo fully covers the cell no matter
    what shape it was uploaded as."""
    target_w = max(int(target_w), 10)
    target_h = max(int(target_h), 10)
    im = PILImage.open(io.BytesIO(image_bytes))
    try:
        im = PILImage.exif_transpose(im)  # respect phone-camera orientation
    except Exception:
        pass
    if im.mode not in ("RGB", "RGBA"):
        im = im.convert("RGB")
    im = im.resize((target_w, target_h), PILImage.LANCZOS)
    buf = io.BytesIO()
    im.save(buf, format="PNG")
    buf.seek(0)
    return buf


def _fill_cell_with_image(ws, image_bytes, col_idx, row_idx, cell_w_px, cell_h_px):
    """Stretch `image_bytes` to exactly cell_w_px x cell_h_px and anchor it
    to fully cover the cell (0 offset, no gaps)."""
    buf = _stretch_to_fill(image_bytes, cell_w_px, cell_h_px)
    marker = AnchorMarker(col=col_idx - 1, colOff=0, row=row_idx - 1, rowOff=0)
    size = XDRPositiveSize2D(pixels_to_EMU(cell_w_px), pixels_to_EMU(cell_h_px))
    img = XLImage(buf)
    img.anchor = OneCellAnchor(_from=marker, ext=size)
    ws.add_image(img)


# --------------------------------------------------------------------------- #
# block (row) management
# --------------------------------------------------------------------------- #
def _copy_row_style(ws, src_row, dst_row, ncols=7):
    for c in range(1, ncols + 1):
        src = ws.cell(row=src_row, column=c)
        dst = ws.cell(row=dst_row, column=c)
        dst.font = copy.copy(src.font)
        dst.fill = copy.copy(src.fill)
        dst.border = copy.copy(src.border)
        dst.alignment = copy.copy(src.alignment)
        dst.number_format = src.number_format
    src_h = ws.row_dimensions[src_row].height
    if src_h:
        ws.row_dimensions[dst_row].height = src_h


def _add_status_conditional_formatting(ws, cell_ref):
    """The template color-codes the Status cell (red = Open, green = Closed)
    via conditional formatting on its 3 built-in rows. New rows we insert
    need the same two rules so status colors keep working automatically."""
    red = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
    green = PatternFill(start_color="FF00B050", end_color="FF00B050", fill_type="solid")
    ws.conditional_formatting.add(
        cell_ref, FormulaRule(formula=[f'NOT(ISERROR(SEARCH("OPEN",{cell_ref})))'], fill=red)
    )
    ws.conditional_formatting.add(
        cell_ref, FormulaRule(formula=[f'NOT(ISERROR(SEARCH("CLOSED",{cell_ref})))'], fill=green)
    )


def _ensure_block(ws, index):
    """Return (header_row, data_row) for the item at `index` (0-based),
    inserting and styling a new block if the template's ready-made blocks
    have run out."""
    header_row = FIRST_HEADER_ROW + index * BLOCK_SIZE
    data_row = header_row + 1

    if index < EXISTING_BLOCKS:
        return header_row, data_row

    ws.insert_rows(header_row, amount=BLOCK_SIZE)
    _copy_row_style(ws, FIRST_HEADER_ROW, header_row)
    _copy_row_style(ws, FIRST_HEADER_ROW + 1, data_row)
    for col, text in HEADER_TEXT.items():
        ws.cell(row=header_row, column=col, value=text)
    _add_status_conditional_formatting(ws, f"G{data_row}")
    return header_row, data_row


def _ensure_equal_photo_columns(ws):
    """Widen Corrective Action (E) to match Findings (D) so both the Before
    and After photos always render at the same size."""
    d_dim = ws.column_dimensions[get_column_letter(FINDINGS_COL)]
    e_dim = ws.column_dimensions[get_column_letter(CORRECTIVE_COL)]
    if d_dim.width and (not e_dim.width or e_dim.width != d_dim.width):
        e_dim.width = d_dim.width


# --------------------------------------------------------------------------- #
# public API
# --------------------------------------------------------------------------- #
def generate_report(items, template_path, output_path):
    """
    items: list of dicts, each with:
        location       (str)  - AI-cleaned "Location / Building"
        finding        (str)  - AI-cleaned "Findings" text
        before_bytes   (bytes or None) - BEFORE photo
        after_bytes    (bytes or None) - AFTER photo
        completed_on   (date, optional, default = today)
        status         (str, optional, default = "Open")

    Returns output_path.
    """
    wb = load_workbook(template_path)
    ws = wb[SHEET_NAME]
    _ensure_equal_photo_columns(ws)

    # All data rows share the same row height (the template's tall photo
    # row), so every photo across every observation ends up the same size.
    row_h = _row_height_px(ws.row_dimensions[FIRST_HEADER_ROW + 1].height)
    d_w = _col_width_px(ws.column_dimensions[get_column_letter(FINDINGS_COL)].width)
    e_w = _col_width_px(ws.column_dimensions[get_column_letter(CORRECTIVE_COL)].width)

    for i, item in enumerate(items):
        header_row, data_row = _ensure_block(ws, i)

        location = item.get("location", "")
        finding = item.get("finding", "")
        ws.cell(row=data_row, column=3, value=f"{location}\n\n{finding}".strip())
        ws.cell(row=data_row, column=6, value=item.get("completed_on") or date.today())
        ws.cell(row=data_row, column=7, value=item.get("status") or "Open")

        if item.get("before_bytes"):
            _fill_cell_with_image(ws, item["before_bytes"], col_idx=FINDINGS_COL,
                                   row_idx=data_row, cell_w_px=d_w, cell_h_px=row_h)

        if item.get("after_bytes"):
            _fill_cell_with_image(ws, item["after_bytes"], col_idx=CORRECTIVE_COL,
                                   row_idx=data_row, cell_w_px=e_w, cell_h_px=row_h)

    wb.save(output_path)
    return output_path
