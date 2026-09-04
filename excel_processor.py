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
  C -> Location / Building   (AI-cleaned text)
  D -> Findings               (AI-cleaned text + the BEFORE photo, inserted below the text)
  E -> Corrective Action      (the AFTER photo)
  F -> Completed On           (date)
  G -> Status                 (text, e.g. "Open" / "Closed")

All images are resized locally (Pillow) to fit neatly inside their target
cell while preserving aspect ratio -- the user never has to resize or move
anything by hand.
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

HEADER_TEXT = {
    3: "Location / Building: ",
    4: "Findings ",
    5: "Corrective Action",
    6: "Completed On",
    7: "Status",
}

# Fraction of the data-row height reserved for the wrapped "Findings" text
# before the BEFORE photo starts (D column only).
FINDINGS_TEXT_FRACTION = 0.30
CELL_PADDING_PX = 6


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


def _resize_to_fit(image_bytes, max_w, max_h):
    """Resize (in memory) preserving aspect ratio so it fits max_w x max_h."""
    max_w = max(int(max_w), 10)
    max_h = max(int(max_h), 10)
    im = PILImage.open(io.BytesIO(image_bytes))
    try:
        im = PILImage.exif_transpose(im)  # respect phone-camera orientation
    except Exception:
        pass
    if im.mode not in ("RGB", "RGBA"):
        im = im.convert("RGB")
    im.thumbnail((max_w, max_h), PILImage.LANCZOS)
    buf = io.BytesIO()
    im.save(buf, format="PNG")
    buf.seek(0)
    return buf, im.width, im.height


def _place_image(ws, image_bytes, col_idx, row_idx, cell_w_px, cell_h_px,
                  top_offset_frac=0.0):
    """Resize `image_bytes` to fit the space below `top_offset_frac` of the
    cell, then anchor it centered in that space -- purely local, no manual
    positioning ever required by the user."""
    top_offset_px = cell_h_px * top_offset_frac
    avail_w = cell_w_px - 2 * CELL_PADDING_PX
    avail_h = cell_h_px - top_offset_px - 2 * CELL_PADDING_PX
    buf, w, h = _resize_to_fit(image_bytes, avail_w, avail_h)

    off_x = max(0, (cell_w_px - w) // 2)
    off_y = top_offset_px + max(0, (avail_h - h) // 2)

    marker = AnchorMarker(
        col=col_idx - 1, colOff=pixels_to_EMU(off_x),
        row=row_idx - 1, rowOff=pixels_to_EMU(off_y),
    )
    size = XDRPositiveSize2D(pixels_to_EMU(w), pixels_to_EMU(h))
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

    for i, item in enumerate(items):
        header_row, data_row = _ensure_block(ws, i)

        ws.cell(row=data_row, column=3, value=item.get("location", ""))
        ws.cell(row=data_row, column=4, value=item.get("finding", ""))
        ws.cell(row=data_row, column=6, value=item.get("completed_on") or date.today())
        ws.cell(row=data_row, column=7, value=item.get("status") or "Open")

        d_w = _col_width_px(ws.column_dimensions[get_column_letter(4)].width)
        e_w = _col_width_px(ws.column_dimensions[get_column_letter(5)].width)
        row_h = _row_height_px(ws.row_dimensions[data_row].height)

        if item.get("before_bytes"):
            _place_image(ws, item["before_bytes"], col_idx=4, row_idx=data_row,
                         cell_w_px=d_w, cell_h_px=row_h,
                         top_offset_frac=FINDINGS_TEXT_FRACTION)

        if item.get("after_bytes"):
            _place_image(ws, item["after_bytes"], col_idx=5, row_idx=data_row,
                         cell_w_px=e_w, cell_h_px=row_h,
                         top_offset_frac=0.0)

    wb.save(output_path)
    return output_path
