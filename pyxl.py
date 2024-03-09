from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import AbsoluteAnchor
from openpyxl.drawing.xdr import XDRPoint2D, XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.styles import Alignment


p2e = pixels_to_EMU


def write_img(ws, fname, x, y):
    img = Image(fname)
    img.anchor = AbsoluteAnchor(
        pos=XDRPoint2D(p2e(x), p2e(y)),
        ext=XDRPositiveSize2D(p2e(img.width), p2e(img.height))
    )
    ws.add_image(img)


def write_row(ws, row, data, start_col=1):
    for i, d in enumerate(data):
        ws.cell(row=row, column=start_col+i).value = d


if __name__ == '__main__':
    wb = Workbook()
    ws = wb.active
    write_img(ws, 'img.png', 100, 100)
    write_row(ws, 1, [1, 2, 3])
    wb.save('out.xlsx')
