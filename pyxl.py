from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import AbsoluteAnchor
from openpyxl.drawing.xdr import XDRPoint2D, XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.styles import Alignment


p2e = pixels_to_EMU


def create_img(f_name, size):
    img = Image(f_name)
    img.width, img.height = size
    return img


def write_img(ws, img, crds):
    img.anchor = AbsoluteAnchor(
        pos=XDRPoint2D(p2e(crds[0]), p2e(crds[1])),
        ext=XDRPositiveSize2D(p2e(img.width), p2e(img.height))
    )
    ws.add_image(img)


def write_row(ws, row, data, start_col=1):
    for i, d in enumerate(data):
        ws.cell(row=row, column=start_col+i).value = d


if __name__ == '__main__':
    IMG_LEFT = 63
    CELL_HEIGHT = 75
    IMG_SZ = 100, 100
    PROMPT_CELL_WIDTH = 20
    wb = Workbook()
    ws = wb.active

    ws.column_dimensions['A'].width = PROMPT_CELL_WIDTH

    for i in range(1, 100):
        ws.row_dimensions[i].height = CELL_HEIGHT

    write_row(ws, 1, ['A bengal cat [mbl] resting on a hammock', 0.0015])
    write_img(ws, create_img('img.png', IMG_SZ), (IMG_LEFT, 0))
    wb.save('out.xlsx')
