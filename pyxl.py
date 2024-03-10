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


class DLSheet(object):
    def __init__(self, ws, img_left):
        self.ws = ws
        self.img_left = img_left

    def init_dims(self, prompt_cell_width, cell_height):
        self.ws.column_dimensions['A'].width = prompt_cell_width
        for i in range(1, 100):
            self.ws.row_dimensions[i].height = cell_height

    def append(self, prompt, loss, img):
        write_row(ws, 1, [prompt, loss])
        write_img(ws, img, (self.img_left, 0))


if __name__ == '__main__':
    IMG_LEFT = 63
    CELL_HEIGHT = 75
    IMG_SZ = 100, 100
    PROMPT_CELL_WIDTH = 20
    wb = Workbook()
    ws = wb.active

    dl_sheet = DLSheet(ws, img_left=IMG_LEFT)
    dl_sheet.init_dims(prompt_cell_width=PROMPT_CELL_WIDTH, cell_height=CELL_HEIGHT)
    dl_sheet.append(
        prompt='A bengal cat [mbl] resting on a hammock', loss=0.0015, img=create_img('img.png', IMG_SZ))

    wb.save('out.xlsx')
