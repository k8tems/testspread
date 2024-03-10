from openpyxl import Workbook
from openpyxl.drawing.image import Image
from PIL import Image as PILImage
from openpyxl.drawing.spreadsheet_drawing import AbsoluteAnchor
from openpyxl.drawing.xdr import XDRPoint2D, XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.styles import Alignment


p2e = pixels_to_EMU


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
    def __init__(self, ws, img_left, img_sz):
        self.ws = ws
        self.img_left = img_left
        self.img_sz = img_sz

    def init_dims(self, prompt_cell_width, cell_height):
        self.ws.column_dimensions['A'].width = prompt_cell_width
        for i in range(1, 100):
            self.ws.row_dimensions[i].height = cell_height

    @classmethod
    def create(cls, ws, img_left, img_sz, prompt_cell_width, cell_height):
        obj = cls(ws, img_left, img_sz)
        obj.init_dims(prompt_cell_width, cell_height)
        return obj

    def append(self, prompt, loss, pil_img):
        write_row(ws, 1, [prompt, loss])
        img = Image(pil_img)
        img.width, img.height = self.img_sz
        write_img(ws, img, (self.img_left, 0))


if __name__ == '__main__':
    IMG_LEFT = 210
    CELL_HEIGHT = 75
    IMG_SZ = 100, 100
    PROMPT_CELL_WIDTH = 20
    wb = Workbook()
    ws = wb.active

    dl_sheet = DLSheet.create(
        ws, img_left=IMG_LEFT, img_sz=IMG_SZ,
        prompt_cell_width=PROMPT_CELL_WIDTH, cell_height=CELL_HEIGHT)

    pil_img_1 = PILImage.open('img1.png')  # this is what the incoming data will look like in the NB
    dl_sheet.append(
        prompt='A cinematic photo of a fat bengal cat [mbl] sitting in front of a pizza',
        loss=0.0015, pil_img=pil_img_1)

    pil_img_2 = PILImage.open('img2.png')
    dl_sheet.append(
        prompt='A photo of a miniature bengal cat [mbl] held by a human hand', loss=0.0035, pil_img=pil_img_2)

    wb.save('out.xlsx')
