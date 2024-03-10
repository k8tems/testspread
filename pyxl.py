from openpyxl import Workbook
from openpyxl.drawing.image import Image
from PIL import Image as PILImage
from openpyxl.drawing.spreadsheet_drawing import AbsoluteAnchor
from openpyxl.drawing.xdr import XDRPoint2D, XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU


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
        self.cur_row = 1

    def init_dims(self, prompt_cell_width, cell_height):
        self.ws.column_dimensions['A'].width = prompt_cell_width
        for i in range(1, 100):
            self.ws.row_dimensions[i].height = cell_height

    @classmethod
    def create(cls, ws, img_left, img_sz, prompt_cell_width, cell_height):
        obj = cls(ws, img_left, img_sz)
        obj.init_dims(prompt_cell_width, cell_height)
        return obj

    def append(self, prompt, loss, pil_imgs):
        write_row(ws, self.cur_row, [prompt, loss])
        imgs = [Image(pi) for pi in pil_imgs]
        for i, im in enumerate(imgs):
            im.width, im.height = self.img_sz
            write_img(ws, im, (self.img_left + (i*self.img_sz[0]), (self.cur_row-1) * self.img_sz[0]))
        self.cur_row += 1


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

    pil_img_1_1 = PILImage.open('img1_1.png')  # this is what the incoming data will look like in the NB
    pil_img_1_2 = PILImage.open('img1_2.png')
    dl_sheet.append(
        prompt='A cinematic photo of a fat bengal cat [mbl] sitting in front of a pizza',
        loss=0.0015, pil_imgs=[pil_img_1_1, pil_img_1_2])

    pil_img_2_1 = PILImage.open('img2_1.png')
    pil_img_2_2 = PILImage.open('img2_2.png')
    dl_sheet.append(
        prompt='A photo of a miniature bengal cat [mbl] held by a human hand',
        loss=0.0035, pil_imgs=[pil_img_2_1, pil_img_2_2])

    wb.save('out.xlsx')
