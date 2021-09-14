
import os
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Cm, Pt
from pptx.dml.color import RGBColor


# x 좌표
def x_point(cnt: int):
    if cnt % 4 == 1:
        x = 0
    elif cnt % 4 == 2:
        x = Cm(11.18)
    elif cnt % 4 == 3:
        x = Cm(22.36)
    elif cnt % 4 == 0:
        x = Cm(33.54)

    return x


# 슬라이드에 넣기
def make_slide(x_location: int):

    pic = slide.shapes.add_picture('tmp.png', x_location, img_y, ppt_img_width, ppt_img_height)

    tb = slide.shapes.add_textbox(x_location, ppt_img_height, ppt_img_width, Inches(1))
    tf = tb.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = file_name
    font = run.font
    font.name = '맑은 고딕'
    font.size = Pt(11)

    fill = tb.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(217, 217, 217)



ppt_img_width = Cm(7.1)
ppt_img_height = Cm(20)


prs = Presentation()
prs.slide_height = Inches(9)
prs.slide_width = Inches(16)

cm_to_pixel = 96
cut_height_cm = 17.3
sum_cut = 0


folder = 'target/@UHDC_Part5_PC_MW_상담신청_0831'
folder_path = os.listdir(folder)
img_y = 0

for i, file in enumerate(folder_path):
    image_path = os.path.join(folder, file)
    img = Image.open(image_path)

    width_px, height_px = img.size

    # cm -> px
    width_new_px = 7 * cm_to_pixel
    height_new_px = height_px * width_new_px / width_px

    # image resizing
    img_new = img.resize((int(width_new_px), int(height_new_px)))

    # px -> cm
    cut_height_px = cut_height_cm * cm_to_pixel

    # cutting size
    cut_cnt = int((height_new_px / cut_height_px) + 1)

    for cut in range(cut_cnt):
        top = cut * cut_height_px
        bottom = cut_height_px * (cut + 1)
        cut_size = (0, top, width_new_px, bottom)
        img_crop = img_new.crop(cut_size)
        img_crop.save('tmp.png')

        if cut == 0:
            file_name = file.split('.png')[0]

        else:
            file_name = ''

        sum_cut += 1
        img_x = x_point(sum_cut)

        if sum_cut % 4 == 1:
            slide_layout = prs.slide_layouts[6]
            slide = prs.slides.add_slide(slide_layout)

        make_slide(img_x)

    crop_folder = folder.split('/')[1]
    prs.save(f"{crop_folder}.pptx")

