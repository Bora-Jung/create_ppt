"""
"""
import os
from PIL import Image

# 단위
CM_TO_PIXEL = 37.795275591

# 대상 높이 및 margin
CUT_HEIGHT_CM = 15
MARGIN_HEIGHT_CM = 1.5

def cropping(image_path: str, destination_dir: str = '.'):
    # open image
    origin_image = Image.open(image_path)

    # get size 원래 사이즈
    width_pixel, height_pixel = origin_image.size

########################################################################
    # pixel -> cm(width 7cm 에 맞춰서 수정)
#    width_cm = 26 * CM_TO_PIXEL
   # width_cm = 17 * CM_TO_PIXEL
#    height_cm = height_pixel * width_cm / width_pixel

    # resize image (ANTIALIAS : 높은 해상도의 사진 또는 영상을 낮은 해상도로 변환하거나 나타낼때 깨진 패턴의 형태로 나타나게 되는데 이를 최소화 시켜주는 방법
#    resized_image = origin_image.resize((int(width_cm), int(height_cm)), Image.ANTIALIAS)

    # get resized size
#    _, new_height_pixel = resized_image.size
########################################################################
 
 
    # cm -> pixel
    cut_height_pixel = width_pixel * 0.65 # CUT_HEIGHT_CM * CM_TO_PIXEL
    margin_height_pixel = MARGIN_HEIGHT_CM * CM_TO_PIXEL

    # crop 갯수
    #crop_ct = int(new_height_pixel // cut_height_pixel) + 1
    crop_ct = int(height_pixel // cut_height_pixel) + 1

    # cropped_dir
    image_fname, _ = os.path.splitext(os.path.basename(image_path))

    dest_directory = f'{destination_dir}/{image_fname}'
    os.makedirs(dest_directory, exist_ok=True)

    for i in range(crop_ct):
        top = max(0, i * cut_height_pixel - margin_height_pixel)
        #bottom = min(new_height_pixel, (i + 1) * cut_height_pixel + margin_height_pixel)
        bottom = min(height_pixel, (i + 1) * cut_height_pixel + margin_height_pixel)
        #cropped_image = resized_image.crop((0, top, width_cm, bottom))
        cropped_image = origin_image.crop((0, top, width_pixel, bottom))
        cropped_image.save(f'{dest_directory}/{i}.png', quality=100)


def path_walk(target_path: str, destination_dir: str = './cropped'):
    # target_path 내에 있는 png 파일명 모두 get
    image_paths = [f'{target_path}/{fname}'
                    for fname in os.listdir(target_path)
                    if os.path.basename(fname).endswith('.png')]

    # cropping
    for image_path in image_paths:
        cropping(image_path=image_path,
                 destination_dir=destination_dir)


if __name__ == '__main__':
    import argparse

    description = f"""
    잘하자~
    """

    # make parser
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument('--target', '-t',
                        type=str,
                        help='cropping할 이미지들이 있는 폴더 path',
                        required=True)
    parser.add_argument('--dest', '-d',
                        type=str,
                        default='./cropped',
                        help='cropping 한 뒤 저장할 폴더(없으면 생성)')

    # parse
    args = parser.parse_args()

    # get
    target_path = args.target.replace('/', os.path.sep)
    destination_dir = args.dest

    # run
    path_walk(target_path, destination_dir)

