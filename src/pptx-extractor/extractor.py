import os
import re

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


def write_image(shape, slide_id, shape_no, deck_name):
    image = shape.image
    image_bytes = image.blob
    image_filename = f'{slide_id}{shape_no}.{image.ext}'
    with open(f'images/{deck_name}/{image_filename}', 'wb') as f:
        f.write(image_bytes)


def extract_text(file_path):
    prs = Presentation(file_path)
    deck_name = re.findall(r'input_files/(.*?).pptx', file_path)[0] 
    os.mkdir(f'images/{deck_name}')
    text_list = []
    for slide in prs.slides:
        slide_id = slide.slide_id
        slide_list = []
        for shape in slide.shapes:
            shape_no = 0
            if hasattr(shape, 'text'):
                if shape.text != '':
                    slide_list.append(shape.text)
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                write_image(shape, slide_id, shape_no, deck_name)
            shape_no += 1
        text_list.append(slide_list)

    return text_list


output = extract_text('input_files/3M.pptx')
print(output)