import os
import re

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


class Extractor:

    def __init__(self, file_path):
        self.file_path = file_path
        self.prs = Presentation(file_path)
        self.deck_name = self.set_deck_name()

    def set_deck_name(self):
        return re.findall(r'input_files/(.*?).pptx', self.file_path)[0]

    def write_image(self, shape, slide_id, shape_no):
        image = shape.image
        image_bytes = image.blob
        image_filename = f'{slide_id}{shape_no}.{image.ext}'
        with open(f'images/{self.deck_name}/{image_filename}', 'wb') as f:
            f.write(image_bytes)

    def create_img_dir(self):
        os.mkdir(f'images/{self.deck_name}')

    def extract_text_and_img(self):
        self.create_img_dir()
        text_list = []
        for slide in self.prs.slides:
            slide_id = slide.slide_id
            slide_list = []
            for shape in slide.shapes:
                shape_no = 0
                if hasattr(shape, 'text'):
                    if shape.text != '':
                        slide_list.append(shape.text)
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    self.write_image(shape, slide_id, shape_no)
                shape_no += 1
            text_list.append(slide_list)

        return text_list


output = Extractor('input_files/3M.pptx').extract_text_and_img()
print(output)