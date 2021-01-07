from pptx import Presentation
import os



def extract_text(file_path):
    prs = Presentation(file_path)
    text_list = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, 'text'):
                text_list.append(shape.text)
    return text_list


output = extract_text('input_files/3M.pptx')
print(output)