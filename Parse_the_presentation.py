from pptx import Presentation

presentation = Presentation("C:\machine learning.pptx")
title_and_content = {}

for slide in presentation.slides:
    title = slide.shapes.title.text if slide.shapes.title else "NONE"
    content = ""

    for shape in slide.shapes:
        if shape.has_text_frame and not shape.has_table:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    content += run.text
        elif shape.shape_type == 13:  # Check if the shape is an image (shape_type=13)
            content += "[IMAGE]"

    title_and_content[title] = title_and_content.get(title, "") + content

for key, value in title_and_content.items():
    print(key)
    print(value)
    print()
