from pptx import Presentation

def extract_text_from_shape(shape, title_text=None):
    """
    Extracts text from a shape's text frame paragraphs and runs, excluding the title text.
    """
    text = ""
    if shape.has_text_frame and not shape.has_table:
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                if run.text not in title_text:
                    text += run.text + " "
    return text.strip()



def extract_content_from_slide(slide):
    """
    Extracts the title and content from a slide.
    """
    title = slide.shapes.title.text if slide.shapes.title else "NONE"
    content = ""

    for shape in slide.shapes:
        if shape.shape_type == 13:  # Check if the shape is an image (shape_type=13)
            content += "[IMAGE] "
        else:
            content += extract_text_from_shape(shape, title)+" "

    return title, content



def process_presentation(presentation_path):
    """
    Processes a PowerPoint presentation and returns a dictionary with slide numbers as keys and tuples of title and content as values.
    """
    presentation = Presentation(presentation_path)
    slide_data = {}

    for slide_num, slide in enumerate(presentation.slides, start=1):
        title, content = extract_content_from_slide(slide)
        slide_data[slide_num] = (title, content)

    return slide_data

def print_slide_content(slide_data):
    """
    Prints the slide number, title, and content of slides.
    """
    for slide_num, (title, content) in slide_data.items():
        print(f"Slide {slide_num}")
        print(f"Title: {title}")
        print(f"Content: {content}")
        print()

# Usage example
presentation_path = "C:\machine learning.pptx"
slides_data = process_presentation(presentation_path)
print_slide_content(slides_data)
