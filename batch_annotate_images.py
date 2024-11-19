from enum import Enum
from google.cloud import vision
from PIL import Image, ImageDraw

word_to_ignore=['x','X','Koolstofdioxide','Fosfortrichloride','Ethanol','Difosforpentaoxide','Zwaveltrioxide','Water','Ammoniak','Distikstoftetraoxide','拜耳环境科学']
class FeatureType(Enum):
    PAGE = 1
    BLOCK = 2
    PARA = 3
    WORD = 4
    SYMBOL = 5


def expand_box(bounding_box, expand_factor):
    """Expand the bounding box by adding an expand_factor."""
    x_min = bounding_box.vertices[0].x - expand_factor
    y_min = bounding_box.vertices[0].y - expand_factor
    x_max = bounding_box.vertices[2].x + expand_factor
    y_max = bounding_box.vertices[2].y + expand_factor

    expanded_box = vision.BoundingPoly(
        vertices=[
            vision.Vertex(x=x_min, y=y_min),
            vision.Vertex(x=x_max, y=y_min),
            vision.Vertex(x=x_max, y=y_max),
            vision.Vertex(x=x_min, y=y_max)
        ]
    )

    return expanded_box



def draw_boxes(image, bounds, color, expand_factor):
    """Draw a border around the image using the hints in the vector list."""
    draw = ImageDraw.Draw(image)

    for bound in bounds:
        # Calculate the expanded coordinates
        x_min = bound.vertices[0].x - expand_factor
        y_min = bound.vertices[0].y - expand_factor
        x_max = bound.vertices[2].x + expand_factor
        y_max = bound.vertices[2].y + expand_factor

        draw.rectangle([(x_min, y_min), (x_max, y_max)], outline=color)

    return image


# [START vision_document_text_tutorial_detect_bounds]
def get_document_bounds(image_file):
    """Returns document bounds given an image."""
    client = vision.ImageAnnotatorClient()
    data=[]
    bounds = []
    with open(image_file, "rb") as image_file:
        content = image_file.read()

    image = vision.Image(content=content)

    response = client.document_text_detection(image=image)
    document = response.full_text_annotation

    # Collect specified feature bounds by enumerating all document features
    for page in document.pages:
        for block in page.blocks:
            for paragraph in block.paragraphs:
                paragr = []
                for word in paragraph.words:
                    word_text = ''.join([
                        symbol.text for symbol in word.symbols
                    ])
                    if word_text in word_to_ignore:
                        continue
                    if len(word_text)<=1:
                        paragr.append(word_text)
                        continue

                    print('Word text: {} (confidence: {})'.format(
                        word_text, word.confidence))
                    if True:
                        bounds.append(word.bounding_box)
                        data.append({'Text': word_text, 'Bounds': word.bounding_box})
                if len(paragr)>0:
                    para_text=''.join([
                        wo for wo in paragr
                    ])
                    print('Paragraph text: {} (confidence: {})'.format(
                    para_text, paragraph.confidence))
                    if True:
                        bounds.append(paragraph.bounding_box)
                        data.append({'Text': para_text, 'Bounds': paragraph.bounding_box})

    return data



def render_doc_text(filein, fileout):
    image = Image.open(filein)
    data = get_document_bounds(filein)
    for item in data:
        text = item['Text']
        bounding_box = item['Bounds']
        print(bounding_box)
        expanded_box = expand_box(bounding_box, expand_factor=35)  # Adjust expand_factor as needed

        # Calculate the expanded coordinates
        x_min = expanded_box.vertices[0].x
        y_min = expanded_box.vertices[0].y
        x_max = expanded_box.vertices[2].x
        y_max = expanded_box.vertices[2].y

        # Crop the image using the expanded coordinates
        cropped_image = image.crop((x_min, y_min, x_max, y_max))

        # Save the cropped image with the corresponding text as the filename
        filename = fileout + text + ".jpg"
        cropped_image.save(filename)


    if fileout != 0:
        image.save(fileout)
    else:
        image.show()





if __name__ == "__main__":
    detect_file = "img/t0.jpg"  # Replace with the path to your image file
    out_file = "img/t0/"  # Replace with the desired output file path (optional)

    render_doc_text(detect_file, out_file)
