import math
import os
import pandas as pd
import xlwings as xw
from google.cloud import vision
import time
import re
import unicodedata
import numpy as np
# Set the path to your service account key JSON file
service_account_key_path = "Google_service_key.json"

# Set the environment variable to point to the key file
os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = service_account_key_path
def Google_detect_document_from_excel(excel_file_path):
    """Detects document features in images from an Excel file."""
    client = vision.ImageAnnotatorClient()

    # Get the directory path of the Excel file
    directory = os.path.dirname(excel_file_path)

    workbook = xw.Book(excel_file_path)
    sheet = workbook.sheets[0]
    df = pd.read_excel(excel_file_path, sheet_name=sheet.name)
    # Get the starting column index for the new column
    start_column_index = sheet.used_range.last_cell.column + 1

    # Iterate over the rows in the DataFrame
    for index, row in df.iterrows():
        image_filename = row['Filename']
        ground_truth = row['Groundtruth 1']

        # Construct the absolute path of the image
        image_path = os.path.join(directory, image_filename)

        # Read the image file
        with open(image_path, 'rb') as image_file:
            content = image_file.read()

        image = vision.Image(content=content)

        response = client.document_text_detection(image=image, image_context={"language_hints": ["en"]})

        # Extract the detected words from the response
        detected_words = []
        word_confidences = []
        for page in response.full_text_annotation.pages:
            for block in page.blocks:
                for paragraph in block.paragraphs:
                    for word in paragraph.words:
                        word_text = ''.join([
                            symbol.text for symbol in word.symbols
                        ])
                        detected_words.append(word_text)
                        word_confidences.append(word.confidence)

        # Join the detected words into a single string
        detected_text = ' '.join(detected_words)
        normalized_detected_text = unicodedata.normalize("NFKC", detected_text)
        # Normalize the strings using Unicode normalization form NFKC

        if image_filename.endswith('_0.jpg'):
            page_text=detected_text
            page_num=re.search(r'(.*?)_0\.jpg', image_filename).group(1)
            normalized_page_text = unicodedata.normalize("NFKC", page_text)
        if image_filename.startswith(page_num) and not image_filename.endswith('_0.jpg'):
            if not isinstance(ground_truth, str):
                if math.isnan(ground_truth):
                    continue
            normalized_ground_truth = unicodedata.normalize("NFKC", str(ground_truth))
            if normalized_ground_truth in normalized_page_text:
                Contained_big=True
            else:
                Contained_big=False
            sheet.range((index + 2, start_column_index + 1)).value = Contained_big
            if normalized_ground_truth in normalized_detected_text:
                Contained_small = True
            else:
                Contained_small = False
            sheet.range((index + 2, start_column_index + 2)).value = Contained_small
        try:
            sheet.range((index + 2, start_column_index)).number_format = '@'
            sheet.range((index + 2, start_column_index)).value = str(detected_text)
        except Exception as e:
            print("Error occurred:", e)
            print("Detected Text:", detected_text)
        if word_confidences:
            mean_confidence = np.mean(word_confidences)
        else:
            mean_confidence = 0  # Default value if word_confidences is empty
        sheet.range((index + 2, start_column_index + 3)).value = mean_confidence
    sheet.range((1, start_column_index)).value = 'Detected_text'
    sheet.range((1, start_column_index + 1)).value = 'Contained in big picture'
    sheet.range((1, start_column_index + 2)).value = 'Contained in small picture'
    sheet.range((1, start_column_index + 3)).value = 'Word confidence'
    workbook.save(excel_file_path)
    workbook.close()

def Google_detect_document_sub(excel_file_path):
    """Detects document features in images from an Excel file."""
    client = vision.ImageAnnotatorClient()

    # Get the directory path of the Excel file
    directory = os.path.dirname(excel_file_path)

    workbook = xw.Book(excel_file_path)
    sheet = workbook.sheets[0]
    df = pd.read_excel(excel_file_path, sheet_name=sheet.name)
    # Get the starting column index for the new column
    start_column_index = sheet.used_range.last_cell.column + 1

    # Iterate over the rows in the DataFrame
    for index, row in df.iterrows():
        image_filename = row['Filename']
        ground_truth = str(row['Groundtruth 1'])

        # Construct the absolute path of the image
        image_path = os.path.join(directory, image_filename)

        # Read the image file
        with open(image_path, 'rb') as image_file:
            content = image_file.read()

        image = vision.Image(content=content)

        response = client.document_text_detection(image=image, image_context={"language_hints": ["en"]})

        # Extract the detected words from the response
        detected_words = []
        word_confidences = []
        for page in response.full_text_annotation.pages:
            for block in page.blocks:
                for paragraph in block.paragraphs:
                    for word in paragraph.words:
                        word_text = ''.join([
                            symbol.text for symbol in word.symbols
                        ])
                        detected_words.append(word_text)
                        word_confidences.append(word.confidence)

        # Join the detected words into a single string
        detected_text = ' '.join(detected_words)
        #normalized_detected_text = unicodedata.normalize("NFKC", detected_text)
        if image_filename.endswith('_0.jpg'):
            page_text=detected_text
            page_num=re.search(r'(.*?)_0\.jpg', image_filename).group(1)
        if image_filename.startswith(page_num) and not image_filename.endswith('_0.jpg'):
            if not isinstance(ground_truth, str):
                if math.isnan(ground_truth):
                    continue
            if ground_truth in page_text:
                Contained_big=True
            else:
                Contained_big=False
            sheet.range((index + 2, start_column_index + 1)).value = Contained_big
            if ground_truth in detected_text:
                Contained_small = True
            else:
                Contained_small = False
            sheet.range((index + 2, start_column_index + 2)).value = Contained_small
        try:
            sheet.range((index + 2, start_column_index)).number_format = '@'
            sheet.range((index + 2, start_column_index)).value = str(detected_text)
        except Exception as e:
            print("Error occurred:", e)
            print("Detected Text:", detected_text)
        if word_confidences:
            mean_confidence = np.mean(word_confidences)
        else:
            mean_confidence = 0  # Default value if word_confidences is empty
        sheet.range((index + 2, start_column_index + 3)).value = mean_confidence
    sheet.range((1, start_column_index)).value = 'Detected_text'
    sheet.range((1, start_column_index + 1)).value = 'Contained in big picture'
    sheet.range((1, start_column_index + 2)).value = 'Contained in small picture'
    sheet.range((1, start_column_index + 3)).value = 'Word confidence'
    workbook.save(excel_file_path)
    workbook.close()

def Google_detect_document_glue(excel_file_path):
    """Detects document features in images from an Excel file."""
    client = vision.ImageAnnotatorClient()

    # Get the directory path of the Excel file
    directory = os.path.dirname(excel_file_path)

    workbook = xw.Book(excel_file_path)
    sheet = workbook.sheets[0]
    df = pd.read_excel(excel_file_path, sheet_name=sheet.name)
    # Get the starting column index for the new column
    start_column_index = sheet.used_range.last_cell.column + 1

    # Iterate over the rows in the DataFrame
    for index, row in df.iterrows():
        image_filename = row['Filename']
        ground_truth = row['Groundtruth 1']

        # Construct the absolute path of the image
        image_path = os.path.join(directory, image_filename)

        # Read the image file
        with open(image_path, 'rb') as image_file:
            content = image_file.read()

        image = vision.Image(content=content)

        response = client.document_text_detection(image=image, image_context={"language_hints": ["en"]})

        # Extract the detected words from the response
        detected_words = []
        word_confidences = []
        for page in response.full_text_annotation.pages:
            for block in page.blocks:
                for paragraph in block.paragraphs:
                    for word in paragraph.words:
                        word_text = ''.join([
                            symbol.text for symbol in word.symbols
                        ])
                        detected_words.append(word_text)
                        word_confidences.append(word.confidence)

        # Join the detected words into a single string
        detected_text = ''.join(detected_words)
        normalized_detected_text = unicodedata.normalize("NFKC", detected_text)
        # Normalize the strings using Unicode normalization form NFKC

        if image_filename.endswith('_0.jpg'):
            page_text=detected_text
            page_num=re.search(r'(.*?)_0\.jpg', image_filename).group(1)
            normalized_page_text = unicodedata.normalize("NFKC", page_text)
        if image_filename.startswith(page_num) and not image_filename.endswith('_0.jpg'):
            if not isinstance(ground_truth, str):
                if math.isnan(ground_truth):
                    continue
            normalized_ground_truth = unicodedata.normalize("NFKC", str(ground_truth))
            if normalized_ground_truth in normalized_page_text:
                Contained_big=True
            else:
                Contained_big=False
            sheet.range((index + 2, start_column_index + 1)).value = Contained_big
            if normalized_ground_truth in normalized_detected_text:
                Contained_small = True
            else:
                Contained_small = False
            sheet.range((index + 2, start_column_index + 2)).value = Contained_small
        try:
            sheet.range((index + 2, start_column_index)).number_format = '@'
            sheet.range((index + 2, start_column_index)).value = str(detected_text)
        except Exception as e:
            print("Error occurred:", e)
            print("Detected Text:", detected_text)

    sheet.range((1, start_column_index)).value = 'Detected_text_glue'
    sheet.range((1, start_column_index + 1)).value = 'Contained in big picture glue'
    sheet.range((1, start_column_index + 2)).value = 'Contained in small picture glue'
    workbook.save(excel_file_path)
    workbook.close()


start_time = time.time()
Google_detect_document_glue('Page4/Dataset_table_page4.xlsx')
end_time = time.time()
execution_time = end_time - start_time

print(f"The program ran for {execution_time} seconds.")