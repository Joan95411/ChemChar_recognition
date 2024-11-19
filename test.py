import xlrd
import Levenshtein
import unicodedata
import pandas as pd
import xlwings as xw
import math
import numpy as np
import difflib

def calculate_wer(ground_truth, predicted):
    # Split the strings into words
    ground_truth_words = ground_truth.split()
    predicted_words = predicted.split()

    # Calculate the Levenshtein distance between the words
    distance = Levenshtein.distance(ground_truth_words, predicted_words)
    print(ground_truth_words,predicted_words)
    print(distance,len(ground_truth_words))
    # Calculate the Word Error Rate,
    wer = distance / len(ground_truth_words)

    return wer
def calculate_cer(ground_truth, predicted):
    # Calculate the Levenshtein distance between the characters
    distance = Levenshtein.distance(ground_truth, predicted)
    print(ground_truth,predicted)
    print(distance,len(ground_truth))
    # Calculate the Character Error Rate
    cer = distance / len(ground_truth)

    return cer
def calscore(excel_file_path):

    workbook = xw.Book(excel_file_path)
    sheet = workbook.sheets[0]
    df = pd.read_excel(excel_file_path, sheet_name=sheet.name)

    # Iterate over the rows in the DataFrame
    for index, row in df.iterrows():
        image_filename = row['Filename']
        detected_text = row['Detected_text']
        detected_text_glue = row['Detected_text']
        ground_truth = row['Groundtruth 1']
        if image_filename.endswith('_0.jpg'):
            continue
        if not isinstance(ground_truth, str):
            if math.isnan(ground_truth):
                continue
        if not isinstance(detected_text, str):
            if math.isnan(detected_text):
                detected_text=''
                print(detected_text)

        normalized_detected_text = unicodedata.normalize("NFKC", str(detected_text))
        normalized_ground_truth = unicodedata.normalize("NFKC", str(ground_truth))
        print(normalized_ground_truth,normalized_detected_text)
        wer=calculate_wer(normalized_ground_truth,normalized_detected_text)
        wersub = calculate_wer(str(ground_truth), str(detected_text))
        print('wer:'+str(wer))
        print('wersub:' + str(wersub))
        cer=calculate_cer(normalized_ground_truth,normalized_detected_text)
        cersub = calculate_cer(str(ground_truth), str(detected_text))
        print('cer:' + str(cer))
        print('cersub:' + str(cersub))
        score=Levenshtein.ratio(normalized_ground_truth,normalized_detected_text)
        print('levenshtein ratio:' + str(score))


def detect_subscript(word):
    for symbol in word:
        if is_subscript(symbol):
            print(f"{symbol} is a subscript")
        else:
            print(f"{symbol} is not a subscript")
def is_subscript(symbol):
    for char in symbol:
        if "\u2080" <= char <= "\u2089":
            return True
    return False


def subscriptfinder():
    workbook = xlrd.open_workbook('Page1/Dataset_table_subscripts_page1.xlsx')

    # Select the worksheet you want to read
    worksheet = workbook.sheet_by_index(3)

    # Find the index of the column with the name 'detected text'
    detected_text_col_index = None
    for col in range(worksheet.ncols):
        cell = worksheet.cell(0, col)  # Assuming the column name is in the first row (row 0)
        if cell.value == 'Detected_text':
            detected_text_col_index = col
            break  # Exit the loop once found

    # If 'detected text' column was found, read that column
    if detected_text_col_index is not None:
        for row in range(1, worksheet.nrows):  # Start from row 1 to skip the header row
            cell = worksheet.cell(row, detected_text_col_index)
            value = cell.value

            xf_index = cell.xf_index

            # Check if the cell has a valid style
            if xf_index is not None:
                cell_format = workbook.xf_list[xf_index]
                font = workbook.font_list[cell_format.font_index]

                if font.escapement == 1:
                    subscript = True
                else:
                    subscript = False

                print(f"Value: {value}, Subscript: {subscript}")
            
    else:
        print("Column 'detected text' not found in the worksheet.")

# Call the function to execute
#subscriptfinder()


def calnumrank(excel_file_path):
    #calculate the most represented ground truth label
    workbook = xw.Book(excel_file_path)
    sheet = workbook.sheets[0]

    df = pd.read_excel(excel_file_path, sheet_name=sheet.name)
    pic1 = {}
    pic2 = {}
    pic3 = {}
    pic4 = {}
    pic5 = {}
    pic6 = {}
    pic7 = {}
    pic8 = {}

    # Initialize dictionaries with ground truth as keys and 0 as values
    ground_truths = set(df['Groundtruth 1'])
    for ground_truth in ground_truths:
        pic1[ground_truth] = 0
        pic2[ground_truth] = 0
        pic3[ground_truth] = 0
        pic4[ground_truth] = 0
        pic5[ground_truth] = 0
        pic6[ground_truth] = 0
        pic7[ground_truth] = 0
        pic8[ground_truth] = 0

    # Iterate over the rows in the DataFrame
    for index, row in df.iterrows():
        image_filename = row['Filename']
        ground_truth = row['Groundtruth 1']
        if image_filename.endswith('_0.jpg'):
            continue

        if not isinstance(ground_truth, str):
            if math.isnan(ground_truth):
                continue

        if image_filename.endswith('_1.jpeg'):
            pic1[ground_truth] += 1
        if image_filename.endswith('_2.jpeg'):
            pic2[ground_truth] += 1
        if image_filename.endswith('_3.jpeg'):
            pic3[ground_truth] += 1
        if image_filename.endswith('_4.jpeg'):
            pic4[ground_truth] += 1
        if image_filename.endswith('_5.jpeg'):
            pic5[ground_truth] += 1
        if image_filename.endswith('_6.jpeg'):
            pic6[ground_truth] += 1
        if image_filename.endswith('_7.jpeg'):
            pic7[ground_truth] += 1
        if image_filename.endswith('_8.jpeg'):
            pic8[ground_truth] += 1

    # Find the maximum count for each dictionary
    max_counts = {
        'pic1': max(pic1, key=pic1.get),
        'pic2': max(pic2, key=pic2.get),
        'pic3': max(pic3, key=pic3.get),
        'pic4': max(pic4, key=pic4.get),
        'pic5': max(pic5, key=pic5.get),
        'pic6': max(pic6, key=pic6.get),
        'pic7': max(pic7, key=pic7.get),
        'pic8': max(pic8, key=pic8.get)
    }

    print(max_counts)
    workbook.close()

def test():
    binary_list = [True, False, True, True, False, False, True, False, True]

    true_count = sum(binary_list)  # Count the number of True values
    total_count = len(binary_list)  # Get the total number of values in the list

    percentage_true = (true_count / total_count) * 100

    print(f"The percentage of True values: {percentage_true:.2f}%")
    print(f"{np.mean(binary_list)*100:.2f}%")

# missing_characters = [char[2] for char in difflib.ndiff('C2H6', 'C ₂ H 6')
#                                       if char.startswith('-')]
# print(missing_characters)