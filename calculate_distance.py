import Levenshtein
import unicodedata
import pandas as pd
import xlwings as xw
import math
import difflib
import re
import numpy as np
def calculate_wer(ground_truth, predicted):
    # Split the strings into words
    ground_truth_words = ground_truth.split()
    predicted_words = predicted.split()
    print(ground_truth_words,predicted_words)
    # Calculate the Levenshtein distance between the words
    distance = Levenshtein.distance(ground_truth_words, predicted_words)

    # Calculate the Word Error Rate
    wer = distance / len(ground_truth_words)

    return wer

def calculate_cer(ground_truth, predicted):

    distance = Levenshtein.distance(ground_truth, predicted)

    # Calculate the Character Error Rate
    cer = distance / len(ground_truth)

    return cer

def calculate_cer_confi(ground_truth, predicted, confidence_scores):
    total_weighted_distance = 0
    total_weight = 0

    for i in range(len(ground_truth)):
        distance = Levenshtein.distance(ground_truth[i], predicted[i])
        confidence_score = confidence_scores[i]

        # Calculate weighted distance based on confidence score
        weighted_distance = distance * (1 - confidence_score)

        # Accumulate weighted distance and total weight
        total_weighted_distance += weighted_distance
        total_weight += (1 - confidence_score)

    # Calculate Character Error Rate with weighted distance
    cer_confi = total_weighted_distance / total_weight

    return cer_confi

def calscore(excel_file_path):

    workbook = xw.Book(excel_file_path)
    for sheet in workbook.sheets:
        if sheet.name in ['Google', 'Azure']:
            df = pd.read_excel(excel_file_path, sheet_name=sheet.name)
            start_column_index = sheet.used_range.last_cell.column + 1

            # Iterate over the rows in the DataFrame
            for index, row in df.iterrows():
                image_filename = row['Filename']
                detected_text = row['Detected_text']
                ground_truth = row['Groundtruth 1']
                if image_filename.endswith('_0.jpg'):
                    continue
                if not isinstance(ground_truth, str):
                    if math.isnan(ground_truth):
                        continue
                if not isinstance(detected_text, str):
                    if math.isnan(detected_text):
                        detected_text=' '

                normalized_detected_text = unicodedata.normalize("NFKC", detected_text)
                normalized_ground_truth = unicodedata.normalize("NFKC", ground_truth)
                wer=calculate_wer(str(normalized_ground_truth),str(normalized_detected_text))
                cer=calculate_cer(str(normalized_ground_truth),str(normalized_detected_text))
                score=Levenshtein.ratio(normalized_ground_truth,normalized_detected_text)
                try:
                    sheet.range((index + 2, start_column_index)).value = wer
                    sheet.range((index + 2, start_column_index+1)).value = cer
                    sheet.range((index + 2, start_column_index + 2)).value = score
                except Exception as e:
                    print("Error occurred:", e)
                    print("Detected Text:", detected_text)
            sheet.range((1, start_column_index)).value = 'Word error rate'
            sheet.range((1, start_column_index + 1)).value = 'Character error rate'
            sheet.range((1, start_column_index + 2)).value = 'Levenshtein ratio'
    workbook.save(excel_file_path)
    workbook.close()
def calscore_norm_glue(excel_file_path):

    workbook = xw.Book(excel_file_path)
    for sheet in workbook.sheets:
        df = pd.read_excel(excel_file_path, sheet_name=sheet.name)
        start_column_index = sheet.used_range.last_cell.column + 1

        # Iterate over the rows in the DataFrame
        for index, row in df.iterrows():
            image_filename = row['Filename']
            detected_text = row['Detected_text_glue']
            ground_truth = row['Groundtruth 1']
            if image_filename.endswith('_0.jpg'):
                continue
            if not isinstance(ground_truth, str):
                if math.isnan(ground_truth):
                    continue
            if not isinstance(detected_text, str):
                if math.isnan(detected_text):
                    detected_text=''

            normalized_detected_text = unicodedata.normalize("NFKC", str(detected_text))
            normalized_ground_truth = unicodedata.normalize("NFKC", str(ground_truth))
            wer=calculate_wer(normalized_ground_truth,normalized_detected_text)
            cer=calculate_cer(normalized_ground_truth,normalized_detected_text)
            score=Levenshtein.ratio(normalized_ground_truth,normalized_detected_text)
            try:
                sheet.range((index + 2, start_column_index)).value = wer
                sheet.range((index + 2, start_column_index+1)).value = cer
                sheet.range((index + 2, start_column_index + 2)).value = score
            except Exception as e:
                print("Error occurred:", e)
                print("Detected Text:", detected_text)
        sheet.range((1, start_column_index)).value = 'Word error rate(G)'
        sheet.range((1, start_column_index + 1)).value = 'Character error rate(G)'
        sheet.range((1, start_column_index + 2)).value = 'Levenshtein ratio(G)'
    workbook.save(excel_file_path)
    workbook.close()
def calscore_sub(excel_file_path):

    workbook = xw.Book(excel_file_path)
    for sheet in workbook.sheets:
        df = pd.read_excel(excel_file_path, sheet_name=sheet.name)
        start_column_index = sheet.used_range.last_cell.column + 1

        # Iterate over the rows in the DataFrame
        for index, row in df.iterrows():
            image_filename = row['Filename']
            detected_text = row['Detected_text']
            ground_truth = row['Groundtruth 1']
            if image_filename.endswith('_0.jpg'):
                continue
            if not isinstance(ground_truth, str):
                if math.isnan(ground_truth):
                    continue
            if not isinstance(detected_text, str):
                if math.isnan(detected_text):
                    detected_text=' '

            wer=calculate_wer(str(ground_truth),str(detected_text))
            cer=calculate_cer(str(ground_truth),str(detected_text))
            score=Levenshtein.ratio(str(ground_truth),str(detected_text))
            try:
                sheet.range((index + 2, start_column_index)).value = wer
                sheet.range((index + 2, start_column_index+1)).value = cer
                sheet.range((index + 2, start_column_index + 2)).value = score
            except Exception as e:
                print("Error occurred:", e)
                print("Detected Text:", detected_text)
        sheet.range((1, start_column_index)).value = 'Word error rate'
        sheet.range((1, start_column_index + 1)).value = 'Character error rate'
        sheet.range((1, start_column_index + 2)).value = 'Levenshtein ratio'
    workbook.save(excel_file_path)
    workbook.close()
def calscore_glue(excel_file_path):

    workbook = xw.Book(excel_file_path)
    for sheet in workbook.sheets:
        df = pd.read_excel(excel_file_path, sheet_name=sheet.name)
        start_column_index = sheet.used_range.last_cell.column + 1

        # Iterate over the rows in the DataFrame
        for index, row in df.iterrows():
            image_filename = row['Filename']
            detected_text = row['Detected_text_glue']
            ground_truth = row['Groundtruth 1']
            if image_filename.endswith('_0.jpg'):
                continue
            if not isinstance(ground_truth, str):
                if math.isnan(ground_truth):
                    continue
            if not isinstance(detected_text, str):
                if math.isnan(detected_text):
                    detected_text=''

            wer=calculate_wer(str(ground_truth),str(detected_text))
            cer=calculate_cer(str(ground_truth),str(detected_text))
            score=Levenshtein.ratio(str(ground_truth),str(detected_text))
            try:
                sheet.range((index + 2, start_column_index)).value = wer
                sheet.range((index + 2, start_column_index+1)).value = cer
                sheet.range((index + 2, start_column_index + 2)).value = score
            except Exception as e:
                print("Error occurred:", e)
                print("Detected Text:", detected_text)
        sheet.range((1, start_column_index)).value = 'Word error rate(G)'
        sheet.range((1, start_column_index + 1)).value = 'Character error rate(G)'
        sheet.range((1, start_column_index + 2)).value = 'Levenshtein ratio(G)'
    workbook.save(excel_file_path)
    workbook.close()

def Azure_contained(excel_file_path):

    workbook = xw.Book(excel_file_path)
    sheet = workbook.sheets[1]
    df = pd.read_excel(excel_file_path,sheet_name=sheet.name)
    # Get the starting column index for the new column
    start_column_index = sheet.used_range.last_cell.column + 1
    # Iterate over the rows in the DataFrame
    for index, row in df.iterrows():
        image_filename = row['Filename']
        ground_truth = row['Groundtruth 1']
        detected_text =row['Detected_text_glue']

        if image_filename.endswith('_0.jpg'):
            page_text = detected_text
            page_num = re.search(r'(.*?)_0\.jpg', image_filename).group(1)
        if image_filename.startswith(page_num) and not image_filename.endswith('_0.jpg'):
            if not isinstance(ground_truth, str):
                if math.isnan(ground_truth):
                    continue
            if str(ground_truth) in str(page_text):
                Contained_big = True
            else:
                Contained_big = False
            sheet.range((index + 2, start_column_index )).value = Contained_big
            if str(ground_truth) in str(detected_text):
                Contained_small = True
            else:
                Contained_small = False
            sheet.range((index + 2, start_column_index + 1)).value = Contained_small

    sheet.range((1, start_column_index )).value = 'Contained in big picture glue'
    sheet.range((1, start_column_index + 1)).value = 'Contained in small picture glue'
    workbook.save(excel_file_path)
    workbook.close()
def Google_contained(excel_file_path):

    workbook = xw.Book(excel_file_path)
    sheet = workbook.sheets[0]
    df = pd.read_excel(excel_file_path,sheet_name=sheet.name)
    # Get the starting column index for the new column
    start_column_index = sheet.used_range.last_cell.column + 1
    # Iterate over the rows in the DataFrame
    for index, row in df.iterrows():
        image_filename = row['Filename']
        ground_truth = row['Groundtruth 1']
        detected_text =row['Detected_text']

        if image_filename.endswith('_0.jpg'):
            page_text = detected_text
            page_num = re.search(r'(.*?)_0\.jpg', image_filename).group(1)
        if image_filename.startswith(page_num) and not image_filename.endswith('_0.jpg'):
            if not isinstance(ground_truth, str):
                if math.isnan(ground_truth):
                    continue
            if str(ground_truth) in str(page_text):
                Contained_big = True
            else:
                Contained_big = False
            sheet.range((index + 2, start_column_index )).value = Contained_big
            if str(ground_truth) in str(detected_text):
                Contained_small = True
            else:
                Contained_small = False
            sheet.range((index + 2, start_column_index + 1)).value = Contained_small

    sheet.range((1, start_column_index)).value = 'Contained in big picture'
    sheet.range((1, start_column_index + 1)).value = 'Contained in small picture'
    workbook.save(excel_file_path)
    workbook.close()

def Google_contained_glue(excel_file_path):

    workbook = xw.Book(excel_file_path)
    sheet = workbook.sheets[0]
    df = pd.read_excel(excel_file_path,sheet_name=sheet.name)
    # Get the starting column index for the new column
    start_column_index = sheet.used_range.last_cell.column + 1
    # Iterate over the rows in the DataFrame
    for index, row in df.iterrows():
        image_filename = row['Filename']
        ground_truth = row['Groundtruth 1']
        detected_text =row['Detected_text_glue']

        if image_filename.endswith('_0.jpg'):
            page_text = detected_text
            page_num = re.search(r'(.*?)_0\.jpg', image_filename).group(1)
        if image_filename.startswith(page_num) and not image_filename.endswith('_0.jpg'):
            if not isinstance(ground_truth, str):
                if math.isnan(ground_truth):
                    continue
            if str(ground_truth) in str(page_text):
                Contained_big = True
            else:
                Contained_big = False
            sheet.range((index + 2, start_column_index )).value = Contained_big
            if str(ground_truth) in str(detected_text):
                Contained_small = True
            else:
                Contained_small = False
            sheet.range((index + 2, start_column_index + 1)).value = Contained_small

    sheet.range((1, start_column_index)).value = 'Contained in big picture glue'
    sheet.range((1, start_column_index + 1)).value = 'Contained in small picture glue'
    workbook.save(excel_file_path)
    workbook.close()

def calconfi(excel_file_path):

    workbook = xw.Book(excel_file_path)
    for sheet in workbook.sheets:
        df = pd.read_excel(excel_file_path, sheet_name=sheet.name)

        pic1_counts = 0
        pic2_counts = 0
        pic3_counts = 0
        pic4_counts = 0
        pic5_counts = 0
        pic6_counts = 0
        pic7_counts = 0
        pic8_counts = 0
        wer_all=[]
        pic1_wer = []
        pic2_wer = []
        pic3_wer = []
        pic4_wer = []
        pic5_wer = []
        pic6_wer = []
        pic7_wer = []
        pic8_wer = []
        cer_all=[]
        pic1_cer = []
        pic2_cer = []
        pic3_cer = []
        pic4_cer = []
        pic5_cer = []
        pic6_cer = []
        pic7_cer = []
        pic8_cer = []
        pic1_lev = []
        pic2_lev = []
        pic3_lev = []
        pic4_lev = []
        pic5_lev = []
        pic6_lev = []
        pic7_lev = []
        pic8_lev = []
        lev_all = []
        pic1_all = 0
        pic2_all = 0
        pic3_all = 0
        pic4_all = 0
        pic5_all = 0
        pic6_all = 0
        pic7_all = 0
        pic8_all = 0
        pic1_confidence = []
        pic2_confidence = []
        pic3_confidence = []
        pic4_confidence = []
        pic5_confidence = []
        pic6_confidence = []
        pic7_confidence = []
        pic8_confidence = []
        confi_all=[]
        # Iterate over the rows in the DataFrame
        for index, row in df.iterrows():
            image_filename = row['Filename']
            contained = row['Contained in small picture']
            confidence = row['Word confidence']
            wer=row['Word error rate']
            cer = row['Character error rate']
            lev=row['Levenshtein ratio']
            if image_filename.endswith('_0.jpg'):
                continue
            if not isinstance(contained, str):
                if math.isnan(contained):
                    continue
            confi_all.append(confidence)
            wer_all.append(wer)
            cer_all.append(cer)
            lev_all.append(lev)
            if image_filename.endswith('_1.jpeg'):
                pic1_all+=1
                pic1_confidence.append(confidence)
                pic1_counts += contained
                pic1_wer.append(wer)
                pic1_cer .append(cer)
                pic1_lev.append(lev)
            if image_filename.endswith('_2.jpeg'):
                pic2_all += 1
                pic2_confidence.append(confidence)
                pic2_counts += contained
                pic2_wer .append(wer)
                pic2_cer .append(cer)
                pic2_lev.append(lev)
            if image_filename.endswith('_3.jpeg'):
                pic3_all += 1
                pic3_confidence.append(confidence)
                pic3_counts += contained
                pic3_wer .append(wer)
                pic3_cer .append(cer)
                pic3_lev.append(lev)
            if image_filename.endswith('_4.jpeg'):
                pic4_all += 1
                pic4_confidence.append(confidence)
                pic4_counts += contained
                pic4_wer .append(wer)
                pic4_cer .append(cer)
                pic4_lev.append(lev)
            if image_filename.endswith('_5.jpeg'):
                pic5_all += 1
                pic5_confidence.append(confidence)
                pic5_counts += contained
                pic5_wer .append(wer)
                pic5_cer .append(cer)
                pic5_lev.append(lev)
            if image_filename.endswith('_6.jpeg'):
                pic6_all += 1
                pic6_confidence.append(confidence)
                pic6_counts += contained
                pic6_wer .append(wer)
                pic6_cer .append(cer)
                pic6_lev.append(lev)
            if image_filename.endswith('_7.jpeg'):
                pic7_all += 1
                pic7_confidence.append(confidence)
                pic7_counts += contained
                pic7_wer .append(wer)
                pic7_cer .append(cer)
                pic7_lev.append(lev)
            if image_filename.endswith('_8.jpeg'):
                pic8_all += 1
                pic8_confidence.append(confidence)
                pic8_counts += contained
                pic8_wer .append(wer)
                pic8_cer .append(cer)
                pic8_lev.append(lev)

        wer_mean = {
            'pic1': np.mean(pic1_wer),
            'pic2': np.mean(pic2_wer),
            'pic3': np.mean(pic3_wer),
            'pic4': np.mean(pic4_wer),
            'pic5': np.mean(pic5_wer),
            'pic6': np.mean(pic6_wer),
            'pic7': np.mean(pic7_wer),
            'pic8': np.mean(pic8_wer),
            'wer': np.mean(wer_all)
        }
        cer_mean = {
            'pic1': np.mean(pic1_cer),
            'pic2': np.mean(pic2_cer),
            'pic3': np.mean(pic3_cer),
            'pic4': np.mean(pic4_cer),
            'pic5': np.mean(pic5_cer),
            'pic6': np.mean(pic6_cer),
            'pic7': np.mean(pic7_cer),
            'pic8': np.mean(pic8_cer),
            'cer': np.mean(cer_all)
        }
        lev_mean = {
            'pic1': np.mean(pic1_lev),
            'pic2': np.mean(pic2_lev),
            'pic3': np.mean(pic3_lev),
            'pic4': np.mean(pic4_lev),
            'pic5': np.mean(pic5_lev),
            'pic6': np.mean(pic6_lev),
            'pic7': np.mean(pic7_lev),
            'pic8': np.mean(pic8_lev),
            'Levenshtein ratio': np.mean(lev_all)
        }
        true_percentages = {
            'pic1': '{:.2%}'.format(pic1_counts / pic1_all),
            'pic2': '{:.2%}'.format(pic2_counts / pic2_all),
            'pic3': '{:.2%}'.format(pic3_counts / pic3_all),
            'pic4': '{:.2%}'.format(pic4_counts / pic4_all),
            'pic5': '{:.2%}'.format(pic5_counts / pic5_all),
            'pic6': '{:.2%}'.format(pic6_counts / pic6_all),
            'pic7': '{:.2%}'.format(pic7_counts / pic7_all),
            'pic8': '{:.2%}'.format(pic8_counts / pic8_all)
        }
        mean_confidence = {
            'pic1': np.mean(pic1_confidence),
            'pic2': np.mean(pic2_confidence),
            'pic3': np.mean(pic3_confidence),
            'pic4': np.mean(pic4_confidence),
            'pic5': np.mean(pic5_confidence),
            'pic6': np.mean(pic6_confidence),
            'pic7': np.mean(pic7_confidence),
            'pic8': np.mean(pic8_confidence),
            'Confidence':np.mean(confi_all)
        }

        print(sheet.name,true_percentages)
        print(sheet.name,mean_confidence)
        print(sheet.name,wer_mean)
        print(sheet.name,cer_mean)
        print(sheet.name,lev_mean)
    workbook.close()

def calculate_everything(excel_file_path):

    workbook = xw.Book(excel_file_path)
    for sheet in workbook.sheets:
        if sheet.name in ['Google', 'Azure']:
            df = pd.read_excel(excel_file_path, sheet_name=sheet.name)

            data = {
                "pic1": {"cb": [],"cs": [], "confi":[],"wer": [],"cer":[],"lev":[],"cbg":[],"csg":[],"werg":[],"cerg":[],"levg":[]},
                "pic2": {"cb": [],"cs": [], "confi":[],"wer": [],"cer":[],"lev":[],"cbg":[],"csg":[],"werg":[],"cerg":[],"levg":[]},
                "pic3": {"cb": [],"cs": [], "confi":[],"wer": [],"cer":[],"lev":[],"cbg":[],"csg":[],"werg":[],"cerg":[],"levg":[]},
                "pic4": {"cb": [],"cs": [], "confi":[],"wer": [],"cer":[],"lev":[],"cbg":[],"csg":[],"werg":[],"cerg":[],"levg":[]},
                "pic5": {"cb": [],"cs": [], "confi":[],"wer": [],"cer":[],"lev":[],"cbg":[],"csg":[],"werg":[],"cerg":[],"levg":[]},
                "pic6": {"cb": [],"cs": [], "confi":[],"wer": [],"cer":[],"lev":[],"cbg":[],"csg":[],"werg":[],"cerg":[],"levg":[]},
                "pic7": {"cb": [],"cs": [], "confi":[],"wer": [],"cer":[],"lev":[],"cbg":[],"csg":[],"werg":[],"cerg":[],"levg":[]},
                "pic8": {"cb": [],"cs": [], "confi":[],"wer": [],"cer":[],"lev":[],"cbg":[],"csg":[],"werg":[],"cerg":[],"levg":[]}
            }

            #no need to calculate overall wer/cer/confidence mean which represents the average mean value of the page performance, but the page has 8 different writings,
            #it doesn't make sense to mix them up, we calculate the overall result for each answer, then organize the figures and compare
            # Iterate over the rows in the DataFrame
            for index, row in df.iterrows():
                image_filename = row['Filename']
                contained_big = row['Contained in big picture']
                contained_small = row['Contained in small picture']
                confidence = row['Word confidence']
                wer=row['Word error rate']
                cer = row['Character error rate']
                lev=row['Levenshtein ratio']
                contained_big_glue = row['Contained in big picture glue']
                contained_small_glue = row['Contained in small picture glue']
                werG = row['Word error rate(G)']
                cerG = row['Character error rate(G)']
                levG = row['Levenshtein ratio(G)']
                if image_filename.endswith('_0.jpg'):
                    continue
                if not isinstance(contained_small, str):
                    if math.isnan(contained_small):
                        continue

                if image_filename.endswith('_1.jpeg'):
                    data["pic1"]["cb"].append(contained_big)
                    data["pic1"]["cs"].append(contained_small)
                    data["pic1"]["confi"].append(confidence)
                    data["pic1"]["wer"].append(wer)
                    data["pic1"]["cer"] .append(cer)
                    data["pic1"]["lev"].append(lev)
                    data["pic1"]["cbg"].append(contained_big_glue)
                    data["pic1"]["csg"].append(contained_small_glue)
                    data["pic1"]["werg"].append(werG)
                    data["pic1"]["cerg"].append(cerG)
                    data["pic1"]["levg"].append(levG)
                if image_filename.endswith('_2.jpeg'):
                    data["pic2"]["cb"].append(contained_big)
                    data["pic2"]["cs"].append(contained_small)
                    data["pic2"]["confi"].append(confidence)
                    data["pic2"]["wer"].append(wer)
                    data["pic2"]["cer"].append(cer)
                    data["pic2"]["lev"].append(lev)
                    data["pic2"]["cbg"].append(contained_big_glue)
                    data["pic2"]["csg"].append(contained_small_glue)
                    data["pic2"]["werg"].append(werG)
                    data["pic2"]["cerg"].append(cerG)
                    data["pic2"]["levg"].append(levG)
                if image_filename.endswith('_3.jpeg'):
                    data["pic3"]["cb"].append(contained_big)
                    data["pic3"]["cs"].append(contained_small)
                    data["pic3"]["confi"].append(confidence)
                    data["pic3"]["wer"].append(wer)
                    data["pic3"]["cer"].append(cer)
                    data["pic3"]["lev"].append(lev)
                    data["pic3"]["cbg"].append(contained_big_glue)
                    data["pic3"]["csg"].append(contained_small_glue)
                    data["pic3"]["werg"].append(werG)
                    data["pic3"]["cerg"].append(cerG)
                    data["pic3"]["levg"].append(levG)
                if image_filename.endswith('_4.jpeg'):
                    data["pic4"]["cb"].append(contained_big)
                    data["pic4"]["cs"].append(contained_small)
                    data["pic4"]["confi"].append(confidence)
                    data["pic4"]["wer"].append(wer)
                    data["pic4"]["cer"].append(cer)
                    data["pic4"]["lev"].append(lev)
                    data["pic4"]["cbg"].append(contained_big_glue)
                    data["pic4"]["csg"].append(contained_small_glue)
                    data["pic4"]["werg"].append(werG)
                    data["pic4"]["cerg"].append(cerG)
                    data["pic4"]["levg"].append(levG)
                if image_filename.endswith('_5.jpeg'):
                    data["pic5"]["cb"].append(contained_big)
                    data["pic5"]["cs"].append(contained_small)
                    data["pic5"]["confi"].append(confidence)
                    data["pic5"]["wer"].append(wer)
                    data["pic5"]["cer"].append(cer)
                    data["pic5"]["lev"].append(lev)
                    data["pic5"]["cbg"].append(contained_big_glue)
                    data["pic5"]["csg"].append(contained_small_glue)
                    data["pic5"]["werg"].append(werG)
                    data["pic5"]["cerg"].append(cerG)
                    data["pic5"]["levg"].append(levG)
                if image_filename.endswith('_6.jpeg'):
                    data["pic6"]["cb"].append(contained_big)
                    data["pic6"]["cs"].append(contained_small)
                    data["pic6"]["confi"].append(confidence)
                    data["pic6"]["wer"].append(wer)
                    data["pic6"]["cer"].append(cer)
                    data["pic6"]["lev"].append(lev)
                    data["pic6"]["cbg"].append(contained_big_glue)
                    data["pic6"]["csg"].append(contained_small_glue)
                    data["pic6"]["werg"].append(werG)
                    data["pic6"]["cerg"].append(cerG)
                    data["pic6"]["levg"].append(levG)
                if image_filename.endswith('_7.jpeg'):
                    data["pic7"]["cb"].append(contained_big)
                    data["pic7"]["cs"].append(contained_small)
                    data["pic7"]["confi"].append(confidence)
                    data["pic7"]["wer"].append(wer)
                    data["pic7"]["cer"].append(cer)
                    data["pic7"]["lev"].append(lev)
                    data["pic7"]["cbg"].append(contained_big_glue)
                    data["pic7"]["csg"].append(contained_small_glue)
                    data["pic7"]["werg"].append(werG)
                    data["pic7"]["cerg"].append(cerG)
                    data["pic7"]["levg"].append(levG)
                if image_filename.endswith('_8.jpeg'):
                    data["pic8"]["cb"].append(contained_big)
                    data["pic8"]["cs"].append(contained_small)
                    data["pic8"]["confi"].append(confidence)
                    data["pic8"]["wer"].append(wer)
                    data["pic8"]["cer"].append(cer)
                    data["pic8"]["lev"].append(lev)
                    data["pic8"]["cbg"].append(contained_big_glue)
                    data["pic8"]["csg"].append(contained_small_glue)
                    data["pic8"]["werg"].append(werG)
                    data["pic8"]["cerg"].append(cerG)
                    data["pic8"]["levg"].append(levG)

            pic_names = ["pic1", "pic2", "pic3", "pic4", "pic5", "pic6", "pic7", "pic8"]
            cb_mean = {pic_name: np.mean(data[pic_name]['cb']) for pic_name in pic_names}
            cs_mean = {pic_name: np.mean(data[pic_name]['cs']) for pic_name in pic_names}
            confi_mean = {pic_name: np.mean(data[pic_name]['confi']) for pic_name in pic_names}
            wer_mean = {pic_name: np.mean(data[pic_name]['wer']) for pic_name in pic_names}
            cer_mean = {pic_name: np.mean(data[pic_name]['cer']) for pic_name in pic_names}
            lev_mean = {pic_name: np.mean(data[pic_name]['lev']) for pic_name in pic_names}
            cbg_mean = {pic_name: np.mean(data[pic_name]['cbg']) for pic_name in pic_names}
            csg_mean = {pic_name: np.mean(data[pic_name]['csg']) for pic_name in pic_names}
            werg_mean = {pic_name: np.mean(data[pic_name]['werg']) for pic_name in pic_names}
            cerg_mean = {pic_name: np.mean(data[pic_name]['cerg']) for pic_name in pic_names}
            levg_mean = {pic_name: np.mean(data[pic_name]['levg']) for pic_name in pic_names}
            new_sheet = workbook.sheets.add(name=sheet.name+'_stats')
            # Prepare data for writing
            data_to_write = {
                pic_name: [cb_mean[pic_name], cs_mean[pic_name], confi_mean[pic_name], wer_mean[pic_name],
                           cer_mean[pic_name], lev_mean[pic_name], cbg_mean[pic_name], csg_mean[pic_name],
                           werg_mean[pic_name], cerg_mean[pic_name], levg_mean[pic_name]]
                for pic_name in pic_names
            }
            means_list = ["cb_mean", "cs_mean", "confi_mean", "wer_mean", "cer_mean", "lev_mean", "cbg_mean", "csg_mean", "werg_mean",
                          "cerg_mean", "levg_mean"]
            # Write the data to the new sheet
            new_sheet.range('B1').value = means_list  # Write the column names
            new_sheet.range('A2:A9').value = [[pic_name] for pic_name in pic_names]   # Write the row names

            # Write the mean values
            for row, pic_name in enumerate(pic_names):
                new_sheet.range(f'B{row + 2}').value = data_to_write[pic_name]



def missingchar(excel_file_path):

    workbook = xw.Book(excel_file_path)
    for sheet in workbook.sheets:
        if sheet.name in ['Google', 'Azure']:
            df = pd.read_excel(excel_file_path, sheet_name=sheet.name)
            pictures = {'pic1': [], 'pic2': [], 'pic3': [], 'pic4': [], 'pic5': [], 'pic6': [], 'pic7': [], 'pic8': []}

            for index, row in df.iterrows():
                image_filename = row['Filename']
                detected_text = row['Detected_text']
                ground_truth = row['Groundtruth 1']

                if image_filename.endswith('_0.jpg'):
                    continue

                if not isinstance(ground_truth, str):
                    if math.isnan(ground_truth):
                        continue
                if not isinstance(detected_text, str):
                    if math.isnan(detected_text):
                        detected_text = ''

                normalized_detected_text = unicodedata.normalize("NFKC", str(detected_text))
                normalized_ground_truth = unicodedata.normalize("NFKC", str(ground_truth))
                missing_characters = [char[2] for char in difflib.ndiff(normalized_ground_truth, normalized_detected_text)
                                      if char.startswith('-')]

                if image_filename.endswith('_1.jpeg'):
                    pictures['pic1'].extend(missing_characters)
                elif image_filename.endswith('_2.jpeg'):
                    pictures['pic2'].extend(missing_characters)
                elif image_filename.endswith('_3.jpeg'):
                    pictures['pic3'].extend(missing_characters)
                elif image_filename.endswith('_4.jpeg'):
                    pictures['pic4'].extend(missing_characters)
                elif image_filename.endswith('_5.jpeg'):
                    pictures['pic5'].extend(missing_characters)
                elif image_filename.endswith('_6.jpeg'):
                    pictures['pic6'].extend(missing_characters)
                elif image_filename.endswith('_7.jpeg'):
                    pictures['pic7'].extend(missing_characters)
                elif image_filename.endswith('_8.jpeg'):
                    pictures['pic8'].extend(missing_characters)
            print(sheet.name)

            new_sheet_name = f'{sheet.name}_Top4MissingChars'
            new_sheet = workbook.sheets.add(new_sheet_name)

            row_index = 1  # Start at the first row
            for pic_name, missing_characters in pictures.items():
                total_characters = len(missing_characters)
                if total_characters > 0:
                    character_counts = {char: missing_characters.count(char) for char in set(missing_characters)}
                    sorted_characters = sorted(character_counts.items(), key=lambda x: x[1], reverse=True)[:4]
                    new_sheet.range(f'A{row_index}').value = f'Most Missing Characters in {pic_name}'
                    new_sheet.range(f'A{row_index + 1}').value = 'Character'
                    new_sheet.range(f'B{row_index + 1}').value = 'Percentage'
                    for i, (char, count) in enumerate(sorted_characters):
                        percentage = (count / total_characters) * 100
                        new_sheet.range(f'A{row_index + 2 + i}').number_format = '@'
                        new_sheet.range(f'A{row_index + 2 + i}').value = char
                        new_sheet.range(f'B{row_index + 2 + i}').value = f'{percentage:.2f}%'  # Format percentage
                    row_index += len(sorted_characters) + 3  # Move to the next available row
                else:
                    new_sheet.range(f'A{row_index}').value = f'No missing characters found in {pic_name}'
                    row_index += 1  # Move to the next available row


#Google_contained('Page4/Dataset_table_subscripts_page4.xlsx')
#Google_contained_glue('Page1/Dataset_table_subscripts_page1.xlsx')
#Azure_contained('Page1/Dataset_table_subscripts_page1.xlsx')
#calscore_sub('Page2/Dataset_table_subscripts_page2.xlsx')
calscore('Page1/Dataset_table_page1.xlsx')
#calculate_everything('Page3/Dataset_table_page3.xlsx')
#missingchar('Page4/Dataset_table_page4.xlsx')