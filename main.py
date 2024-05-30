import datetime
import os
import fitz
import pytesseract
from PIL import Image
import json
import pandas as pd
import var

PDF_FN = var.PDF_FN
JSON_FN = var.JSON_FN
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
DF_REPORT = pd.DataFrame(columns=['page_num', 'field', 'passed'])
HEBREW_TO_ENGLISH_MONTH = var.HEBREW_TO_ENGLISH_MONTH


def pdf_to_obj(pdf_fn=PDF_FN):
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    pdf_document = fitz.open(pdf_fn)
    pdf_data = []
    for page_num in range(pdf_document.page_count):
        # Get the current page
        page = pdf_document.load_page(page_num)

        # Render the page as an image
        pix = page.get_pixmap()
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        # Use Tesseract to extract text from the image
        page_text = pytesseract.image_to_string(img, lang='heb+eng',
                                                config='--psm 12 --oem 1')
        pdf_data.append(page_text.split("\n\n"))
    return pdf_data


def json_to_dic(json_fn=JSON_FN):
    with open(json_fn, encoding="utf8") as f:
        data = json.load(f)
    return data


def find_data_in_nested_json(json_obj, target_value):
    if isinstance(json_obj, dict):
        for key, value in json_obj.items():
            if target_value[1:] == str(value) or target_value == str(value):
                return True
            elif isinstance(value, (dict, list)):
                if find_data_in_nested_json(value, target_value):
                    return True
    elif isinstance(json_obj, list):
        for item in json_obj:
            if target_value == str(item):
                return True
            elif isinstance(item, (dict, list)):
                if find_data_in_nested_json(item, target_value):
                    return True
    return False


def check_pdf_data_on_json(pdf_content_lst, json_dic):
    for idx, data in enumerate(pdf_content_lst):
        for d in data:
            if find_data_in_nested_json(json_dic, d):
                DF_REPORT.loc[len(DF_REPORT.index)] = ["page" + str(idx), d, True]
            elif "לשנה שנסתיימה ביום " in d:
                day, month, year = d[19:].split()
                english_month = HEBREW_TO_ENGLISH_MONTH[month]
                date_obj = datetime.datetime.strptime(f'{day} {english_month} {year}', '%d %B %Y').date()
                formatted_date = date_obj.strftime('%d/%m/%Y')
                if find_data_in_nested_json(json_dic, formatted_date):
                    DF_REPORT.loc[len(DF_REPORT.index)] = ["page" + str(idx), d, True]
                else:
                    DF_REPORT.loc[len(DF_REPORT.index)] = ["page" + str(idx), d, False]
            else:
                DF_REPORT.loc[len(DF_REPORT.index)] = ["page" + str(idx), d, False]


def report_df_to_excel(df=DF_REPORT):
    today = str(datetime.datetime.now()).split('.')[0].replace(':', '_')
    excel_fn = "test_pdf_data_on_json_" + today + ".xlsx"
    df.to_excel(excel_writer=os.path.join("excel_reports", excel_fn), sheet_name="result test - pdf on json")


if __name__ == '__main__':
    pdf_content_lst = pdf_to_obj()
    json_dic = json_to_dic()
    check_pdf_data_on_json(pdf_content_lst, json_dic)
    report_df_to_excel()
