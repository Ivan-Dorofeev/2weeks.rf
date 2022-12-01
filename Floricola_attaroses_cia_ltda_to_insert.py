import datetime
import math
import dicttoxml
import pandas as pd
import os
from ABBYY import CloudOCR
from transliterate import translit
from dateutil import parser


def pdf_convertor_to_excel(file_name, company_name):
    """Converts pdf file to xlsx format (excel)"""

    ocr_engine = CloudOCR(application_id='3310c60a-f8f7-43b9-b6b4-a1fb58f35da3', password='LkTD7NDTiCiHBt9hULyGO1hu')
    pdf_file_path = os.path.join(company_name, file_name)
    with open(f'{pdf_file_path}', 'rb') as pdf_file:
        pdf_convert = ocr_engine.process_and_download({pdf_file.name: pdf_file}, exportFormat='xlsx',
                                                      language='English')
    excel_file = file_name.split(".")[0]
    excel_file_path = os.path.join(company_name, excel_file)
    with open(f'{excel_file_path}.xlsx', 'wb') as f:
        f.write(pdf_convert['xlsx'].getbuffer())


def read_excel_file(excel_file):
    """Read Excel file and return xml"""
    read_file = pd.read_excel(excel_file, index_col=None)
    file_rows = read_file.values
    all_rows = []
    result = {
        'INVOICE_NUMBER': 0,
        'AVIA_TICKET': 0,
        'ROSE_WEIGHT': 0,
        'ALSTROMETRIA_WEIGHT': 0,
        'FULL_PLACES_COUNT': 0,
        'BOX_PLACES_COUNT': 0,
        'QB_BOX_PLACES_COUNT': 0,
        'AMS_DATA': 0,
        'MSK_DATA': 0,
        'AWB_CONTRAGENT': 0,
        'AWB_BID': 0,
        'AWB': 0,
        'PRICOOL_CONTRAGENT': 0,
        'TRACK': 0,
        'TRANSPORT_COMPANY': 0,
        'MARKING/NOTIFY': 0,
        'CONTRAGENT': 0,
        'TOT.STEMS': 0
    }
    for row in file_rows:
        row_without_nan = []
        for text in row.tolist():
            if type(text) is float:
                if math.isnan(text):
                    continue
            row_without_nan.append(text)
        if row_without_nan != []:
            all_rows.append(row_without_nan)
    result['AWB_CONTRAGENT'] = 'FLORICOLA ATTAROSES CIA.LTDA.'
    for row in all_rows:
        if 'SHIPPING DATE' in str(row):
            result['INVOICE_NUMBER'] = all_rows[all_rows.index(row) + 1][0][2:]
            ams_date = all_rows[all_rows.index(row) + 2][0]
            result['AMS_DATA'] = datetime.datetime.strftime(parser.parse(ams_date), '%d%m%Y')
        elif 'AWB' in str(row):
            result['AVIA_TICKET'] = all_rows[all_rows.index(row) + 1][0]
            result['TRANSPORT_COMPANY'] = all_rows[all_rows.index(row) + 1][2]
        elif 'FULL BOXES' in str(row):
            result['BOX_PLACES_COUNT'] = all_rows[all_rows.index(row) + 1][-1]
            result['FULL_PLACES_COUNT'] = all_rows[all_rows.index(row) + 1][-2]
        elif str(row[0]).startswith('Due Date'):
            msk_date = str(row[0]).split('Date')[-1].strip()
            result['MSK_DATA'] = datetime.datetime.strftime(parser.parse(msk_date), '%d%m%Y')
        elif 'Totals' in str(row):
            result['AWB'] = row[-1].split(' ')[-1]
            result['TOT.STEMS'] = row[2]
        elif str(row[0]).startswith('CONSIGNED'):
            result['MARKING/NOTIFY'] = row[-1]
        elif 'HB' in str(row):
            products = [{
                'name': translit(str(row[2]), "ru"),
                'count': row[6],
                'price': str(row[7]).split(' ')[-1],
                'sum': str(row[8]).split(' ')[-1],
                'total_stems': row[5],
            }]
            while True:
                row = all_rows[all_rows.index(row) + 1]
                if 'Totals' in str(row):
                    break
                products.append({
                    'name': translit(str(row[0]), "ru"),
                    'count': row[2],
                    'price': str(row[-2]).split(' ')[-1],
                    'sum': str(row[-1]).split(' ')[-1],
                    'total_stems': row[-3],
                })
            result['PRODUCTS'] = products

    xml_file = dicttoxml.dicttoxml(result)
    print(xml_file)
    write_file = excel_file[:-5] + '.xml'
    with open(write_file, 'wb') as xml_ff:
        xml_ff.write(xml_file)


def remove_excel_files(company_name):
    for root, dirs, files in os.walk(company_name):
        for file in files:
            if file.endswith(".xlsx"):
                file_path = os.path.join(company_name, file)
                os.remove(file_path)


if __name__ == "__main__":
    company_name = 'FLORICOLA ATTAROSES CIA.LTDA.'

    for root, dirs, files in os.walk(company_name):
        for file in files:
            if file.endswith(".pdf"):
                pdf_convertor_to_excel(file, company_name)
    for root, dirs, files in os.walk(company_name):
        for file in files:
            if file.endswith(".xlsx"):
                file_path = os.path.join(company_name, file)
                dataframe_from_file = read_excel_file(file_path)
    remove_excel_files(company_name)
