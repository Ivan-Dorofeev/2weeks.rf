import datetime
import json
import math
import re
import dicttoxml
import pandas as pd
import os
from ABBYY import CloudOCR
from transliterate import translit
from dateutil import parser
import warnings

warnings.simplefilter("ignore")


def pdf_convertor_to_excel(file_name, company_name):
    """Converts pdf file to xlsx format (excel)"""

    ocr_engine = CloudOCR(application_id='3310c60a-f8f7-43b9-b6b4-a1fb58f35da3', password='LkTD7NDTiCiHBt9hULyGO1hu')
    pdf_file_path = os.path.join(company_name, file_name)
    with open(f'{pdf_file_path}', 'rb') as pdf_file:
        pdf_convert = ocr_engine.process_and_download({pdf_file.name: pdf_file}, exportFormat='xlsx',
                                                      language='English')
    excel_file = f'Floreloy_{file_name.split(".")[0]}'
    excel_file_path = os.path.join(company_name, excel_file)
    with open(f'{excel_file_path}.xlsx', 'wb') as f:
        f.write(pdf_convert['xlsx'].getbuffer())


def remove_excel_files(company_name):
    """Delete .xlsx files"""
    for root, dirs, files in os.walk(company_name):
        for file in files:
            if file.endswith(".xlsx"):
                file_path = os.path.join(company_name, file)
                os.remove(file_path)


def prepare_box(box):
    """Check and prepare box to mix or sorted"""
    names = []
    for product in box:
        names.append(product['name'])
    if len(set(names)) == 1:
        packing = 0
        for p in box:
            packing += int(p['total_stems'])
        for p in box:
            p['total_stems'] = packing
        return box
    else:
        new_box = box[0]
        new_box['is_mixed'] = True
        new_box['name'] = 'МИКС'
        new_box['total_stems'] = sum(p['total_stems'] for p in box)
        new_box['count'] = sum(p['count'] for p in box)
        new_box['sum'] = sum(p['sum'] for p in box)
        return [new_box]


def prepare_marking(mark):
    """check marking in file and conditions """
    with open('marking.json', 'r') as mark_file:
        marking = json.load(mark_file)[0]['marks']
    if mark in marking:
        return mark
    elif ' ' in mark:
        mark = ''.join(mark.split(' ')[1:])
        if mark in marking:
            return mark
    elif '-' in mark:
        mark.replace('-', '')
        if mark in marking:
            return mark
    return 'NA'


def read_excel_file(excel_file):
    """Read Excel file and save to xml"""
    read_file = pd.read_excel(excel_file, index_col=None, header=None)
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
        'PRICOOL_CONTRAGENT': 0,
        'TRACK': 0,
        'TRANSPORT_COMPANY': 0,
        'MARKING/NOTIFY': 0,
        'CONTRAGENT': 0,
        'PRODUCTS': []
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

    result['AWB_CONTRAGENT'] = 'ECUCARGA CIA LTDA'
    result['MARKING/NOTIFY'] = 'FP21'
    result['PRICOOL_CONTRAGENT'] = 'Floreloy'
    with open('traslate.json', 'r') as tr_file:
        translate_dict = json.load(tr_file)
    for row in all_rows:
        if 'INVOICE:' in str(row) and len(row) == 1:
            result['INVOICE_NUMBER'] = all_rows[all_rows.index(row) + 3][0]
            ams_date = all_rows[all_rows.index(row) + 2][0]
            result['AMS_DATA'] = datetime.datetime.strftime(parser.parse(ams_date), '%Y%m%d')
        elif 'AWB' in str(row):
            result['AVIA_TICKET'] = row[1]
        elif 'HB-R' in str(row):
            if isinstance(row[-2], str):
                result['BOX_PLACES_COUNT'] = int(str(row[-2]).split(' ')[0])
            else:
                result['BOX_PLACES_COUNT'] = int(row[-2])
        elif 'Varieties/Length' in str(row):
            for i in all_rows:
                if 'Full Boxes:' in str(i):
                    result['FULL_PLACES_COUNT'] = float(str(i[0]).split(' ')[-1])
                    break
            row_index = all_rows.index(row)
            box = []
            box_count = 0
            while True:
                row_index += 1
                row = all_rows[row_index]
                if 'Full Boxes:' in str(row):
                    if box:
                        prepared_box = prepare_box(box)
                        for p in prepared_box:
                            result['PRODUCTS'].append(p)
                    break
                if 'HB-R' in str(row):
                    box_count += 1
                    if box:
                        prepared_box = prepare_box(box)
                        for p in prepared_box:
                            result['PRODUCTS'].append(p)
                        box = []

                    if not result['MARKING/NOTIFY']:
                        marking = prepare_marking(str(row[0]))
                        result['MARKING/NOTIFY'] = marking
                    continue

                re_find = re.findall(r'(.{1,}\d{1,3})', row[-3])[0]
                size = re.findall(r'\d{1,3}', re_find)[0]
                name = str(re.findall(r'\D+', re_find)[0]).strip()
                product_name = str(name)[3:] if str(name).startswith('GR ') else str(name)
                nomenklatur_name = result['PRICOOL_CONTRAGENT']
                box.append(
                    {
                        'name': str(translate_dict[product_name]) if translate_dict.get(product_name) else translit(
                            product_name, "ru"),
                        'count': row[0],
                        'nomenclature_characteristic': f'{nomenklatur_name}, {int(size)} см',
                        'size': int(size),
                        'price': row[-2],
                        'sum': row[-1],
                        'total_stems': row[0],
                        'is_mixed': False,
                    })
    result['ROSE_WEIGHT'] = float(25) * box_count
    date_from_ams = datetime.datetime.strptime(result['AMS_DATA'], '%Y%m%d') + datetime.timedelta(days=5)
    result['MSK_DATA'] = datetime.datetime.strftime(date_from_ams, '%Y%m%d')

    xml_file = dicttoxml.dicttoxml(result)
    write_file = excel_file[:-5] + '.xml'
    with open(write_file, 'wb') as xml_ff:
        xml_ff.write(xml_file)


if __name__ == "__main__":
    company_name = 'invoices/Floreloy'

    for root, dirs, files in os.walk(company_name):
        for file in files:
            if file.endswith(".pdf"):
                pdf_convertor_to_excel(file, company_name)
    for root, dirs, files in os.walk(company_name):
        for file in files:
            if file.endswith(".xlsx"):
                print(file)
                file_path = os.path.join(company_name, file)
                dataframe_from_file = read_excel_file(file_path)
    remove_excel_files(company_name)
