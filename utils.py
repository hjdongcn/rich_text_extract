from PIL import Image
import subprocess
import pytesseract
import cv2
import pandas as pd
import zipfile
import rarfile
import os
import eml_parser
import base64
import datetime
import json
import pathlib
from bs4 import BeautifulSoup
import re
import urllib.parse
import xml.etree.ElementTree as ET
from paddleocr import PaddleOCR, PPStructure
from paddleocr import paddleocr
import csv

import logging
import time

ocr = PaddleOCR(use_angle_cls=True)
table_engine = PPStructure(show_log=True)
paddleocr.logging.disable(logging.DEBUG)

# convert doc file to docx file
def convert_doc_to_docx(doc_path):
    subprocess.run(['libreoffice', '--headless', '--convert-to', 'docx', doc_path])

# convert doc file to html file
def convert_doc_to_html(doc_path):
    subprocess.run(['libreoffice', '--headless', '--convert-to', 'html', '--convert-images-to', 'jpg', doc_path])

# OCR by Python library
def ocr_lib(img_path):
    # image = Image.open(img_path)
    img = cv2.imread(img_path)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    # # 虚化处理
    # blurMedian = cv2.medianBlur(gray, 3)	# 中值虚化处理
    # blurGaussian = cv2.GaussianBlur(gray,(5,5),0)	# 高斯虚化处理
    # otsuThreshold1 = cv2.threshold(blurMedian, 0, 255, cv2.THRESH_BINARY+cv2.THRESH_OTSU)[1]
    # thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
    # cnts = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    # cnts = cnts[0] if len(cnts) == 2 else cnts[1]
    ocr_text = pytesseract.image_to_string(gray, lang='chi_sim+eng')
    return ocr_text


# OCR by API
def ocr_api(img_path, csv_path=None):
    result = table_engine(img_path)
    ocr_text = ''
    now_line = 0.0
    for line in result:
        if line['type'] == 'table':
            html = line['res']['html']
            soup = BeautifulSoup(html, 'html.parser')
            # 查找HTML中的表格（假设您只有一个表格）
            table = soup.find('table')

            # 如果有多个表格，可以使用find_all来获取所有表格

            # 打开CSV文件以写入数据
            if csv_path == None:
                csv_path = '/home/norainy/jingsai/output1/' + str(time.time()) + '.csv'
            with open(csv_path, 'w', newline='', encoding='utf-8') as csv_file:
                csv_writer = csv.writer(csv_file)

                # 遍历表格的行和列，将数据写入CSV文件
                for row in table.find_all('tr'):
                    csv_row = []
                    for cell in row.find_all(['th', 'td']):
                        csv_row.append(cell.get_text(strip=True))
                    csv_writer.writerow(csv_row)
            return False
        else:
            for i in line['res']:
                if (i['text_region'][0][1] + i['text_region'][3][1])/2 > now_line:
                    now_line = i['text_region'][3][1]
                    ocr_text += '\n' + i['text']
                else:
                    ocr_text += ' ' + i['text']
                # ocr_text += i['text']    
    return ocr_text
    # img = cv2.imread(img_path)
    # gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
    # blurred_image = cv2.GaussianBlur(gray, (5, 5), 0)
    # result = ocr.ocr(img_path)
    # now_line = 0.0
    # ocr_text = ''
    # for idx in range(len(result)):
    #     # ocr_text += result[idx][1][0]
    #     res = result[idx]
    #     # print(res)
    #     # print(res[0][0][1])
    #     if (res[0][0][1] + res[0][3][1])/2 > now_line:
    #         now_line = res[0][3][1]
    #         ocr_text += '\n' + res[1][0]
    #     else:
    #         ocr_text += ' ' + res[1][0]
    # return ocr_text

def excel_to_json(excel_file_path):
    # 使用pandas读取Excel文件的所有工作表
    sheets = pd.read_excel(excel_file_path, sheet_name=None)
    # 创建一个字典，用于存储每个工作表的JSON数据
    json_data = {}
    # 遍历每个工作表，并将其转换为JSON格式
    for sheet_name, df in sheets.items():
        # 将DataFrame转换为JSON格式
        json_data[sheet_name] = df.to_json(orient='records', lines=True, force_ascii=False)
    return json_data

def excel_to_csv(excel_file_path, root_path):
    if not os.path.exists(root_path):
        os.mkdir(root_path)
    # 读取 Excel 文件
    xls_file = pd.ExcelFile(excel_file_path)

    # 获取 Excel 文件中的工作表列表
    sheet_names = xls_file.sheet_names
    # print(sheet_names)
    # 遍历每个工作表并将其保存为 CSV 文件
    for sheet_name in sheet_names:
        # 从 Excel 文件中读取工作表数据
        df = xls_file.parse(sheet_name)
        
        # 将工作表数据保存为 CSV 文件
        csv_filename = f'{root_path}/{sheet_name}.csv'
        df.to_csv(csv_filename, index=False, encoding='utf-8')


# 解压压缩包
def unzip_remove(zip_file_path, extract_to_path):
    with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
        zip_ref.extractall(extract_to_path)
    os.remove(zip_file_path)
    return 

# unzip rar file
def unrar(rar_file, output_dir):
    """Unzips a rar file to the specified output directory.

    Args:
    rar_file: The rar file to unzip.
    output_dir: The output directory to unzip the rar file to.
    """

    try:
        with rarfile.RarFile(rar_file) as rf:
            rf.extractall(output_dir)
        print("RAR file extraction successful.")
    except Exception as e:
        print(f"Error while extracting RAR file: {e}")

# 处理eml文件
def process_eml(eml_file_path, file_dir_path):
    ep = eml_parser.EmlParser(include_raw_body=True, include_attachment_data=True)
    parsed_eml = ep.decode_email(eml_file_path)
    eml_text = ''
    # print(ep.get_raw_body_text(eml_file_path))
    # out_path = pathlib.Path(eml_file_path[:eml_file_path.rfind('/') + 1])
    if not os.path.exists(file_dir_path):
        os.mkdir(file_dir_path)
    out_path = pathlib.Path(file_dir_path)
    
    if 'attachment' in parsed_eml:
        for a in parsed_eml['attachment']:
            out_filepath = out_path / a['filename']

            print(f'\tWriting attachment: {out_filepath}')
            with out_filepath.open('wb') as a_out:
                a_out.write(base64.b64decode(a['raw']))
        
    # print(parsed_eml)
    for i in ['subject', 'from', 'to']:
        eml_text += i + ': ' + str(parsed_eml['header'][i]) + '\n'
    for b in parsed_eml['body']:
        # print('--------------')
        # print(b)
        for i in ['email', 'domain', 'ip']:
            eml_text += i + ': ' + str(b[i]) + '\n'
        soup = BeautifulSoup(b['content'], 'html.parser')
        eml_text += soup.get_text(separator=" ", strip=True)
        # print(soup.get_text(separator=' ', strip=True))
        # res1=re.findall(r"<p[^>]*>(.*?)</p>|<td[^>]*>(.*?)</td>", str(b['content']))
        # res1 = b['content']
        # for r in res1:
        #     eml_text += r[0]+r[1]+'\n'
    # print(parsed_eml['body'][1]) # attachment body header
    # print(parsed_eml) # attachment body header
    

    # print(parsed_eml['attachment'])
    # return json.dumps(parsed_eml, default=json_serial, ensure_ascii=False)
    return eml_text

def json_serial(obj):
  if isinstance(obj, datetime.datetime):
      serial = obj.isoformat()
      return serial
  
# 生成txt文件
def generate_txt(file, content):
    with open(file, 'w') as f:
        f.write(content)
        
# 处理html文件
def html_to_txt(html_dir, html_name):
    html_path = os.path.join(html_dir, html_name)
    # 读取HTML文件内容
    with open(html_path, "r", encoding="utf-8") as f:
        html_content = f.read()

    # 使用Beautiful Soup解析HTML
    soup = BeautifulSoup(html_content, "html.parser")
    text = ""
    # 遍历HTML文档节点
    for element in soup.descendants:
        # 提取文本内容
        if isinstance(element, str) and 'page' not in element.strip():
            text = text + element.strip()
            # print(element.strip())
        # 提取图片名称
        elif element.name == "img":
            src = element.get("src")
            src = urllib.parse.unquote(src)
            # print("src: ", src)
            # img_name = re.search(r'/([^/]+)$', src)
            # if img_name:
            #     print("Image Name:", img_name.group(1))
            img_path = os.path.join(html_dir, src)
            image_txt = ocr_api(img_path)
            if image_txt != False:
            # text = text + '\n From image:\n' + image_txt
                text = text + image_txt
    return text


def sort_by_number(filename):
    # 使用正则表达式提取文件名中的数字部分
    match = re.search(r'\d+', filename)
    if match:
        return int(match.group())
    else:
        return filename



# 处理html文件
def pptx_to_txt(pptx_path, pptx_dir):
    os.rename(pptx_path, pptx_dir+'.zip')
    unzip_remove(pptx_dir+'.zip', pptx_dir)
    text = ''
    img_dir = os.path.join(pptx_dir, 'ppt', 'media')
    slide_dir = os.path.join(pptx_dir, 'ppt', 'slides')
    for root, dirs, files in os.walk(slide_dir):
        # files = sorted(files, key=sort_by_number)  # 对文件排序
        for file in files:
            file_path = os.path.join(root, file)
            with open(file_path) as f:
                # pass
                xml_root = ET.parse(f).getroot()
                for elem in xml_root.iter():
                    namespace, element_name = elem.tag.split("}", 1)
                    # print(element_name)
                    if element_name == 't':
                        # print(ET.tostring(elem, encoding='unicode'))
                        content = elem.text
                        text += content
    for root, dirs, files in os.walk(img_dir):
        for file in files:
            file_path = os.path.join(root, file)
            extension = file_path.split('.')[-1]
            # print(file_path)
            if extension in ['jpg', 'jpeg', 'png']:
                t = ocr_api(file_path)
                if t != False:
                    text += t
    return text



if __name__ == "__main__":
    start = time.time()
    text = ocr_api('/home/norainy/jingsai/赛题材料/麒麟SSL+VPN+Windows客户端使用手册_html_98bf9c15f498bb9a.png')
    # print(text)
    # text = process_eml('/home/norainy/jingsai/题目1：富文本敏感信息泄露检测/赛题材料/xxx部门弱口令漏洞问题和整改 2023-05-25T17_27_32+08_00.eml')
    # print(text)
    # text = excel_to_csv('/home/norainy/jingsai/赛题材料/wps/资产梳理.et', '/home/norainy/jingsai/赛题材料/wps')
    print(text)
    end = time.time()
    print(end-start)