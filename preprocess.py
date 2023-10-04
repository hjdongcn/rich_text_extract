from utils import unzip_remove, unrar, ocr_api, excel_to_json, unzip_remove, generate_txt, process_eml
import utils
import argparse
import os
import subprocess
import shutil
import time
import tqdm
import re

def traverse_folder(folder_path, output_dir):
    for root, dirs, files in os.walk(folder_path):
        # root表示当前目录路径 str
        # dirs表示当前目录下的子文件夹列表 list
        # files表示当前目录下的文件列表 list
        # print(root, " | ", dirs, " | ", files)
        for file in files:
            # print(file)
            file_path = os.path.join(root, file) # 文件路径
            file_name, file_extension = os.path.splitext(file) # 文件名字和后缀
            file_txt_path = os.path.join(output_dir, file + '.txt')
            file_dir_path = os.path.join(output_dir, file_name) # 文件名字的dir
            file_extension= file_extension.lower()
            # 处理纯文本文件
            if file_extension == '.txt':
                # print(file_path, folder_path)
                shutil.copy(file_path, output_dir)

            # 处理图片文件中的文本
            elif file_extension in ['.png', '.jpg', '.jpeg']:
                text_from_image = ocr_api(file_path, file_dir_path+'.csv')
                if text_from_image != False:
                    generate_txt(file_txt_path, text_from_image)
                # os.rename(file_path, file_path+'.ed')

            # 处理邮件
            elif file_extension == '.eml':
                text_from_eml = process_eml(file_path, file_dir_path)
                generate_txt(file_txt_path, str(text_from_eml))
                if os.path.exists(file_dir_path):
                    traverse_folder(file_dir_path, output_dir)
                    shutil.rmtree(file_dir_path)
                # os.rename(file_path, file_path+'.ed')
            
            # 处理word文档
            elif file_extension in ['.wps', '.docx', '.doc']:
                # 构建要执行的命令
                command = f'libreoffice --convert-to html {file_path} --outdir {file_dir_path}'
                # 使用subprocess.run()函数执行命令
                try:
                    subprocess.run(command, shell=True, check=True)
                    # print("命令执行成功，已将文件转换为html。")
                except subprocess.CalledProcessError as e:
                    print("命令执行出错：", e)
                text_from_doc = utils.html_to_txt(file_dir_path, file_name+'.html')
                generate_txt(file_txt_path, str(text_from_doc))
                shutil.rmtree(file_dir_path)
                # os.rename(file_path, file_path+'.ed')

            # 处理ppt文档
            elif file_extension in ['.ppt', '.dps', '.pptx']:
                # 构建要执行的命令
                command = f'libreoffice --headless --convert-to pptx {file_path} --outdir {output_dir}'
                # 使用subprocess.run()函数执行命令
                try:
                    subprocess.run(command, shell=True, check=True)
                    # print("命令执行成功，已将文件转换为pptx。")
                except subprocess.CalledProcessError as e:
                    print("命令执行出错：", e)
                text_from_ppt = utils.pptx_to_txt(file_dir_path + '.pptx', file_dir_path)
                generate_txt(file_txt_path, str(text_from_ppt))
                shutil.rmtree(file_dir_path)
                # os.rename(file_path, file_path+'.ed')

            # 处理表格
            elif file_extension in ['.et', '.xlsx']:
                # text_from_table = excel_to_json(file_path)
                # generate_txt(file_txt_path, str(text_from_table))
                utils.excel_to_csv(file_path, file_dir_path)
                # os.rename(file_path, file_path+'.ed')

            # 处理SAM
            elif file_name == "sam" and file_extension != '.zip':
                # 构建要执行的命令
                command = f"samdump2 -o {file_txt_path} {root}/system{file_extension} {root}/{file_name}{file_extension}"
                # 使用subprocess.run()函数执行命令
                try:
                    subprocess.run(command, shell=True, check=True)
                    # print("命令执行成功，已将注册表文件转换为XML。")
                except subprocess.CalledProcessError as e:
                    print("命令执行出错：", e)
                    exit()
                # os.rename(file_path, file_path+'.ed')
                # sys_path = os.path.join(root, f'system{file_extension}')
                # os.rename(sys_path, sys_path+'.ed')
            
            # 解压压缩包
            elif file_extension == '.zip':
            # if file_extension == '.zip':
                unzip_remove(file_path, root)
                traverse_folder(file_path[:-4], output_dir)

            # system 跳过
            elif file_name == "system":
                pass

            else:
                shutil.copy(file_path, output_dir)
            
            print(file_txt_path)
            # print(file)

        # for dir in dirs:
        #     print("dir--------------", dir)
        #     dir_path = os.path.join(root, dir)
        #     traverse_folder(dir_path, output_dir)

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--file", "-f", help="The zip file.")
    parser.add_argument("--output", '-o', help="The output directory.")
    args = parser.parse_args()

    file_path = args.file
    output_dir = args.output
    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)
        shutil.rmtree(file_path[:-4])

    os.makedirs(output_dir)
    # print(output_path)
    unrar(file_path, file_path[:-4])
    # output_path = file_path
    traverse_folder(file_path[:-4], output_dir)  

if __name__ == "__main__":
    start = time.time()
    main()
    end = time.time()
    print(end-start)