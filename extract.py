#!/bin/env python
# -*- coding: utf-8 -*-

__author__ = 'rublog'


import logging
import openpyxl
import ast
from openpyxl import Workbook, load_workbook
import json
import os


logger = logging.getLogger()
logger.setLevel('DEBUG')
BASIC_FORMAT = "%(asctime)s:%(levelname)s:%(message)s"
DATE_FORMAT = '%Y-%m-%d %H:%M:%S'
formatter = logging.Formatter(BASIC_FORMAT, DATE_FORMAT)
cons_log = logging.StreamHandler()  # 输出到控制台的handler
cons_log.setFormatter(formatter)
cons_log.setLevel('INFO')  # 也可以不设置，不设置就默认用logger的level
file_log = logging.FileHandler('output.log', encoding='utf-8')  # 输出到文件的handler
file_log.setFormatter(formatter)
logger.addHandler(cons_log)
logger.addHandler(file_log)
logger.info('this is info')
logger.debug('this is debug')


def save_customer_id(file_name, customer_id):
    with open(str(file_name) + '.json', mode='a+', encoding='utf-8') as f:
        f.write(str(customer_id))


def main():
    # 将获取的json资料读取保存到xlsx
    # 为啥是xlsx，因为表格相对比较直观好整理资料
    xls_book = load_workbook('20.xlsx')
    xls_sheet = xls_book.active
    file_names = os.listdir()
    for file_name in file_names:
        if os.path.splitext(file_name)[1] == '.json':
            # print(file_name)
            logger.info("正在处理" + file_name)
            with open(file_name, mode='r', encoding='utf-8') as f:
                txt_content = f.readlines()
                with open("wrong_lins.json", mode='a+', encoding='utf-8') as k:
                    for line in txt_content:
                        try:
                            json_content = json.loads(line)
                            xls_content = list(json_content.values())
                            # print(xls_content)
                            xls_sheet.append(xls_content)
                        except:
                            k.write(line)
                        finally:
                            pass
                logger.info(file_name + "的数据正在保存中~")
    xls_book.save("20.xlsx")
    


if __name__ == "__main__":
    main()
