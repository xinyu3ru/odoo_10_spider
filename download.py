#!/bin/env python
# -*- coding: utf-8 -*-

__author__ = 'rublog'

from typing import Any, Union

from requests import Session, Request
import re
import logging
import random
import time
import config

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

base_url = config.BASEURL
e_mail = config.EMAIL
password = config.PASSWORD
host = config.HOST


def search_customer(s, start_customer_id, headers, user_id='1781'):
    # 搜索关键词客户
    search_url = base_url + '/web/dataset/search_read'
    payload_search_template = """{{"jsonrpc":"2.0","method":"call","params":{{"model":"res.partner","fields":["sale_order_count","display_name","title","parent_id","street","country_id","debit","email","city","credit","color","zip","street2","type","opportunity_count","function","phone","image_small","state_id","is_company","category_id","meeting_count","mobile","__last_update"],"domain":[],"context":{{"lang":"zh_CN","tz":"Asia/Chongqing","uid":{2},"search_default_customer":1,"params":{{"action":55}}}},"offset":{0},"limit":40,"sort":""}},"id":{1}}}"""
    payload_search = payload_search_template.format(start_customer_id, random_id(), user_id).encode('utf-8')
    req_search = Request('post', search_url, headers=headers, data=payload_search)
    prepped_search_name = s.prepare_request(req_search)
    r = s.send(prepped_search_name)
    logger.info("搜索客户，服务器返回代码" + str(r.status_code))
    # logger.info(r.text)
    logger.info("搜索客户，获取的客户信息")
    logger.info(r.json())
    customer_id = r.json()
    save_customer_id(start_customer_id, customer_id)
    return customer_id


def save_customer_id(file_name, customer_id):
    # 获取到的资料保存到json文件
    with open(str(file_name) + '.json', mode='a+', encoding='utf-8') as f:
        f.write(str(customer_id))


def main():
    # 创建基础session和获取csrf
    s = Session()
    headers = {'Host': host,
               'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:71.0) Gecko/20100101 Firefox/71.0',
               'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
               'Accept-Language': 'zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2',
               'Accept-Encoding': 'gzip, deflate',
               'Connection': 'keep-alive',
               'Upgrade-Insecure-Requests': '1'
               }
    url = base_url + '/web/login'
    req_start = Request('get', url, headers=headers)
    prepped = s.prepare_request(req_start)
    r = s.send(prepped)

    content = r.text
    csrf_token = re.findall('csrf_token" value="(.*?)"', content)[0]
    logger.info("获取到的csrf_token是" + csrf_token)
    # 登陆账号
    payload = {'csrf_token': csrf_token, 'login': e_mail, 'password': password}
    login_url = base_url + "/web/login"
    req_login = Request('post', login_url, headers=headers, params=payload)
    prepped_login = s.prepare_request(req_login)
    r = s.send(prepped_login)
    logger.info("登陆页面，服务器返回代码" + str(r.status_code))
    # logger.info(r.request.headers)


    # 组合浏览器头，获取列表
    headers = {'Host': host,
               'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:72.0) Gecko/20100101 Firefox/72.0',
               'Accept': 'application/json, text/javascript, */*; q=0.01',
               'Accept-Language': 'zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2',
               'Accept-Encoding': 'gzip, deflate',
               'Content-Type': 'application/json',
               'X-Requested-With': 'XMLHttpRequest',
               'Origin': base_url,
               'Connection': 'keep-alive',
               'cookie': 'hibext_instdsigdipv2=1',
               'Referer': base_url + '/web'
               }
    for i in r.cookies.items():
        yyy = '='.join(i)
        logger.info("组合获得header中的session_id是" + yyy)
    headers['cookie'] = ';'.join([headers['cookie'], yyy])
    headers['cookie'] = ';'.join([headers['cookie'], str("增加功能13".encode('utf-8'))])
    customer_num = 50160 #50160
    while customer_num >= 40:
        search_customer(s, customer_num, headers)
        customer_num = customer_num - 40
        time.sleep(random.randint(1, 5))



def random_id() -> object:
    return random.randrange(130000000, 799999999)


if __name__ == "__main__":
    main()
