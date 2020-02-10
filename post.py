#!/bin/env python
# -*- coding: utf-8 -*-

__author__ = 'rublog'

from typing import Any, Union

from requests import Session, Request
import re
import logging
import random
import openpyxl
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


def get_customer_info(s, customer_id, headers):
    # 获取账号信息
    read_url = base_url + '/web/dataset/call_kw/res.partner/read'
    payload_read_template = """{{"jsonrpc":"2.0","method":"call","params":{{"model":"res.partner","method":"read","args":[[{0}],["sale_order_count","bank_account_count","property_product_pricelist","opt_out","title","company_id","parent_id","contracts_count","supplier_invoice_count","property_supplier_payment_term_id","customer","fax","crm_lead_id","child_ids","user_ids","name","deliver_rule","commercial_partner_id","message_follower_ids","total_invoiced","notify_email","currency_id","street","country_id","debit","supplier","large_product_location","ref","email","picking_warn","city","display_customer_delivery","product_ship_price","tax_rate","active","ship_price","finished_qty","zip","property_account_payable_id","gilding_ids","credit","message_bounce","comment","sale_warn","image","trust","property_payment_term_id","often_product_line","user_id","property_delivery_carrier_id","street2","type","date_type","opportunity_count","function","picking_warn_msg","phosee_partner","payment_token_count","consignee","activities_count","phone","task_count","fixed_value","state_id","invoice_warn_msg","website","company_type","third_party_partner","property_account_receivable_id","message_ids","need_tax","invoice_warn","property_account_position_id","company_name","property_stock_supplier","is_company","property_stock_customer","forthcoming_death","category_id","lang","meeting_count","mobile","interval","last_deliver_date","sale_warn_msg","display_name","__last_update"]],"kwargs":{{"context":{{"lang":"zh_CN","tz":"Asia/Chongqing","uid":1781,"search_default_customer":1,"params":{{"action":55}},"bin_size":true}}}}}},"id":{1}}}"""
    payload_read = payload_read_template.format(customer_id, random_id()).encode('utf-8')
    req_read = Request('post', read_url, headers=headers, data=payload_read)
    prepped_read_name = s.prepare_request(req_read)
    r = s.send(prepped_read_name)
    logger.info("读取客户信息，服务器返回代码" + str(r.status_code))
    # logger.info(r.text)
    logger.info("读取客户信息，获取的客户信息")
    # logger.info(r.json())
    customer_info = r.json()['result'][0]
    logger.info(customer_info)
    return customer_info


def search_customer_id(s, customer_key_word, headers, user_id='1781'):
    # 搜索关键词客户
    search_url = base_url + '/web/dataset/search_read'
    payload_search_template = """{{"jsonrpc":"2.0","method":"call","params":{{"model":"res.partner","fields":["sale_order_count","display_name","title","parent_id","street","country_id","debit","email","city","zip","credit","color","street2","type","opportunity_count","function","phone","image_small","state_id","is_company","category_id","meeting_count","mobile","__last_update"],"domain":[["customer","=",1],["parent_id","=",false],"|","|",["display_name","ilike","{0}"],["ref","=","{0}"],["email","ilike","{0}"]],"context":{{"lang":"zh_CN","tz":"Asia/Chongqing","uid":{1},"search_default_customer":1,"params":{{"action":55}}}},"offset":0,"limit":40,"sort":""}},"id":{2}}}"""
    payload_search = payload_search_template.format(customer_key_word, user_id, random_id()).encode('utf-8')
    req_search = Request('post', search_url, headers=headers, data=payload_search)
    prepped_search_name = s.prepare_request(req_search)
    r = s.send(prepped_search_name)
    logger.info("搜索客户，服务器返回代码" + str(r.status_code))
    # logger.info(r.text)
    logger.info("搜索客户，获取的客户信息")
    logger.info(r.json())
    customer_id = r.json()['result']['records'][0]['id']
    return customer_id


def write_customer_info(s, payload_edit, headers):
    # 获取账号信息
    write_url = base_url + '/web/dataset/call_kw/res.partner/write'
    req_read = Request('post', write_url, headers=headers, data=payload_edit.encode('utf-8'))
    prepped_read_name = s.prepare_request(req_read)
    r = s.send(prepped_read_name)
    logger.info("写入客户信息，服务器返回代码" + str(r.status_code))
    # logger.info(r.text)
    logger.info("写入客户信息，服务器返回的信息")
    # logger.info(r.json())
    customer_info = r.json()
    logger.info(customer_info)
    return customer_info


def judge_user_id(customer_info, user_id='1781'):
    if str(customer_info['user_id']) == user_id:
        return 0
    else:
        return 1


def judge_phone(customer_info):
    if customer_info['phone'] in ['False', 'false', '']:
        return ['phone', customer_info['mobile']]
    elif not customer_info['phone']:
        return ['phone', customer_info['mobile']]
    elif customer_info['mobile'] in ['False', 'false', '']:
        return ['mobile', customer_info['phone']]
    elif not customer_info['mobile']:
        return ['mobile', customer_info['phone']]
    else:
        return ['both have']


def judge_name(customer_info):
    col_a = customer_info['name']
    if col_a[0:2] in ["石家", "唐山", "秦皇", "邯郸", "邢台", "保定", "张家", "承德", "沧州", "廊坊", "太原", "衡水", "大同", "阳泉", "长治",
                      "晋城", "朔州", "晋中", "运城", "忻州", "临汾", "吕梁", "呼和", "包头", "乌海", "赤峰", "通辽", "鄂尔", "呼伦", "巴彦",
                      "乌兰", "兴安", "锡林", "阿拉", "沈阳", "大连", "鞍山", "抚顺", "本溪", "丹东", "锦州", "营口", "阜新", "辽阳", "盘锦",
                      "铁岭", "朝阳", "葫芦", "长春", "吉林", "四平", "辽源", "通化", "白山", "松原", "白城", "延边", "鲜族", "哈尔", "齐齐",
                      "鸡西", "鹤岗", "双鸭", "大庆", "伊春", "佳木", "七台", "牡丹", "黑河", "绥化", "大兴", "南京", "无锡", "徐州", "常州",
                      "苏州", "南通", "连云", "淮安", "盐城", "扬州", "镇江", "泰州", "宿迁", "杭州", "宁波", "温州", "嘉兴", "湖州", "绍兴",
                      "金华", "衢州", "舟山", "台州", "丽水", "合肥", "芜湖", "蚌埠", "淮南", "马鞍", "淮北", "铜陵", "安庆", "黄山", "滁州",
                      "阜阳", "宿州", "六安", "亳州", "池州", "宣城", "福州", "厦门", "莆田", "三明", "泉州", "漳州", "南平", "龙岩", "宁德",
                      "南昌", "景德", "萍乡", "九江", "新余", "鹰潭", "赣州", "吉安", "宜春", "抚州", "上饶", "济南", "青岛", "淄博", "枣庄",
                      "东营", "烟台", "潍坊", "济宁", "泰安", "威海", "日照", "莱芜", "临沂", "德州", "聊城", "滨州", "菏泽", "郑州", "开封",
                      "洛阳", "平顶", "安阳", "鹤壁", "新乡", "焦作", "濮阳", "许昌", "漯河", "三门", "南阳", "商丘", "信阳", "周口", "驻马",
                      "济源", "武汉", "黄石", "十堰", "宜昌", "襄阳", "鄂州", "荆门", "孝感", "荆州", "黄冈", "咸宁", "随州", "恩施", "长沙",
                      "株洲", "湘潭", "衡阳", "邵阳", "岳阳", "常德", "张家", "益阳", "郴州", "永州", "怀化", "娄底", "湘西", "广州", "韶关",
                      "深圳", "珠海", "汕头", "佛山", "江门", "湛江", "茂名", "肇庆", "惠州", "梅州", "汕尾", "河源", "阳江", "清远", "东莞",
                      "中山", "潮州", "揭阳", "云浮", "南宁", "柳州", "桂林", "梧州", "北海", "防城", "钦州", "贵港", "玉林", "百色", "贺州",
                      "河池", "来宾", "崇左", "海口", "三亚", "三沙", "儋州", "成都", "自贡", "攀枝", "泸州", "德阳", "绵阳", "广元", "遂宁",
                      "内江", "乐山", "南充", "眉山", "宜宾", "广安", "达州", "雅安", "巴中", "资阳", "阿坝", "甘孜", "凉山", "贵阳", "六盘",
                      "遵义", "安顺", "毕节", "铜仁", "黔西", "黔东", "黔南", "昆明", "曲靖", "玉溪", "保山", "昭通", "丽江", "普洱", "临沧",
                      "楚雄", "红河", "文山", "西双", "大理", "德宏", "怒江", "迪庆", "拉萨", "昌都", "山南", "日喀", "那曲", "阿里", "林芝",
                      "西安", "铜川", "宝鸡", "咸阳", "渭南", "延安", "汉中", "榆林", "安康", "商洛", "兰州", "嘉峪", "金昌", "白银", "天水",
                      "武威", "张掖", "平凉", "酒泉", "庆阳", "定西", "陇南", "临夏", "甘南", "西宁", "海东", "海北", "黄南", "海南", "果洛",
                      "玉树", "海西", "银川", "石嘴", "吴忠", "固原", "中卫", "乌鲁", "克拉", "吐鲁", "哈密", "昌吉", "博尔", "巴音", "阿克",
                      "克孜", "喀什", "和田", "伊犁", "塔城", "阿勒"]:
        return 1
    elif col_a[0:2] in ["澳洲", "北京", "天津", "上海", "重庆", "河北", "山西", "辽宁", "吉林", "黑龙江", "江苏", "浙江", "安徽", "福建", "江西",
                        "山东", "河南", "湖北", "湖南", "广东", "海南", "四川", "贵州", "云南", "陕西", "甘肃", "青海", "台湾", "内蒙古", "内蒙",
                        "广西", "西藏", "宁夏", "新疆", "香港", "澳门", "美国"]:
        return 0
    else:
        logger.info(col_a)
        pass


def save_vcf(customer_info):
    vcards_template = """BEGIN:VCARD
VERSION:3.0
EMAIL;TYPE=INTERNET:{0}
FN:{1}
TEL;CELL:{2}
END:VCARD
"""
    one_vcard = vcards_template.format(customer_info['email'], customer_info['name'], customer_info['mobile'])
    with open('mobile.vcf', mode='a+', encoding='utf-8') as f:
        f.write(one_vcard)


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

# 打开客户信息表
    wb = openpyxl.load_workbook('2020.xlsx')
    wx = wb["rublog"]

    # 搜索客户
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
    # ['='.join(str(i)) for i in r.cookies.items()]])
    # customer_id = search_customer_id(s, '西安萝卜头摄影', headers)
    # print("搜索到的客户信息如下")
    # print(customer_id)
    # customer_info = get_customer_info(s, 772, headers)
    # print(customer_info)
    max_row = wx.max_row
    for i in range(3, max_row):
        col_num = "A" + str(i)
        col_a = wx[col_num].value
        customer_id = search_customer_id(s, col_a, headers)
        customer_info = get_customer_info(s, customer_id, headers)
        edit_info = ""
        payload_template = """{{"jsonrpc":"2.0","method":"call","params":{{"model":"res.partner","method":"write","args":[[{0}],{1}],"kwargs":{{"context":{{"lang":"zh_CN","tz":"Asia/Chongqing","uid":{2},"search_default_customer":1,"params":{{"action":55}}}}}}}},"id":{3}}}"""
        if judge_name(customer_info):
            full_name = customer_info['zip'][0:2] + customer_info['name']
            edit_info = '{"name":"' + full_name + '"'
            customer_info['name'] = full_name
            if customer_info['child_ids']:
                child_ids_tmp = ',"child_ids":[[4,'
                for j in customer_info['child_ids']:
                    child_ids_tmp = child_ids_tmp + str(j) + ',false],[4,'
                edit_info = "{0}{1}{2}".format(edit_info, child_ids_tmp[0:-4], ']')
        logger.info('添加完name的edit_info')
        logger.info(edit_info)
        if judge_user_id(customer_info):
            if edit_info:
                edit_info = edit_info + ',"user_id":1781'
            else:
                edit_info = '{"user_id":1781'
        logger.info('添加完user_id的edit_info')
        logger.info(edit_info)
        cell_phone = judge_phone(customer_info)
        if cell_phone[0] == 'mobile':
            customer_info['mobile'] = cell_phone[1]
            if edit_info:
                edit_info = edit_info + ',"mobile":"' + str(cell_phone[1]) + '"'
            else:
                edit_info = '"mobile":"' + str(cell_phone[1]) + '"'
        elif cell_phone[0] == 'phone':
            customer_info['phone'] = cell_phone[1]
            if edit_info:
                edit_info = edit_info + ',"phone":"' + str(cell_phone[1]) + '"'
            else:
                edit_info = '{"phone":"' + str(cell_phone[1]) + '"'
        else:
            pass
        logger.info('添加完phone的edit_info')
        logger.info(edit_info)
        if edit_info:
            edit_info = edit_info + '}'
        else:
            continue
        payload_edit = payload_template.format(customer_id, edit_info, '1781', random_id())
        logger.info('组合之后的payload片段')
        logger.info(payload)
        result = write_customer_info(s, payload_edit, headers)
        save_vcf(customer_info)
        wx[col_num] = customer_info['name']
        cell_d_num = "D" + str(i)
        wx[cell_d_num] = '是'
        cell_f_num = "F" + str(i)
        wx[cell_f_num] = '是'
        wb.save('2020.xlsx')


def random_id() -> object:
    return random.randrange(130000000, 799999999)


if __name__ == "__main__":
    main()
