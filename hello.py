import json
import xlwt

cn_value = []
tw_value = []
en_value = []
cn_key = {}
tw_key = {}
en_key = {}


def find_str(_json, n, s, value, key):
    if isinstance(_json, dict):
        for index in _json:
            if isinstance(_json[index], str):
                value.append(_json[index])
                key[s + "." + index] = len(value)
            else:
                find_str(_json[index], n + 1, s + "." + index, value, key)
    else:
        for index in range(len(_json)):
            value.append(_json[index])
            key[s + "[" + str(index) + "]"] = len(value)


def read_file(file):
    with open(file, 'r', encoding='utf8') as fr:
        data = json.load(fr)  # 用json中的load方法，将json串转换成字典
    return data


tw = read_file('./zh-tw.json')
cn = read_file('./zh-cn.json')
en = read_file('./en-us.json')
find_str(cn, 0, "$i18n", cn_value, cn_key)
find_str(tw, 0, "$i18n", tw_value, tw_key)
find_str(en, 0, "$i18n", en_value, en_key)
book = xlwt.Workbook()  # 创建一个excel对象
sheet = book.add_sheet('Sheet1', cell_overwrite_ok=True)  # 添加一个sheet页
title = ["string_id", "ZH-CN", "ZH-TW", "English"]
for i in range(len(title)):
    sheet.write(0, i, title[i])  # 将title数组中的字段写入到0行i列中
for i in cn_key:
    sheet.write(cn_key[i] + 1, 0, i)
    sheet.write(cn_key[i] + 1, 1, cn_value[cn_key[i]])
    sheet.write(cn_key[i] + 1, 2, tw_value[tw_key[i]])
    sheet.write(cn_key[i] + 1, 3, en_value[en_key[i]])
book.save('demo.xls')
