import json
import xlwt

cn_value = []
tw_value = []
en_value = []
key = []


def find_str(_json, n, s, value, add):
    for index in _json:
        if isinstance(_json[index], str):
            value.append(_json[index])
            if add:
                key.append(s + "." + index)
        else:
            find_str(_json[index], n + 1, s + "." + index, value, add)


def read_file(file):
    with open(file, 'r', encoding='utf8') as fr:
        data = json.load(fr)  # 用json中的load方法，将json串转换成字典
    return data


tw = read_file('./zh-tw.json')
cn = read_file('./zh-cn.json')
en = read_file('./en-us.json')
find_str(cn, 0, "lang", cn_value, True)
find_str(tw, 0, "lang", tw_value, False)
find_str(en, 0, "lang", en_value, False)
book = xlwt.Workbook()  # 创建一个excel对象
sheet = book.add_sheet('Sheet1', cell_overwrite_ok=True)  # 添加一个sheet页
title = ["string_id", "ZH-CN", "ZH-TW", "English"]
for i in range(len(title)):
    sheet.write(0, i, title[i])  # 将title数组中的字段写入到0行i列中
for i in range(len(key)):
    sheet.write(i + 1, 0, key[i])
    sheet.write(i + 1, 1, cn_value[i])
    sheet.write(i + 1, 2, tw_value[i])
    sheet.write(i + 1, 3, en_value[i])
book.save('demo.xls')
