import json

import xlrd


def read_file(filename):
    workbook = xlrd.open_workbook(filename)
    table1 = workbook.sheets()[0]
    rows1 = table1.nrows
    data_lang = []
    for v in range(0, rows1):
        values = table1.row_values(v)
        data_lang.append({
            "language": values[0],
            "zh-CN": values[1],
            "zh-TW": values[2],
            "en-US": values[3],
            "ja-JP": values[4],
            "ko-KR": values[5],
            "vi-VN": values[6],
            "id-ID": values[7],
            "th-TH": values[8],
        })
    table2 = workbook.sheets()[1]
    rows2 = table2.nrows
    data_table = []
    for v in range(1, rows2):
        values = table2.row_values(v)
        data_table.append({
            "language": values[0],
            "zh-CN": values[1],
            "zh-TW": values[2],
            "en-US": values[3],
            "ja-JP": values[4],
            "ko-KR": values[5],
            "vi-VN": values[6],
            "id-ID": values[7],
            "th-TH": values[8],
        })
    table3 = workbook.sheets()[2]
    rows3 = table3.nrows
    data_dotnet = []
    for v in range(1, rows3):
        values = table3.row_values(v)
        data_dotnet.append({
            "language": values[0],
            "string": values[1],
            "zh-CN": values[2],
            "zh-TW": values[3],
            "en-US": values[4],
            "ja-JP": values[5],
            "ko-KR": values[6],
            "vi-VN": values[7],
            "id-ID": values[8],
            "th-TH": values[9],
        })
    data = {
        "lang": data_lang,
        "table": data_table,
        "dotnet": data_dotnet,
    }
    return data


if __name__ == "__main__":
    lang = read_file("lang.xlsx")
    with open("lang.json", "w", encoding='utf-8') as f:
        f.write(json.dumps(lang["lang"], ensure_ascii=False, indent=2))
    with open("table.json", "w", encoding='utf-8') as f:
        f.write(json.dumps(lang["table"], ensure_ascii=False, indent=2))
    with open("dotnet.json", "w", encoding='utf-8') as f:
        f.write(json.dumps(lang["dotnet"], ensure_ascii=False, indent=2))
