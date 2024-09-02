import pandas as pd
import re


df = pd.read_excel("报名2.xlsx")


def process_str_to_list(input_string):
    # 使用正则表达式替换多个空格和制表符为一个空格
    output_string = re.sub(r'\s+', ' ', input_string)

    # 按照 "#题目" 切分字符串
    items = output_string.split("#题目")

    # 去掉空串
    items = [("#题目 " + item.strip()) for item in items if item.strip()]
    dic = {}
    # 输出结果
    for item in items:
        # print(item)
        pattern = r'#题目\s+(.*?)\s+问题.*?用户填写\s+(.*?)(?=\s*#|$)'
        matches = re.findall(pattern, item)

        # 输出结果
        for match in matches:
            dic[match[0]] = match[1]
    return dic

ls = []
for i in range(0, len(df)):
    # 从第5列开始到最后一列合并为字符串 前4列为个人信息
    dic_ = {}
    result = ' '.join(df.iloc[i, 5:].astype(str)).replace('NaN','').replace("nan",'')
    dic_.update(dict(df.iloc[i,0:4]))
    dic_.update(process_str_to_list(result))
    ls.append(dic_)

df_new = pd.DataFrame(ls)

df_new.to_excel("报名2_new.xlsx")