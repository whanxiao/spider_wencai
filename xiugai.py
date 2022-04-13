import requests
import json
import pandas as pd
import execjs
import os

def main():
    l = get_name()
    for name in l:
        print(name)
        dic_data = get_dic(name)
        if dic_data == 0:
            continue
        else:
            export_excel(dic_data)
            print('end')





def get_name():
    l = []
    df = pd.read_excel("gupiao.xls")  # 读取项目名称列,不要列名
    df_li = df.values.tolist()
    result = []
    for s_li in df_li:
        result.append(s_li[0])
    for i in result:
        l.append(i)
    return l



def get_dic(i):
    with open('./aes.min.js', 'r') as f:
        jscontent = f.read()
    context = execjs.compile(jscontent)
    hexinv = context.call("v")

    url = "http://www.iwencai.com/customized/chart/get-robot-data"

    payload = json.dumps({
        "question": i,
        "perpage": 50,
        "page": 1,
        "secondary_intent": "stock",
        "log_info": "{\"input_type\":\"click\"}",
        "source": "Ths_iwencai_Xuangu",
        "version": "2.0",
        "query_area": "",
        "block_list": "",
        "add_info": "{\"urp\":{\"scene\":1,\"company\":1,\"business\":1},\"contentType\":\"json\",\"searchInfo\":true}"
    })
    headers = {
        'Accept': 'application/json, text/plain, */*',
        'Accept-Language': 'zh-CN,zh;q=0.9,ru;q=0.8,ja;q=0.7',
        'Cache-control': 'no-cache',
        'Connection': 'keep-alive',
        'Content-Type': 'application/json',
        'Cookie': 'other_uid=Ths_iwencai_Xuangu_3p4g5xi8tag6om0non8helyrcykqvpne; ta_random_userid=xlcblzk71e; cid=84c9f5bda5a7f8b8163e182baa4719251649577284; iwencaisearchquery=%E7%91%9E%E5%BA%86%E6%97%B6%E4%BB%A3%E6%96%B0%E8%83%BD%E6%BA%90%E7%A7%91%E6%8A%80%E6%9C%89%E9%99%90%E5%85%AC%E5%8F%B8; PHPSESSID=29186cb1e086080d27498be805afe0e5; cid=84c9f5bda5a7f8b8163e182baa4719251649577284; ComputerID=84c9f5bda5a7f8b8163e182baa4719251649577284; WafStatus=0; v=A_7wSd9OEPo8g0TeYvqPEW6kTx9FP8K5VAN2nagHasE8S5CBEM8SySSTxqt7',
        'Origin': 'http://www.iwencai.com',
        'Pragma': 'no-cache',
        'Referer': 'http://www.iwencai.com/unifiedwap/result?w=%E9%92%B1%E6%B1%9F%E6%91%A9%E6%89%98%20%E5%AD%90%E5%85%AC%E5%8F%B8&querytype=stock',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.75 Safari/537.36',
        'hexin-v': hexinv
    }

    response = requests.request("POST", url, headers=headers, data=payload)
    # print(response.text)
    b = json.loads(response.text)
    try:
        a = b['data']['answer'][0]["txt"][0]["content"]["components"][1]['data']
        return a
    except Exception:
        print('有异常')
        return 0


    # time.sleep(10)


def export_excel(dic_data):
    '''判读表是否存在  存在就直接读取，不存在就创建表
        读取完表 判断表是否为空，为空则从索引为0的行开始写
        不为空则从行数的下一列插入数据
        数据添加完毕后将数据写到表里'''

    if not os.path.exists("test.xlsx"):   # 判断 "test.xlsx" 是否存在，不存在就创建，存在就读取
        df = pd.DataFrame(columns=('股票简称','子公司名称', '参控比例', '子公司主营业务', '子公司净利润', '被参控公司是否报表合并', '被参控公司参控关系', '被参控公司是否上市'))
        df = df.set_index('子公司名称')
        df.to_excel("test.xlsx")

    df = pd.read_excel("test.xlsx")
    pf = pd.DataFrame(list(dic_data))

    pf = pf.copy()
    pf = pf.reindex(columns=['子公司名称', '股票简称', '参控比例', '子公司主营业务', '子公司净利润', '被参控公司是否报表合并', '被参控公司参控关系', '被参控公司是否上市'])
    # df = df.append(pf,ignore_index=True)
    df = pd.concat([df, pf], axis=0, join='outer', ignore_index=True, sort=False)
    df = df.set_index('子公司名称')

    df.to_excel("test.xlsx")


if __name__ == '__main__':
    main()
    print('保存完毕')