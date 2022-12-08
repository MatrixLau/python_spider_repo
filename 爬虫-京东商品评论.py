# coding=utf-8
import requests
import json
import time
import xlwt
import xlutils.copy
import xlrd

def start(page):
    # score表示为评论的数据类型 如 0:全部评价 1:好评 2:中评 3:差评 5:追加评价 （不确定）
    score = '0'
    # producitid 商品id 
    ### 如何获取？
    ### 1.京东打开商品详情页，复制url   如：https://item.jd.com/100021965322.html
    ### 2.在url中找到productId=后面的数字或者如上面的100021965322.html取后面的数字
    productId = '100021965322'

    url = 'https://club.jd.com/comment/productPageComments.action?&productId='+ productId +'&score='+ score +'&sortType=0&page='+ str(page) +'&pageSize=10&isShadowSku=0&fold=1'
    headers= {
        "User-Agent": "Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.198 Mobile Safari/537.36"
    }
    time.sleep(2)
    test = requests.get(url=url, headers= headers)
    data = json.loads(test.text)
    return data

def parse(data):   # 解析页面

    items = data['comments']
    for i in items:
        yield (
            i['referenceName'],   #商品名
            i['nickname'],   #用户昵称
            i['id'],   #用户id
            i['content'],   #评论内容
            i['creationTime'],   #评论时间
            i['score']   #评论分数
        )

def excel(items):   #第一次写入
    newTable = "dataFetched.xls"   #创建文件
    wb = xlwt.Workbook("encoding='utf-8")

    ws = wb.add_sheet('sheet1')   #创建表
    headDate = ['itemName', 'nickname', 'id', 'content', 'creationTime','score']   #定义标题
    for i in range(0,6):   #for循环遍历写入 对应解析页面取出的数据的标题数
        ws.write(0, i, headDate[i], xlwt.easyxf('font: bold on'))

    index = 1   #行数

    for data in items:   #items是十条数据 data是其中一条
        for i in range(0,6):   #列数 对应解析页面取出的数据的标题数

            print(data[i])
            ws.write(index, i, data[i])   #行 列 数据（一条一条自己写入）
        print('______________________')
        index += 1   #等上一行写完了 在继续追加行数
        wb.save(newTable)

def another(items, j):   #如果不是第一次写入 以后的就是追加数据了 需要另一个函数
    index = (j-1) * 10 + 1   #这里是 每次写入都从11 21 31..等开始 所以我才传入数据 代表着从哪里开始写入
    data = xlrd.open_workbook('dataFetched.xls')
    ws = xlutils.copy.copy(data)
    table = ws.get_sheet(0)   # 进入表
    for test in items:
        for i in range(0,6):   #跟excel同理 对应解析页面取出的数据的标题数
            print(test[i])

            table.write(index, i, test[i])   # 只要分配好 自己写入
        print('_______________________')

        index += 1
        ws.save('dataFetched.xls')



def main():
    j = 1   #页面数
    judge = True   #判断写入是否为第一次

    for i in range(0, 100):
        time.sleep(1.5)   #加延迟以免被封IP
        first = start(j)
        test = parse(first)

        if judge:
            excel(test)
            judge = False
        else:
            another(test, j)
        print(str(j) + ' page(s) fetched\n')
        j = j + 1


if __name__ == '__main__':
    main()
