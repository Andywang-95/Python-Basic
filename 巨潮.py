import requests

# 功能一
def search():
    while True:
        keyword = input('A股資料查詢\n\n請輸入要搜索的關鍵詞： ')
        params = {
            'keyWord':keyword,
            'maxNum':10
        }
        search_res = requests.post('http://www.cninfo.com.cn/new/information/topSearch/query',params=params)
        search_data = search_res.json()
        if search_data == []:
            print('無搜索結果！\n')
            continue
        index = 1
        search_dict = {}
        for options in search_data:
            if options['category'] == 'A股':
                search_dict[index] = options
                print('【{}】 {} {} 代碼：{}'.format(index,options['category'],options['zwjc'],options['code']))
                index += 1
        print('【0】 重新搜索關鍵字...\n')
        while True:
            choice = int(input('請輸入要查看的序號： '))
            if choice in range(1,len(search_dict)+1):
                print('\n請稍後，正在跳轉至【{}】\n'.format(search_dict[choice]['zwjc']))
                return search_dict[choice]
            elif choice == 0:
                break
            else:
                
                print('序號不存在!')
                continue

# 功能二
def select(target_dict):
    if target_dict['code'][0] == '6':
        column = 'sse'
        plate = 'sh'
    else:
        column = 'szse'
        plate = 'sz'
    category_dict = {
        '1': 'category_ndbg_szsh;',
        '2': 'category_bndbg_szsh;',
        '3': 'category_rcjy_szsh;'
    }
    choice = ''
    while True:
        n = input('請輸入你想要查看的項目：\n1.年報  2.半年報  3.日常經營\n**如有多个项目，请用【Enter】隔开，完成后请输入【n】前往下一步**\n')
        if n != 'n' and int(n) < 4:
            choice += category_dict[n]
            continue
        break   
    date1 = input('請依照格式輸入查詢起始日期（2021-01-01）：\n')
    date2 = input('請依照格式輸入查詢截至日期（2021-01-01）：\n')
    num = 1
    pdf_list = []
    while True:
        select_url = 'http://www.cninfo.com.cn/new/hisAnnouncement/query'
        Form_data = {
            'stock': target_dict['code']+','+target_dict['orgId'],
            'tabName': 'fulltext',
            'pageSize': '30',
            'pageNum': num,
            'column': column,
            'category': choice,
            'plate': plate,
            'seDate': date1+'~'+date2,
            'searchkey': '',
            'secid': '',
            'sortName':'' ,
            'sortType':'' ,
            'isHLtitle': 'true'
        }
        res = requests.post(select_url,data=Form_data)
        if res.json()['announcements'] is None:
            if pdf_list == []:
                return '無搜索結果'
            else:
                return pdf_list
        for i in res.json()['announcements']:
            pdf_list.append([i['announcementTitle']+'.'+i['adjunctType'],'http://static.cninfo.com.cn/'+i['adjunctUrl']])
        num += 1

# # 功能三
def download(list_pdf):
    for pdf in list_pdf:
        res_pdf = requests.get(pdf[1])
        d_pdf = res_pdf.content
        with open(pdf[0],'wb') as f:
            f.write(d_pdf)
            print('已完成《{}》的下载'.format(pdf[0].split('.')[0]))

# 功能整合
def main():
    # 搜索股票，获取股票id信息等
    information = search()
    # 输入各个参数，筛选报告
    final_list = select(information)
    # 下载报告
    finish= download(final_list)
main()


