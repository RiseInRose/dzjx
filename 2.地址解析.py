import re
import requests
import xlwt
import xlrd
import time

def findadr(excel_name,hang):
    wb = xlwt.Workbook()
    sheet = wb.add_sheet('sheet1')
    bk = xlrd.open_workbook(excel_name)
    shxrange = range(bk.nsheets)
    try:
        sh = bk.sheet_by_name("sheet1")
    except:
        print("no sheet in %s named Sheet1" % excel_name)
    # 获取行数
    nrows = sh.nrows
    # 获取列数
    ncols = sh.ncols
    # print("nrows %d, ncols %d" % (nrows, ncols))

    row_list = []
    # 获取各行数据
    for i in range(1, nrows):
        row_data = sh.row_values(i)
        row_list.append(row_data)
        # print(row_list[2][1])
        # 按照列表，打开每一个网址，获取分享链接
    for i in range(2, nrows-1):
        url = row_list[i][1]
        # print(url)
        try:
            p = r'(page/./././)(.+)(.html)'
            vid = re.search(p, str(url)).group(2)
            sheet.write(i, hang+1, vid)
            # 下面是电脑端解析出真实地址。现在不用。
            # print(vid)
            crack_url = "http://vv.video.qq.com/geturl?vid=" + vid + "&otype=xml&platform=1&ran=0%2E9652906153351068"
            headers = {'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
                       'Accept-Encoding': 'gzip, deflate, compress',
                       'Accept-Language': 'en-us;q=0.5,en;q=0.3',
                       'Cache-Control': 'max-age=0',
                       'Connection': 'keep-alive',
                       'User-Agent': 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:22.0) Gecko/20100101 Firefox/22.0'}
            r = requests.get(url=crack_url, headers=headers)
            # print(r.content)
            p2 = r'(type><url>)(.+)(</url><urlbk>)'
            try:
                video_url = re.search(p2, str(r.content)).group(2)
                sheet.write(i, hang, video_url)
                print(video_url)
                print('the %d is ok' % (i - 1))
            except:
                print('match failure')


        except:
            print('there is something wrong ')
        wb.save(excel_name3)

# Nowtime = time.strftime('%Y-%m-%d',time.localtime(time.time()))
# excel_name = str(Nowtime)+'腾讯'+'.xls'
# excel_name2 = str(Nowtime)+'腾讯分享地址'+'.xls'
# excel_name3 = str(Nowtime)+'腾讯解析地址'+'.xls'
excel_name ='腾讯'+'.xls'
excel_name3 = '腾讯解析地址'+'.xls'
findadr(excel_name,1)
print("done")

# p =r'(&vid=)(.+)(&auto)'



