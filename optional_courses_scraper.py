import requests
import re
import xlwt as x
import time

url0 = 'http://210.42.121.134/stu/choose_PubLsn_list.jsp'
urllogin = 'http://210.42.121.134/servlet/Login'
user_agent = "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:51.0) Gecko/20100101 Firefox/51.0"
urlpicture = 'http://210.42.121.134/servlet/GenImg'
headers={"user-agent":user_agent}
cookies = {}

codeid = input('input your id: ')
codepw = input('input your password: ')


def login():
    respicture = requests.get(urlpicture, headers=headers)
    cookies['JSESSIONID'] = respicture.cookies.get('JSESSIONID')
    with open('code.jpg', 'wb') as img:
        img.write(respicture.content)
    code = input('input the code: ')
    form = {'id': codeid, "pwd": codepw, "xdvfb": code}  # 登录所需信息
    requests.post(urllogin, headers=headers, data=form, cookies=cookies)  # 模拟登录教务系统


def get_code(url):
    req = requests.get(url=url, headers=headers, cookies=cookies)  # 向url发送get请求，获取整个页面
    return req  # 返回页面的源代码

    # match = '<tr >\n    \n    \n    <td >(.+)</td>\n    \n     <td>(.+)</td>\n     <td><font color="#FF0000">(.+)</font>/(.+)</td>\n    <td >(.+)</td> \n    <td>(.+)</td>\n    <td>(.+)</td>\n    <td >(.+)</td>\n    <td >(.+)</td>\n     <td >(.+)</td>\n    \n    \n   \n    \n \n    <td >\n\t    <div class="overflow">\n\t    \t(.+)\n(.+)\n\t\t</div>\n\t</td>\n       <td >\n\n    <div class="overflow">\n\t    \t\n\t    \t \t (.+)\n\t    \t \t \n\t</div>\n   </td>'
    # match = '<td ?>(.+)</td>\r\n.+\r\n.+<td ?>(.+)</td>\r\n.+<td ?><.+>(.+)<.+>/(.+)</td>\r\n.+<td ?>(.+)</td> ?\r\n.+<td ?>(.+)</td>\r\n.+<td ?>(.+)</td>\r\n.+<td ?>(.+)</td>\r\n.+<td ?>(.+)</td>\r\n.+<td ?>(.+)</td>\r\n.+\r\n.+\r\n.+\r\n.+\r\n ?\r\n.+<td ?>\r\n\t.+<.+>\r\n\t.+\t(.+)\n(.+)\r\n\t\t</div>\r\n\t</td>\r\n.+<td ?>\r\n\r\n.+<.+>\r\n\t.+\t\r\n\t.+\t ?(.+)\r\n.+\r\n\t</div>'


def grab():
    file = x.Workbook('E:\Python\schoolsystemscrape\Courses')
    table = file.add_sheet('PublicOptional_398')

    match = '<td ?>(.+)</td>'
    match02 = '\t \t (.+)\n?\r\n'  # 课程类型
    match03 = '<div class="overflow">\r\n\t    \t(.+);?\n(.+)\r\n\t\t</div>|<div class="overflow">\r\n\t    \t\r\n\t\t</div>'  # 上课时间地点
    # (.+)\n(.+)\r\n\t\t</div>\r\n\t</td>\r\n.+<td ?>\r\n\r\n.+<.+>\r\n\t.+\t\r\n\t.+
    codes = get_code(url0).text
    massage = re.findall(match, codes)
    massage2 = re.findall(match02, codes)
    massage3 = re.findall(match03, codes)
    fail = re.findall('验证码',codes)

    if len(fail):
        raise Exception('need login')

    for i in range(2, 28):
        codes = get_code(
            'http://210.42.121.134/stu/choose_PubLsn_list.jsp?XiaoQu=0&credit=0&keyword=&pageNum=' + str(i)).text
        Add = re.findall(match, codes)
        Add02 = re.findall(match02, codes)
        Add03 = re.findall(match03, codes)
        massage += Add
        massage2 += Add02
        massage3 += Add03

    massage = [x for x in massage if 'onclick' not in x]
    massage = [' ' if '&nbsp;' in x else x for x in massage]
    # leng = len(massage[0])

    for i in range(2, len(massage), 9):
        t = massage[i]
        t = t.replace('</font>', '')
        massage[i] = t.replace('<font color="#FF0000">', '')

    for i in range(len(massage2)):
        for j in range(9):
            table.write(i, j, massage[i * 9 + j])
        try:
            table.write(i, 10, massage2[i])
        except:
            break
        try:
            table.write(i, 9, massage3[i][0] + massage3[i][1])
        except:
            table.write(i, 9, massage3[i][0])
        else:
            continue

    file.save('E:\Python\schoolsystemscrape\Courses.xls')

try:
    grab()
except:
    login()
    grab()
