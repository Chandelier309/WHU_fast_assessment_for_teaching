import re
import xlwt as ex
from selenium import webdriver
import time
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import WebDriverException

dcap = dict(DesiredCapabilities.PHANTOMJS)
dcap["phantomjs.page.settings.userAgent"] = \
    "Mozilla/5.0 (Linux; Android 5.1.1; Nexus 6 Build/LYZ28E) AppleWebKit/537.36 " \
    "(KHTML, like Gecko) Chrome/48.0.2564.23 Mobile Safari/537.36"

phan = 'D:/Python/Browsers/phantomjs-2.1.1-windows/bin/phantomjs.exe'
Mozi = 'D:/Mozilla Firefox/firefox.exe'
chro = 'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe'
url0 = 'http://210.42.121.134/stu/choose_PubLsn_list.jsp'
url = 'http://210.42.121.134/servlet/Login'
urllogin = 'http://210.42.121.134/index.jsp'
urlpicture = 'http://210.42.121.134/servlet/GenImg'

xpath = '//body/div/ul/li[3]/a'
xpath2 = "//div[1]/form/div/select[1]"
xpath3 = '/html/body/div[1]/form/div/select[2]'
xpath4 = '/html/body/div[1]/form/div/select[3]'
xpath5 = '/html/body/div[1]/form/input[2]'
css_nextpage = 'html body div.page_nav div.total_count a.navegate'
match = '第\d+/(\d+)页'

driver = webdriver.PhantomJS(executable_path=phan, desired_capabilities=dcap)

driver.get(urllogin)

cookies = driver.get_cookies()
for cookie in cookies:
    try:
        driver.add_cookie(cookie)
    except WebDriverException:
        print('cookies的添加出现异常,不过这没什么问题')

driver.save_screenshot('screenshot.png')
codeid = input('input your id: ')
codepw = input('input your password: ')
code = input('input the verifying code: ')

driver.find_element_by_id("qxftest").send_keys(codeid)
driver.find_element_by_name("pwd").send_keys(codepw)
driver.find_element_by_xpath("//*[@id='loginInputBox']/tbody/tr[3]/td[2]/div[1]/input").send_keys(code)
driver.find_element_by_id("loginBtn").click()

try:
    driver.find_element_by_id("btn2").click()
except NoSuchElementException:
    print('老铁,我估计你验证码输错了.再给你一次机会')
    driver.get(urllogin)

    cookies = driver.get_cookies()
    for cookie in cookies:
        try:
            driver.add_cookie(cookie)
        except WebDriverException:
            print('cookies的添加出现异常,不过这没什么问题')

    driver.save_screenshot('screenshot.png')
    code = input('input the verifying code: ')

    driver.find_element_by_id("qxftest").send_keys(codeid)
    driver.find_element_by_name("pwd").send_keys(codepw)

    driver.find_element_by_xpath("//*[@id='loginInputBox']/tbody/tr[3]/td[2]/div[1]/input").send_keys(code)
    driver.find_element_by_id("loginBtn").click()
    driver.find_element_by_id("btn2").click()


def frame_switch(name):
    driver.switch_to.frame(driver.find_element_by_name(name))


def frame_switch_id(idd):
    driver.switch_to.frame(driver.find_element_by_id(idd))


def choose_college(kk):
    colleges_list = driver.find_element_by_xpath(xpath2)
    diff_colleges = colleges_list.find_elements_by_tag_name("option")
    college_we_are_looking_for = diff_colleges[kk]
    return college_we_are_looking_for


def choose_academic(ii):
    acadd = driver.find_element_by_xpath(xpath3)
    academicss = acadd.find_elements_by_tag_name("option")
    academicc = academicss[ii]
    return academicc


def choose_subject(jj):
    subjj = driver.find_element_by_xpath(xpath4)
    subjectss = subjj.find_elements_by_tag_name("option")
    subjectt = subjectss[jj]
    return subjectt


def choose_grade(ll):
    gradd = driver.find_element_by_xpath('/html/body/div[1]/form/select')
    gradess = gradd.find_elements_by_tag_name('option')
    gradee = gradess[ll]
    return gradee


def class_infor(nn):
    classess = driver.find_elements_by_tag_name('tr')
    classe = classess[1:len(classess)]
    classinfor = classe[nn].find_elements_by_tag_name('td')
    return classinfor


def record(nn, mm, countt):
    classinfor = class_infor(nn)
    collesheet.write(countt, mm + 3, classinfor[mm].text)
    print(classinfor[mm].text)


def iterate(kk, ii, jj, ll, countt, nn):
    collegee = choose_college(kk)
    collegee.click()
    academicc = choose_academic(ii)
    academicc.click()
    subjectt = choose_subject(jj)
    subjectt.click()
    gradee = choose_grade(ll)
    gradee.click()
    collesheet.write(countt, 0, academicc.text)
    collesheet.write(countt, 1, subjectt.text)
    collesheet.write(countt, 2, gradee.text)

    classinfor = class_infor(nn)
    numclassinfor = len(classinfor)

    for m in range(numclassinfor):
        record(nn, m, count)

frame_switch('main')
driver.find_element_by_xpath(xpath).click()
frame_switch_id('iframe2')

coll = driver.find_element_by_xpath(xpath2)
colleges = coll.find_elements_by_tag_name("option")
numcol = len(colleges)

myexcel = ex.Workbook('E:\Python\schoolsystemscrape\Courses_专业课.xls')
collesheet = myexcel.add_sheet('一緒に帰ろう', cell_overwrite_ok=True)
count = 0

try:
    for k in range(numcol):
        choose_college(k).click()
        acad = driver.find_element_by_xpath(xpath3)
        academics = acad.find_elements_by_tag_name("option")
        numaca = len(academics)

        for i in range(numaca):
            college = choose_college(k)
            college.click()
            academic = choose_academic(i)
            academic.click()
            subj = driver.find_element_by_xpath(xpath4)
            subjects = subj.find_elements_by_tag_name("option")
            numsub = len(subjects)

            for j in range(1, numsub):
                college = choose_college(k)
                college.click()
                academic = choose_academic(i)
                academic.click()
                subject = choose_subject(j)
                subject.click()
                grad = driver.find_element_by_xpath('/html/body/div[1]/form/select')
                grades = grad.find_elements_by_tag_name('option')
                numgrad = len(grades)

                for l in range(5, numgrad):
                    college = choose_college(k)
                    college.click()
                    academic = choose_academic(i)
                    academic.click()
                    subject = choose_subject(j)
                    subject.click()
                    grade = choose_grade(l)
                    grade.click()
                    driver.find_element_by_xpath('/html/body/div[1]/form/input[2]').click()
                    navegate = driver.find_element_by_class_name('total_count')
                    determine_page_number = int(re.findall(match, navegate.text)[0])
                    if determine_page_number == 0:
                        break
                    else:
                        classes = driver.find_elements_by_tag_name('tr')
                        classes = classes[1:len(classes)]
                        numcla = len(classes)
                        for n in range(numcla):
                            iterate(k, i, j, l, count, n)
                            count += 1
                        if determine_page_number > 1:
                            for x in range(2, determine_page_number + 1):
                                pagelink = driver.find_elements_by_class_name('navegate')
                                for page in pagelink:
                                    if page.text == str(x):
                                        page.click()
                                        classes = driver.find_elements_by_tag_name('tr')
                                        classes = classes[1:len(classes)]
                                        numcla = len(classes)
                                        driver.save_screenshot('screenshot.png')
                                        for n in range(numcla):
                                            iterate(k, i, j, l, count, n)
                                            count += 1
                                        break
                    print(determine_page_number)
        time.sleep(1)
        college = choose_college(k)

finally:
    myexcel.save('Courses_专业课_2017.xls')
