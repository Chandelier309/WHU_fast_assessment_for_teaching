from selenium.common.exceptions import NoSuchElementException
from selenium import webdriver
import time


def main():
    chrome = 'D:\Python\Browsers\chromedriver_win32\chromedriver.exe'
    url_login = 'http://210.42.121.134/index.jsp'
    xpath = '//*[@id="btn4"]'
    Fail_to_login = True
    k = 2

    driver = webdriver.Chrome(executable_path=chrome)

    def login():
        driver.get(url_login)
        driver.save_screenshot('screenshot.png')
        code = input('input the verifying code: ')
        username = input('用户名\n')
        cipher = input('密码\n')
        driver.find_element_by_id("qxftest").send_keys(username)
        driver.find_element_by_name("pwd").send_keys(cipher)
        driver.find_element_by_xpath("//*[@id='loginInputBox']/tbody/tr[3]/td[2]/div[1]/input").send_keys(code)
        driver.find_element_by_id("loginBtn").click()
        driver.find_element_by_xpath(xpath).click()

    def frame_switch(name):
        driver.switch_to.frame(driver.find_element_by_name(name))

    def frame_switch_id(idd):
        driver.switch_to.frame(driver.find_element_by_id(idd))

    def xpath_class(number_2_11):
        xpath_out = '/html/body/table/tbody/tr[' + str(number_2_11) + ']/td[7]/input[1]'
        return xpath_out

    def xpath_aspect(number, score, page, two_or_three):
        if page == 1:
            xpath_out = '/html/body/form/div/div[1]/table/tbody/tr[' \
                        + str(number) + ']/td[' + str(two_or_three) \
                        + ']/div/div[2]/img[' \
                        + str(score) + ']'
        else:
            xpath_out = '/html/body/form/div/div[2]/table[1]/tbody/tr[' \
                        + str(number) + ']/td[' + str(two_or_three) \
                        + ']/div/div[2]/img[' \
                        + str(score) + ']'
        return xpath_out

    def send_score(index, score, page):
        try:
            driver.find_element_by_xpath(xpath_aspect(index, score, page, 2)).click()
        except NoSuchElementException:
            driver.find_element_by_xpath(xpath_aspect(index, score, page, 3)).click()

    def assess(index_teacher):
        course = driver.find_element_by_xpath(xpath_class(index_teacher))
        course.click()
        driver.switch_to.default_content()
        driver.switch_to.frame('overLayFrame')
        info = driver.find_element_by_xpath('/html/body/table/tbody/tr[2]')
        print(info.text)
        scores = input('给Ta个分吧\n')
        for j in range(2, 11):
            send_score(j, scores, 1)
        driver.find_element_by_partial_link_text('课程评价').click()
        time.sleep(2)
        for j in range(2, 6):
            try:
                send_score(j, scores, 2)
            except NoSuchElementException:
                pass
        time.sleep(1)
        word = input('给Ta个评价吧\n')
        if word == '':
            word = '祝老师身体健康,万事如意,阖家幸福!'
        driver.find_element_by_xpath('/html/body/form/div/div[2]/table[2]/tbody/tr[2]/td/textarea').send_keys(word)
        driver.find_element_by_xpath('/html/body/form/div/div[2]/table[2]/tbody/tr[3]/th/input[14]').click()
        driver.find_element_by_partial_link_text('课程评价').click()
        verify = input('又是验证码!\n')
        driver.find_element_by_name('xdvfb').send_keys(verify)
        driver.find_element_by_name('Submit').click()

    while Fail_to_login:
        try:
            login()
            Fail_to_login = False
        except NoSuchElementException:
            print('老铁,我估计你验证码输错了.再给你一次机会')
    frame_switch('main')
    frame_switch_id('iframe0')

    while True:
        try:
            driver.find_element_by_xpath(xpath_class(k))
        except NoSuchElementException:
            break
        time.sleep(1)
        try:
            assess(k)
        except NoSuchElementException:
            pass
        time.sleep(1)
        k += 1
        driver.get('http://210.42.121.134/stu/stu_index.jsp')
        frame_switch('main')
        frame_switch_id('iframe0')

    driver.close()

if __name__ == '__main__':
    main()
