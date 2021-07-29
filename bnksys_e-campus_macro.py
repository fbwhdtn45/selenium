import selenium
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
import time
import sys
import os
from tkinter import *
from tkinter import messagebox

window = Tk()
width = window.winfo_screenwidth()
height = window.winfo_screenheight()
id = ""
pw = ""
window.title("BNK")
window.geometry('{}x{}+{}+{}'.format(200, 100, int((width-200)/2), int((height-100)/2)))
window.resizable(False,False)
closed = False

def callback(event) :
    click()
    
window.bind('<Return>', callback)
label = Label(window,text="E-CAMPUS 수강 매크로")
label.grid(row=0,column=1)

label = Label(window,text="ID ")
label.grid(row=1,column=0)

id_text = Entry(window)
id_text.grid(row=1,column=1)

label = Label(window,text="PW ")
label.grid(row=2,column=0)

pw_text = Entry(window)
pw_text.grid(row=2,column=1)

def click() :
    global id, pw
    if (id_text.get() == "") or (pw_text.get() == "") :
        messagebox.showinfo("오류","아이디/비밀번호 입력")
    else :
        id = id_text.get()
        pw = pw_text.get()
        window.destroy()

def closing() :
    global closed
    window.destroy()
    closed = True

button = Button(window,text="확인",command=click)
button.grid(row=3,column=1,sticky=W+E+N+S)

id_text.focus()
window.protocol("WM_DELETE_WINDOW", closing)
window.mainloop()
session_time = 0

def letsgo() :
    global id,pw
    try :
        chromedriver_path = os.path.join(getattr(sys,'_MEIPASS'), "chromedriver.exe")
    except :
        chromedriver_path = os.path.join(os.path.abspath("."),"chromedriver.exe")
    try :
        driver = webdriver.Chrome(chromedriver_path)  # 크롬 드라이버로 크롬 켜기
        driver.implicitly_wait(1) # 로딩 기다리기
        url = 'https://bnksys.bnk-ecampus.co.kr/'        # 접속할 웹 사이트 주소 (e-campus)
        driver.get(url)
        driver.find_element_by_id("BSIS").find_element_by_tag_name('a').click() # 회사이름(시스템) 클릭
        driver.find_element_by_class_name("id").send_keys(id) # 아이디 입력
        driver.find_element_by_class_name("pw").send_keys(pw) # 비밀번호 입력
        driver.find_element_by_class_name("pw").send_keys(Keys.RETURN) # 로그인
        driver.find_element_by_class_name("btn-more").click() # 더보기 
        items = driver.find_element_by_class_name("tab_con_wrap").find_element_by_tag_name("tbody").find_elements_by_tag_name("tr")
    except : 
        return 0

    def go_to_learn(ele) :
        try :
            ele.find_element_by_tag_name("a").send_keys(Keys.CONTROL + "\n")
            driver.switch_to.window(driver.window_handles[1])
            time.sleep(1)

            items = driver.find_elements_by_class_name("tb_box")[1].find_element_by_tag_name('tbody').find_elements_by_tag_name('tr')
            items_cnt = len(items)
            i = 0
            while i < items_cnt :
                items = driver.find_elements_by_class_name("tb_box")[1].find_element_by_tag_name('tbody').find_elements_by_tag_name('tr')
                if learning(items[i]) == 0 :
                    return 0
                i += 1

            driver.close()
            driver.switch_to.window(driver.window_handles[0])
            time.sleep(1)
            return 1
        except : 
            return 0

    def learning(ele) :
        global session_time
        try :
            ele1 = ele.find_elements_by_tag_name('td')
            if ele1[3].text != '학습완료' :
                timer = ele1[2].find_element_by_class_name('lesson_time').text[:-1]
                time.sleep(1)
                ele1[4].find_element_by_tag_name('button').send_keys(Keys.CONTROL + "\n")
                driver.switch_to.window(driver.window_handles[2])
                time.sleep(1)
                
                # iframe 인 경우
                try :
                    # iframe 접근
                    driver.switch_to.frame(driver.find_element_by_tag_name("iframe"))
                    driver.find_element_by_xpath('//*[@id="player"]/div[7]/div[3]/button').click()
                    lesson_time = 60*(int(timer)+1)+30
                    session_time += lesson_time
                    if session_time > 3600 :
                        session_time = 0
                        return 0
                    time.sleep(lesson_time) # lesson_time
                    driver.switch_to.default_content()
                    driver.close()
                    try :
                        driver.switch_to.window(driver.window_handles[1])
                    except :
                        return 0
                    time.sleep(1)
                    return 1
                # 일반 플레이어인 경우
                except :
                    # 이어서 재생
                    try :
                        driver.find_element_by_xpath('//*[@id="kollus_player"]/div[10]/div/button[1]').click()                   
                    # 재생버튼 클릭
                    except :
                        # 재생버튼 클릭
                        driver.find_element_by_xpath('//*[@id="kollus_player"]/button').click()
                    # 동영상 시간만큼 기다리기
                    finally : 
                        lesson_time = 60*(int(timer)+1)+30
                        session_time += lesson_time
                        if session_time > 3600 :
                            session_time = 0
                            return 0
                        time.sleep(lesson_time) # lesson_time
                        driver.close()
                        try :
                            driver.switch_to.window(driver.window_handles[1])
                        except :
                            return 0
                        time.sleep(1)
                        return 1
        except :
            return 0

    try :
        for item in items :
            lecture = item.find_elements_by_tag_name("td")
            if lecture[3].text != "100.0 %" :
                if go_to_learn(lecture[5]) == 0 :
                    driver.quit()
                    return 0
        return 2
    except :
        return 0

while closed == False :
    try :
        imsi = letsgo()
    except :
        imsi = 0
    if imsi == 2 :
        break