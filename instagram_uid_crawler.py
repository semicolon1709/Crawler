#!/usr/bin/python
# -*- coding: utf-8 -*-

import time
from datetime import datetime
from threading import Thread
from Queue import Queue
from Tkinter import *
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

#  程式一執行產生輸入帳號密碼介面
win=Tk()

#全域帳號、密碼、THREAD數
aUser = None
aPassword = None
NUM_THREADS = None

# cancel按鈕的事件
def FormClose():
    global aUser
    aUser = ""
    global aPassword
    aPassword = ""
    win.destroy()
    exit()
# ok按鈕的事件
def FormOK():
    global aUser
    aUser = edit_user.get()
    global aPassword
    aPassword = edit_password.get()
    global NUM_THREADS
    NUM_THREADS = int(edit_thread.get())
    win.destroy()
#帳號密碼介面設定
win.title("Tk GUI")
win.geometry("300x200")
    #USER
label_user=Label(win, text="User : ")
edit_user=Entry(win, text="")
label_user.grid(column=0, row=0)
edit_user.grid(column=1, row=0, columnspan=6)
    #PASSWORD
label_password=Label(win, text="Password : ")
edit_password=Entry(win, text="", show="*")
label_password.grid(column=0, row=1)
edit_password.grid(column=1, row=1, columnspan=6)
    #Thread Num
label_thread=Label(win, text="Thread Number : ")
edit_thread=Entry(win)
edit_thread.insert(END, '4')
label_thread.grid(column=0, row=2)
edit_thread.grid(column=1, row=2, columnspan=6)
    #BUTTON-連接事件
button_cancel = Button(win, text="Cancel",  width="10", command=FormClose)
button_ok = Button(win, text="OK",width="10", command=FormOK)
button_cancel.grid(column=0, row=3)
button_ok.grid(column=1, row=3)
win.mainloop()

# 爬蟲程式獨立執行FUNCTION
def crawler(line):
    driver = webdriver.Remote("http://localhost:9515", webdriver.DesiredCapabilities.CHROME)
    driver.get("https://www.instagram.com/")
    # driver.maximize_window()
    driver.find_element_by_class_name("_fcn8k").click()
    driver.find_element_by_name("username").send_keys(aUser)
    driver.find_element_by_name("password").send_keys(aPassword)
    driver.find_element_by_class_name("_o0442").click()
    time.sleep(2)
    ActionChains(driver).move_by_offset(50, 0).click().perform()

    driver.find_element_by_class_name("_9x5sw").send_keys(line.decode("utf-8"))
    time.sleep(2)
    ActionChains(driver).send_keys(Keys.ENTER).perform()
    time.sleep(2)

    hotArticles = driver.find_elements_by_class_name("_t5r8b")
    article = hotArticles.__getitem__(0)
    article.click()
    time.sleep(1)

    authList=[]
    try:
        while (TRUE):
            #auth = driver.find_element_by_class_name("_4zhc5").text
            tags = driver.find_element_by_class_name("_nk46a").find_elements_by_tag_name("a")
            tmpTag = ""
            for i,tag in enumerate(tags):
                if i == 0:
                    auth = tag.text;
                else:
                    tmpTag += tag.text+",";
            auth_tag = auth + " - " + tmpTag
            authList.append(auth_tag)
            driver.find_element_by_class_name("coreSpriteRightPaginationArrow").click()
            time.sleep(0.8)
    except NoSuchElementException as e1:
        print(line.strip("\n")+" Finish !!")

    finally:
        filename = unicode("E:/" + line, 'utf8').strip("\n")+"_"+str(len(authList))+".txt"
        with open(filename, "w") as f_id:
            for aItem in authList:
                print aItem
                f_id.write(aItem.encode("utf-8")+ "\n")
            f_id.close()
    driver.close()

# 從QUEUE中 將資料逐一拿出，接著跑爬蟲
def worker():
    while not queue.empty():
        name=queue.get()
        crawler(name)

queue = Queue()
with open("E:/restaurants2.txt","r") as f_rests:
    # 將待搜尋的資料從檔案中放進QUEUE
    for line in f_rests.readlines():
        queue.put(line)

s1 = datetime.now()
# 設定THREAD數量及執行的FUNCTION
threads = map(lambda i: Thread(target=worker), xrange(NUM_THREADS))
#啟動THREAD
map(lambda th: th.start(), threads)
map(lambda th: th.join(), threads)
s2 = datetime.now()

print "All  Finish "+str(s2-s1)+"!!"


