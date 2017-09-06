# coding:utf-8

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import json
from datetime import datetime
import datetime as dt
import calendar
from openpyxl import Workbook
import threading
from selenium.webdriver.support.ui import Select


def dataGrab(year,month):
    web.switch_to.frame('left')
    selectY = Select(web.find_element_by_name('YY'))
    selectY.select_by_value(year)
    selectM = Select(web.find_element_by_name('MM'))
    selectM.select_by_value(month)
    web.find_element_by_name('NEXT').click()

    web.switch_to.default_content()
    web.switch_to.frame('right')
    selectClass = Select(web.find_element_by_name('STR'))
    selectClass.select_by_value('BB104')
    web.find_element_by_name('Query').click()

    web.switch_to.default_content()
    web.switch_to.frame('bottom')

    tmpList = []
    check = [2, 5, 8, 11, 14]
    try:
        test = str(".//*[@class='A9']/tbody/tr[17]/td[3]")
        while web.find_element_by_xpath(test).text:
            check.extend(17)
    except:
        pass



    for i in range(3,10,1):
        for j in check:
            dicMo = {
                'Subject': '',
                'Start Date': '',
                'Start Time': '',
                'End Date': '',
                'End Time': '',
            }
            dicNo = {
                'Subject': '',
                'Start Date': '',
                'Start Time': '',
                'End Date': '',
                'End Time': '',
            }
            dicNi = {
                'Subject': '',
                'Start Date': '',
                'Start Time': '',
                'End Date': '',
                'End Time': '',
            }


            loc_str1= str((".//*[@class='A9']/tbody/tr[{}]/td[{}]").format(j,i))
            loc_str2= str((".//*[@class='A9']/tbody/tr[{}]/td[{}]").format(j+1,i-1))
            loc_str3= str((".//*[@class='A9']/tbody/tr[{}]/td[{}]").format(j+2,i-1))


            morning = web.find_element_by_xpath(loc_str1).text
            if len(morning) != 0:
                dayNum = morning.split('／')[1].split('\n')[0]
                date = year + '-' + month + '-' + dayNum
                if len(morning) > 5:
                    if len(dayNum) ==1:
                        morningSubject = morning.split('／')[1][2:].replace('\n','-')
                    else:
                        morningSubject = morning.split('／')[1][3:].replace('\n','-')
                    dicMo['Start Date'] = date
                    dicMo['End Date'] = date
                    dicMo['Subject'] = morningSubject
                    dicMo['Start Time'] = '9:00 AM'
                    dicMo['End Time'] = '12:00 PM'
                    print(dicMo)
                    tmpList.append(dicMo)


                noon = web.find_element_by_xpath(loc_str2).text
                if len(noon) != 0:
                    dicNo['Start Date'] = date
                    dicNo['End Date'] = date
                    dicNo['Subject']=noon.replace('\n','-')
                    dicNo['Start Time']='13:30 PM'
                    dicNo['End Time']='16:30 PM'
                    print(dicNo)
                    tmpList.append(dicNo)

                night = web.find_element_by_xpath(loc_str3).text
                if len(night) != 0:
                    dicNi['Start Date'] = date
                    dicNi['End Date'] = date
                    dicNi['Subject']=night.replace('\n','-')
                    dicNi['Start Time']='18:00 PM'
                    dicNi['End Time']='21:00 PM'
                    print(dicNi)
                    tmpList.append(dicNi)
    web.switch_to.default_content()
    return tmpList



calendarList = []


# driver_path = 'D:\webdriver\chromedriver'
driver_path = '/Users/yunhan/driver/chromedriver'
web = webdriver.Chrome(driver_path)
web.get('http://140.115.236.11/')







# calendarList.extend(dataGrab('2017','9'))
# calendarList.extend(dataGrab('2017','10'))
# calendarList.extend(dataGrab('2017','11'))
calendarList.extend(dataGrab('2017','12'))
calendarList.extend(dataGrab('2018','1'))
calendarList.extend(dataGrab('2018','2'))




print(calendarList)
web.quit()




