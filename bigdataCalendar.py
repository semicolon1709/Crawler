# coding:utf-8

from selenium import webdriver
from selenium.webdriver.support.ui import Select
import csv
from datetime import datetime
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
            check.append(17)
            break
    except:
        pass



    for i in range(3,10,1):
        for j in check:

            loc_str1= str((".//*[@class='A9']/tbody/tr[{}]/td[{}]").format(j,i))
            loc_str2= str((".//*[@class='A9']/tbody/tr[{}]/td[{}]").format(j+1,i-1))
            loc_str3= str((".//*[@class='A9']/tbody/tr[{}]/td[{}]").format(j+2,i-1))


            morning = web.find_element_by_xpath(loc_str1).text
            if len(morning) != 0:
                monthNum = morning.split('／')[0]
                dayNum = morning.split('／')[1].split('\n')[0]
                date = year + '-' + monthNum + '-' + dayNum
                if len(morning) > 5:
                    if len(dayNum) ==1:
                        morningSubject = morning.split('／')[1][2:].replace('\n','-')
                    else:
                        morningSubject = morning.split('／')[1][3:].replace('\n','-')
                    if morningSubject[0] == '-':
                        morningSubject = morningSubject[1:]
                    tupleMo = (morningSubject,date,'9:00 AM',date,'12:00 PM')
                    print(tupleMo)
                    tmpList.append(tupleMo)


                noon = web.find_element_by_xpath(loc_str2).text
                if len(noon) != 0:
                    noonSubject = noon.replace('\n','-')
                    if noonSubject[0] == '-':
                        noonSubject = noonSubject[1:]
                    tupleNo = (noonSubject, date, '13:30 PM', date, '16:30 PM')
                    print(tupleNo)
                    tmpList.append(tupleNo)

                night = web.find_element_by_xpath(loc_str3).text
                if len(night) != 0:
                    nightSubject = night.replace('\n', '-')
                    if nightSubject[0] == '-':
                        nightSubject = nightSubject[1:]
                    tupleNi = (nightSubject, date, '18:00 PM', date, '21:00 PM')
                    print(tupleNi)
                    tmpList.append(tupleNi)
    web.switch_to.default_content()
    return tmpList


calendarList = []
threads=[]
start = input('輸入開始年月，如"2017-9"')
end = input(('輸入結束年月，如"2018-2"'))
startYear = int(start.split('-')[0])
startMonth = int(start.split('-')[1])
endYear = int(end.split('-')[0])
endMonth = int(end.split('-')[1])



driver_path = 'D:\webdriver\chromedriver'
# driver_path = '/Users/yunhan/driver/chromedriver'
web = webdriver.Chrome(driver_path)
# driver_path = 'D:\webdriver\phantomjs\\bin\phantomjs'
# web = webdriver.PhantomJS(driver_path)
web.get('http://140.115.236.11/')

endTmpMonth = 13
s1 = datetime.now()
for i in range(startYear, endYear + 1, 1):
    if i != startYear:
        startMonth = 1
    if i == endYear:
        endTmpMonth = endMonth+1
    for j in range(startMonth, endTmpMonth, 1):
        calendarList.extend(dataGrab(str(i),str(j)))
s2 = datetime.now()

calendarSet=set(calendarList)
print(calendarSet)
print(str(s2-s1))
headers  =  ['Subject', 'Start Date', 'Start Time', 'End Date', 'End Time']


with open('calendar.csv', 'w', newline='') as f:
    f_csv = csv.writer(f)
    f_csv.writerow(headers)
    f_csv.writerows(calendarSet)
    f.close()
web.quit()
