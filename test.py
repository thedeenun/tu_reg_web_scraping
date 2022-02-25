import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By

PATH = './chromedriver'
driver = webdriver.Chrome(PATH)

url = 'https://web.reg.tu.ac.th/registrar/class_info.asp?lang=th'
driver.get(url)

def select_Faculty(faculty):
    select = driver.find_element_by_tag_name("body > table > tbody > tr.ContentBody > td:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(2) > font:nth-child(2) > select")
    select.click()
    faculty_selected = driver.find_element_by_tag_name("body > table > tbody > tr.ContentBody > td:nth-child(2) > table > tbody > tr:nth-child(4) > td:nth-child(2) > font:nth-child(2) > select > option:nth-child("+str(faculty)+")")
    faculty_selected.click()

def select_year(term):
    select = driver.find_element_by_tag_name("body > table > tbody > tr.ContentBody > td:nth-child(2) > table > tbody > tr:nth-child(6) > td:nth-child(2) > table > tbody > tr:nth-child(1) > td:nth-child(2) > font:nth-child(1) > select")
    select.click()
    term_selected = driver.find_element_by_tag_name("body > table > tbody > tr.ContentBody > td:nth-child(2) > table > tbody > tr:nth-child(6) > td:nth-child(2) > table > tbody > tr:nth-child(1) > td:nth-child(2) > font:nth-child(1) > select > option:nth-child("+str(term)+")")
    term_selected.click()
    select2 = driver.find_element_by_tag_name("body > table > tbody > tr.ContentBody > td:nth-child(2) > table > tbody > tr:nth-child(6) > td:nth-child(2) > table > tbody > tr:nth-child(1) > td:nth-child(2) > font:nth-child(2) > select")
    select2.click()
    year_selected = driver.find_element_by_tag_name("body > table > tbody > tr.ContentBody > td:nth-child(2) > table > tbody > tr:nth-child(6) > td:nth-child(2) > table > tbody > tr:nth-child(1) > td:nth-child(2) > font:nth-child(2) > select > option:nth-child(5)")
    year_selected.click()
    search = driver.find_element_by_tag_name("body > table > tbody > tr.ContentBody > td:nth-child(2) > table > tbody > tr:nth-child(7) > td:nth-child(2) > table > tbody > tr > td > font:nth-child(4) > input[type=""submit""]")
    search.click()

def scrap_Data():
    df = pd.DataFrame()
    row = 4
    page = 1
    try:
        while True:
            rows = len(driver.find_elements(By.XPATH, "/html/body/table/tbody/tr[1]/td[2]/font/font/font/table/tbody/tr"))
            faculty = driver.find_element_by_xpath("/html/body/table/tbody/tr[1]/td[2]/div[1]/font/b")
            year = driver.find_element_by_xpath("/html/body/table/tbody/tr[1]/td[2]/div[3]/font/font/font/b[1]")
            if row == rows and page == 1:
                element = driver.find_element_by_xpath("/html/body/table/tbody/tr[1]/td[2]/font/font/font/table/tbody/tr["+str(row)+"]/td[2]/font/a")
                element.click()
                row = 4
                page += 1
            elif row == rows and page >= 2:
                element = driver.find_element_by_xpath("/html/body/table/tbody/tr[1]/td[2]/font/font/font/table/tbody/tr["+str(row)+"]/td[2]/font/a[2]")
                element.click()
                row = 4
                page += 1
            else:
                while row != rows:
                    for i in driver.find_elements(By.XPATH, "/html/body/table/tbody/tr[1]/td[2]/font/font/font/table/tbody/tr["+str(row)+"]"):
                        campus = i.find_element_by_xpath("td[2]").text
                        course = i.find_element_by_xpath("td[3]").text
                        course_code = i.find_element_by_xpath("td[5]").text
                        subject = i.find_element_by_xpath("td[6]").text
                        credit = i.find_element_by_xpath("td[7]").text
                        section = i.find_element_by_xpath("td[8]").text
                        time_study = i.find_element_by_xpath("td[9]").text
                        #time_exam = i.find_element_by_xpath("td[10]").text
                        receive = i.find_element_by_xpath("td[11]").text
                        remain = i.find_element_by_xpath("td[12]").text
                        status = i.find_element_by_xpath("td[13]").text

                        ### แบ่งชื่อวิชากับชื่ออาจารย์ออกจากกัน
                        subject = subject.splitlines()
                        professor = ""
                        if len(subject) > 1:
                            check = subject[1].startswith("*")
                            if check:
                                subject.remove(subject[1])
                            for j in range(len(subject)):
                                if j == len(subject)-1:
                                    professor += subject[j]
                                    break
                                else:
                                    professor += subject[j+1]+", "

                        ### แบ่ง วัน เวลา และอาคาร
                        time_study = time_study.split(" ")
                        day = ""
                        time = ""
                        room = ""
                        if 1 < len(time_study) < 4:
                            day = time_study[0]
                            time = time_study[1]
                            room = time_study[2]

                        elif len(time_study) > 4:
                            temp = time_study[2].splitlines()
                            day = time_study[0]+", "+temp[1]
                            time = time_study[1]+", "+time_study[3]
                            room = temp[0]+", "+time_study[4]

                        df = df.append({
                            'ศูนย์': campus,
                            'หลักสูตร': course,
                            'รหัสวิชา': course_code,
                            'ชื่อวิชา': subject[0],
                            'คณะ': faculty.text,
                            'หน่วยกิต': credit,
                            'Section': section,
                            'วัน': day,
                            'เวลา': time,
                            'อาคาร': room,
                            'ชื่อผู้สอน': professor,
                            'จำนวนรับ': receive,
                            'เหลือ': remain,
                            'สถานะ': status,
                            'ภาคการศึกษา': year.text
                        }, ignore_index=True)
                    row += 1
    except:
        element = driver.find_element_by_xpath("/html/body/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td/div/ul/li/a")
        element.click()
        print("Ok")
        return df
        pass



select_Faculty(faculty=3)
select_year(term=1)
df = scrap_Data()

select_Faculty(faculty=3)
select_year(year=2)
df = df.append(scrap_Data())

select_Faculty(faculty=6)
select_year(year=1)
df = df.append(scrap_Data())

select_Faculty(faculty=6)
select_year(year=2)
df = df.append(scrap_Data())

df.to_excel("Fact.xlsx", sheet_name='Fact', encoding='utf-8-sig', index=False)
