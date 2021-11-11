import xlrd
import xlwt
from selenium import webdriver
from time import sleep


def chrome(name_list):
    un_find_list = []
    browser = webdriver.Chrome(
        # 驱动
        executable_path="/home/shello/Code/python/src/chromedriver"
    )
    browser.maximize_window()
    browser.get("https://www.gdzyz.cn/manager/users/index.do")
    a = input("请打开录用页面之后再继续...(y/n)")
    if a=="n":
        return
    iframe = browser.find_element_by_id("mission_attend4836614")
    browser.implicitly_wait(5)
    browser.switch_to.frame(iframe)

    for name in name_list:
        userName = browser.find_element_by_id("userName")
        browser.implicitly_wait(15)
        userName.click()
        userName.clear()
        userName.send_keys(name)
        browser.find_element_by_id("f_search_btn").click()
        sleep(2)
        tbody = browser.find_elements_by_css_selector("#maingridgrid > div.l-grid1 > div.l-grid-body.l-grid-body1 > div.l-grid-body-inner > table > tbody")
        if len(tbody)>0:
            checkBox = browser.find_element_by_css_selector("#maingrid\|hcell\|c101 > div > div")
            checkBox.click()
            browser.find_element_by_link_text("补录").click()
            # yes = browser.find_element_by_css_selector("body > div.l-dialog > table > tbody > tr:nth-child(2) > td.l-dialog-cc > div > div.l-dialog-buttons > div > div.l-dialog-btn.l-dialog-btn-ok > div.l-dialog-btn-inner")
            no = browser.find_element_by_css_selector("body > div.l-dialog > table > tbody > tr:nth-child(2) > td.l-dialog-cc > div > div.l-dialog-buttons > div > div.l-dialog-btn.l-dialog-btn-no > div.l-dialog-btn-inner")
            no.click()
            browser.implicitly_wait(5)
            print("已录用："+name)
        else:
            un_find_list.append(name)
            print("找不到："+name)
    # a = input("运行结束，按任意键退出...")
    return un_find_list



def get_name_list(excelPath):
    book = xlrd.open_workbook(excelPath)
    sheet = book.sheet_by_index(0)
    result = []
    for i in range(0, sheet.nrows):
        result.append(sheet.row_values(i)[0])
        print(sheet.row_values(i)[0])
    return result

def write_un_find(un_find_list, excel_path):
    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet('未找到')
    for i in range(len(un_find_list)):
        worksheet.write(i,0, label=un_find_list[i])
    workbook.save(excel_path)

if __name__ == "__main__":
    name_list = get_name_list("/home/shello/Code/python/src/test.xls")
    un_find_list = chrome(name_list)
    write_un_find(un_find_list, "/home/shello/Code/python/src/test_out.xls")
