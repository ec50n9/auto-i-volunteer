import xlrd
import xlwt
from selenium import webdriver


def chrome(name_list):
    un_find_list = name_list
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

    num = int(input("条目数："))
    for i in range(1, num+1):
        selector = "#maingrid\|2\|r1{0:03d}\|c105 > div".format(i)
        print(selector)
        name_cell = browser.find_element_by_css_selector(selector)
        name = name_cell.get_attribute("innerText")
        if name in name_list:
            print("找到："+name)
            name_cell.click()
            un_find_list.remove(name)
        else:
            print("未找到："+name)
    a = input("运行结束，按任意键退出...")
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
    write_un_find(un_find_list, "/home/shello/Code/python/src/test_四小时.xls")

