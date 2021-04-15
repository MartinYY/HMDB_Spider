from google_trans_new import google_translator
from selenium import webdriver
from lxml import etree
import time
import pandas as pd
import xlrd
import random
import os

file_path = os.path.join(os.path.expanduser('~'), "Desktop") + "\\test.xlsx"
driver_path = r"C:\Program Files (x86)\Google\Chrome\Application\chromedriver.exe"
url = "https://hmdb.ca/"
datas = {}
writer = pd.ExcelWriter(file_path, engine='xlsxwriter')

def update_data(datas, index, name):
    # 读取一个excel的文本文件（当前默认时读一个文件的一个sheet页）
    ex = pd.read_excel(file_path, sheet_name=index)
    # 用pd格式化
    df = pd.DataFrame(ex)

    columns = ex.columns.tolist()
    new_columns = ["Common Name", "Description", "Structure", "Chemical Formula",
                   "Average Molecular Weight", "Monoisotopic Molecular Weight", "IUPAC Name",
                   "Traditional Name", "CAS Registry Number", "SMILES", "InChI Identifier", "InChI Key", "Disposition"]

    for column in new_columns:
        if column not in columns:
            df.insert(df.shape[1], column, "")

    for key in datas.keys():
        data = datas[key]
        keys = data.keys()
        for k in keys:
            # 执行修改操作
            df.update(pd.Series(data[k], name=k, index=[key]))

    # 执行数据更新后的保存操作：这里有个问题就是源文件覆盖保存，会没有特定的样式，需要再升级一下
    df.to_excel(writer, sheet_name=name, index=False)


def pasre_page(driver, data):
    html = etree.HTML(driver.page_source)
    trs = html.xpath('//tr')
    for tr in trs:
        th = tr.xpath('./th/text()')

        if len(th) == 0:
            continue
        else:
            th = th[0]

        if th == "Common Name":
            td = tr.xpath('./td/strong/text()')[0]
        elif th == "Description":
            translator = google_translator(timeout=10)
            td = translator.translate(tr.xpath('./td')[0].xpath('string(.)'), 'zh-cn')
        elif th == "Structure":
            td = url + tr.xpath('./td//img/@src')[0]
        elif th == "Chemical Formula":
            td = tr.xpath('./td')
            td = td[0].xpath('string(.)')
        elif th == "SMILES":
            td = tr.xpath('./td/div/text()')[0]
        elif th == "InChI Identifier":
            td = tr.xpath('./td/div/text()')[0]
        elif th == "Disposition":
            td = tr.xpath('./td')[0].xpath('string(.)')
        elif th in data.keys():
            td = tr.xpath('./td/text()')
            if len(td) == 0:
                continue
            else:
                td = td[0]
        else:
            continue

        if len(data[th]) == 0:
            data[th] = td

    try:
        disposition = driver.find_element_by_xpath('//a[contains(text(),"Disposition")]/../../td').text
        data["Disposition"] = disposition
    except:
        return

def get_data(index):
    ex = pd.read_excel(file_path, sheet_name=index)
    df = pd.DataFrame(ex)
    results = {}
    for row in df.itertuples(name="RowData"):
        if pd.isnull(row.Name):
            name = ""
        else:
            name = row.Name
        if pd.isnull(row.HMDB):
            hmdb = ""
        else:
            hmdb = row.HMDB
        data = {"name": name, "hmdb": hmdb}
        results[row.Index] = data
    return results


def getsheet_data(driver):
    wb = xlrd.open_workbook(file_path)
    sheets = wb.sheet_names()
    for i in [0, 1]:
        results = get_data(i)
        print('查询sheet:{}, 列项长度:{}', sheets[i], len(results))
        print('开始查询-------------------------')
        search_data(results, driver)
        print('查询结束-------------------------')
        print('开始更新excel数据-------------------------')
        update_data(datas, i, sheets[i])
        print('更新excel数据结束-------------------------')

    writer.save()
    driver.quit()


def search_data(results, driver):
    for key, value in results.items():
        try:
            if len(value.get("hmdb")) != 0:
                search_win = driver.find_element_by_id('query')
                search_win.send_keys(value.get("hmdb"))
                search_btn = driver.find_element_by_class_name('btn-search')
                search_btn.click()
                time.sleep(random.uniform(2, 4))
            elif len(value.get("name")) != 0:
                search_win = driver.find_element_by_id('query')
                search_win.send_keys(value.get("name"))
                search_btn = driver.find_element_by_class_name('btn-search')
                search_btn.click()
                time.sleep(random.uniform(2, 4))
                content_page = driver.find_element_by_class_name('btn-card')
                content_page.click()
                time.sleep(random.uniform(2, 4))
            else:
                continue
        except Exception as e:
            continue

        data = {
            "Common Name": "",
            "Description": "",
            "Structure": "",
            "Chemical Formula": "",
            "Average Molecular Weight": "",
            "Monoisotopic Molecular Weight": "",
            "IUPAC Name": "",
            "Traditional Name": "",
            "CAS Registry Number": "",
            "SMILES": "",
            "InChI Identifier": "",
            "InChI Key": "",
            "Disposition": ""
        }
        pasre_page(driver, data)
        datas[key] = data


def s2h(seconds):
    '''
    将秒数转为小时数
    '''
    m, s = divmod(seconds, 60)
    h, m = divmod(m, 60)
    return ("%02d小时%02d分钟%02d秒" % (h, m, s))


def main():
    driver = webdriver.Chrome(executable_path=driver_path)
    driver.get(url)

    time.perf_counter()
    getsheet_data(driver)
    print('－－－－－－－－－－－－－－－－－－－－－－－－－－')
    print('爬取完毕，共运行：' + s2h(time.perf_counter()))


if __name__ == '__main__':
    main()
