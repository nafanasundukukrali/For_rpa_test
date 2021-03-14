# - * - coding: utf - 8 -
import pandas as pd
import rpa as r
import win32com.client as win32
import os
import sys

def get_table(file_name):
    r.init()
    r.url('https://yandex.ru/')
    kol = 0
    is_exist = True
    while not r.exist(
            '//*[@id="wd-_topnews-1"]/div/div[3]/div/div/div[1]/a'):
        r.wait()
        kol += 5
        if kol == 60:
            is_exist = False
            break
    if is_exist:
        r.dom('document.querySelectorAll("a").forEach(function(link){link.setAttribute("target", "_self");});')
        r.click('//*[@id="wd-_topnews-1"]/div/div[3]/div/div/div[1]/a')
        xpath_of_table = "/html/body/div[6]/div/div[2]/div[1]/div[1]/div[1]/div[2]/div/div[2]/div/div[2]"  # type: str
        date = r.read(xpath_of_table + '/div[1]/div[1]')
        course = r.read(xpath_of_table + '/div[1]/div[3]')
        changing = r.read(xpath_of_table + '/div[1]/div[2]')
        usd = pd.DataFrame(
            {date: [r.read(xpath_of_table + '/div[' + str(i) + ']/div[1]') for i
                    in range(2, 12)],
             course: [float(r.read(xpath_of_table + '/div[' + str(i) + ']/div[3]').replace(',', '.')) for i
                      in range(2, 12)],
             changing: [float(r.read(xpath_of_table + '/div[' + str(i) + ']/div[2]').replace(',', '.')) for i in
                      range(2, 12)]})
        r.dom('window.history.back()')
        while not r.exist('//*[@id="wd-_topnews-1"]/div/div[3]/div/div/div[2]/a'):
            r.wait()
        kol = 0
        while not r.exist(
                '//*[@id="wd-_topnews-1"]/div/div[3]/div/div/div[2]/a'):  # TODO: Учитывай наличие вообще соединения, наличие блока, загрузка страницы
            r.wait()
            kol += 5
            if kol == 60:
                is_exist = False
                break
        r.dom('document.querySelectorAll("a").forEach(function(link){link.setAttribute("target", "_self");});')
        r.click('//*[@id="wd-_topnews-1"]/div/div[3]/div/div/div[2]/a')
        euro = pd.DataFrame(
            {date: [r.read(xpath_of_table + '/div[' + str(i) + ']/div[1]') for i
                            in range(2, 12)],
             course: [float(r.read(xpath_of_table + '/div[' + str(i) + ']/div[3]').replace(',', '.')) for i
                      in range(2, 12)],
             changing: [float(r.read(xpath_of_table + '/div[' + str(i) + ']/div[2]').replace(',', '.')) for i in
                              range(2, 12)]})
        usd_euro = pd.DataFrame({'Отношение курса': [euro.iloc[i, 1] / usd.iloc[i, 1] for i in range(0, 10)]})
        df = pd.concat([usd, euro, usd_euro], axis=1)
        pd.set_option('max_colwidth', 120)
        reload(sys)
        sys.setdefaultencoding('utf-8')
        df.to_excel(file_name, index=False)
    r.close()


def table_refactoring(file_path):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(file_path)
    excel.Worksheets(1).Activate()
    ws = wb.Worksheets(1)
    for i in range(2, 12):
        ws.Cells(i, 3).NumberFormat = "_-[$$-en-US]* # ##0,00_ "
    for i in range(2, 12):
        ws.Cells(i, 6).NumberFormat = '_-[$' + chr(136) + '-x-euro2]* # ##0,00_-'
    for i in range(2, 12):
        ws.Cells(i, 7).NumberFormat = '0,00'
    for i in range(2, 12):
        ws.Cells(i, 2).NumberFormat = '0,00'
    for i in range(2, 12):
        ws.Cells(i, 5).NumberFormat = '0,00'
    ws.Cells(1, 12).Value = u'Сумма столбца "Курс доллара"'
    ws.Cells(2, 12).FormulaLocal = u'=СУММ(C2:C11)'
    excel.ActiveSheet.Columns.AutoFit()
    number_str = excel.ActiveSheet.UsedRange.Rows.Count
    wb.Save()
    wb.Close()
    return number_str

def send_main(file_path, number_str):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mapi = outlook.GetNameSpace("MAPI")
    sent_mail = mapi.GetDefaultFolder(5)
    messages = list(sent_mail.Items)
    message = messages[0]
    if message.SenderEmailType == "SMTP":
        mail.To = message.SenderEmailAddress
    elif message.SenderEmailType == "EX":
        mail.To = message.Sender.GetExchangeUser().PrimarySmtpAddress
    mail.Subject = mail.To
    mail.Body = u'В таблице {number_str} строк'.format(number_str=number_str)
    attachment = file_path
    mail.Attachments.Add(attachment)
    mail.Send()


if __name__ == "__main__":
    filename = r'today_moex_data.xlsx'
    file_path = os.path.abspath(filename)
    get_table(filename)
    number_str = table_refactoring(file_path)
    send_main(file_path, number_str)
