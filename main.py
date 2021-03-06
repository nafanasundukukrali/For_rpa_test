# - * - coding: utf - 8 -
import pandas as pd
import rpa as r
import win32com.client as win32
import os
import sys
import logging

def get_table(file_name, logging):
    r.init()
    r.url('https://yandex.ru/')
    kol = 0
    is_exist = True
    while not r.present(
            '//*[@id="wd-_topnews-1"]/div/div[3]/div/div/div[1]/a'):
        r.wait(5)
        kol += 5
        if kol >= 30:
            is_exist = False
            break
    if is_exist == False:
        r.dom('alert("Check your connection! Error with yandex.ru page.")')
        r.wait(3)
        r.close()
        logging.info("Check your connection! Check xpath! Error was before USD.")
        sys.exit()
    r.dom('document.querySelectorAll("a").forEach(function(link){link.setAttribute("target", "_self");});')
    r.click('//*[@id="wd-_topnews-1"]/div/div[3]/div/div/div[1]/a')
    try:
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
    except Exception:
        r.close()
        logging.info("Check xpath! Check connection! Error was in USD.")
        sys.exit()
    r.dom('window.history.back()')
    kol = 0
    while not r.present(
            '//*[@id="wd-_topnews-1"]/div/div[3]/div/div/div[2]/a'):
        r.wait(5)
        kol += 5
        if kol >= 30:
            is_exist = False
            break
    if is_exist == False:
        r.dom('alert("Check your connection! Error with yandex.ru page.")')
        r.wait(3)
        r.close()
        logging.info("Check your connection! Check xpath! Error was after USD and before EUR.")
        sys.exit()
    r.dom('document.querySelectorAll("a").forEach(function(link){link.setAttribute("target", "_self");});')
    r.click('//*[@id="wd-_topnews-1"]/div/div[3]/div/div/div[2]/a')
    try:
        euro = pd.DataFrame(
            {date: [r.read(xpath_of_table + '/div[' + str(i) + ']/div[1]') for i
                            in range(2, 12)],
             course: [float(r.read(xpath_of_table + '/div[' + str(i) + ']/div[3]').replace(',', '.')) for i
                      in range(2, 12)],
             changing: [float(r.read(xpath_of_table + '/div[' + str(i) + ']/div[2]').replace(',', '.')) for i in
                              range(2, 12)]})
    except Exception:
        r.dom('alert("Check your connection! Error with yandex.ru page.")')
        r.wait(3)
        r.close()
        logging.info("Check your connection! Check xpath! Error was after USD and in EUR.")
        sys.exit()
    usd_euro = pd.DataFrame({'?????????????????? ??????????': [euro.iloc[i, 1] / usd.iloc[i, 1] for i in range(0, 10)]})
    df = pd.concat([usd, euro, usd_euro], axis=1)
    reload(sys)
    sys.setdefaultencoding('utf-8')
    try:
        df.to_excel(file_name, index=False)
    except Exception:
        r.close()
        logging.info("Error with filename!")
        sys.exit()
    return r


def table_refactoring(file_path, logging, r):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    try:
        wb = excel.Workbooks.Open(file_path)
    except Exception:
        r.dom('alert("Excel Error! Have you close excel?")')
        r.wait(3)
        r.close()
        logging.info("Error with opening excel file")
        sys.exit()
    try:
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
        ws.Cells(1, 12).Value = u'?????????? ?????????????? "???????? ??????????????"'
        ws.Cells(2, 12).FormulaLocal = u'=????????(C2:C11)'
        excel.ActiveSheet.Columns.AutoFit()
        number_str = excel.ActiveSheet.UsedRange.Rows.Count
    except Exception:
        r.dom('alert("Excel: error with data.")')
        r.wait(3)
        r.close()
        logging.info("Excel: error with data.")
        wb.Close()
        sys.exit()
    wb.Save()
    wb.Close()
    return number_str

def send_main(file_path, number_str, logging, r):
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
    mail.Body = u'?? ?????????????? {number_str} ??????????'.format(number_str=number_str)
    attachment = file_path
    mail.Attachments.Add(attachment)
    try:
        mail.Send()
    except:
        r.dom('alert("Outlook error.")')
        r.wait(3)
        r.close()
        logging.info("Outlook")
        sys.exit()
    os.remove(file_path)


if __name__ == "__main__":
    if os._exists("main.log"):
        os.remove("main.log")
    logging.basicConfig(filename="main.log", level=logging.INFO)
    filename = r'today_moex_data.xlsx'
    r = get_table(filename, logging)
    file_path = os.path.abspath(filename)
    number_str = table_refactoring(file_path, logging, r)
    send_main(file_path, number_str, logging, r)
    r.close()
