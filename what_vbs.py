import datetime
import requests
from selenium import webdriver
import cx_Oracle
import smtplib
from email.mime.text import MIMEText

def main():
    try:
        ie_mainte()
        quit_script(0)
    except Exception as e:
        log_error(e)
        send_msg("cpi-it-infra@mail.canon", "", "【エラー】インターネット自動メンテ", "実行フォルダの ie_mainte.log を確認してください")
        quit_script(e.errno)

def ie_mainte():
    print(f"{datetime.datetime.now()}: 処理開始")
    # ウェブページへのアクセスと操作にはSeleniumを使用
    driver = webdriver.Chrome()  # ChromeのWebDriverオブジェクトを作成
    driver.get("http://cipapp.cgn.canon.co.jp/cip/app/WXAH/WXAH_05/WXAHA300.asp")
    wait_ie(driver)
    
    # ログイン操作など、具体的なウェブページの操作
    # driver.find_element_by_id("inuserid").send_keys("ユーザーID")
    # 以下同様に操作を続ける

    # データベース操作
    # connection = cx_Oracle.connect('username', 'password', 'dsn')
    # cursor = connection.cursor()
    # cursor.execute("SELECT * FROM some_table")
    # for row in cursor:
    #     print(row)

    # 終了処理
    driver.quit()
    print(f"{datetime.datetime.now()}: 処理終了")

def wait_ie(driver):
    # ページが完全にロードされるまで待機
    while driver.execute_script("return document.readyState;") != "complete":
        time.sleep(0.5)

def log_error(e):
    print(f"{datetime.datetime.now()}:{e.errno}:{e.strerror}")

def send_msg(to, cc, subject, text):
    msg = MIMEText(text)
    msg["Subject"] = subject
    msg["From"] = "ACH-CSYDB@prec.canon.co.jp"
    msg["To"] = to
    if cc:
        msg["Cc"] = cc

    # SMTPサーバーを通じてメールを送信
    with smtplib.SMTP("nonauth-smtp.global.canon.co.jp", 25) as smtp:
        smtp.send_message(msg)

def quit_script(code):
    exit(code)

if __name__ == "__main__":
    main()
