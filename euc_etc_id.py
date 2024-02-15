from time import sleep
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from config import BINARY_LOCATION,MSEDGE_DRIVER, USER_NAME, USER_PASSWORD

# Edgeのオプション設定
edge_options = Options()
edge_options.use_chromium = True  # ChromiumベースのEdgeを使用することを指定
edge_options.binary_location = BINARY_LOCATION

# Edge WebDriverのパスを指定する
service = Service(executable_path=MSEDGE_DRIVER)

# WebDriverのインスタンスを作成する
driver = webdriver.Edge(service=service, options=edge_options)
username = USER_NAME
password = USER_PASSWORD

def login_kikan_with_seleniun(driver,username,password):
    """基幹システムにログインする関数

    Args:
        driver: Selenium WebDriverのインスタンス
        username: ユーザー名
        password: パスワード

    Returns:
        ログイン後のWebDriverのインスタンス
    """    
    # 新基幹のページを開く
    driver.get("http://core-apserver.prec.canon.co.jp/ZZ_CPIPortal/page/myLogin.jsp")

    # ユーザー入力欄の要素を取得
    user_element = driver.find_element(By.ID, 'user')
    # ユーザー入力欄に文字列を入力
    user_element.send_keys(username)

    # パスワード入力欄の要素を取得
    user_element = driver.find_element(By.ID, 'pass')
    # パスワード入力欄に文字列を入力
    user_element.send_keys(password)

    # ボタンの要素を取得
    submit_element = driver.find_element(By.ID, "submit")
    # ボタンをクリック
    submit_element.click()

    return driver
login_kikan_with_seleniun(driver,username,password)

#!
# 別ウィンドウを開く前に現在ウィンドウのハンドルを取得
current_window_handle = driver.current_window_handle
print(current_window_handle)
link_element = driver.find_element(By.ID, 'sidA_LNK_05')
link_element.click()
# 別ウィンドウが開くまで待つ
WebDriverWait(driver, 10).until(
    EC.number_of_windows_to_be(2)
)

# 別ウィンドウのハンドルを取得
new_window_handle = [
    handle for handle in driver.window_handles if handle != current_window_handle
][0]
print(new_window_handle)
# 別ウィンドウに切り替え
driver.switch_to.window(new_window_handle)

# 別ウィンドウで要素をクリック
element_element = driver.find_element(By.ID, 'a_lnk_01_item')
element_element.click()
sleep(10)


# ブラウザを閉じる
driver.quit()
