from time import sleep
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from config import BINARY_LOCATION,MSEDGE_DRIVER, USER_NAME, USER_PASSWORD,KIKAN_URL, DF_PATH

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
    driver.get(KIKAN_URL)

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
# 別ウィンドウに切り替え
driver.switch_to.window(new_window_handle)

# 棚卸し開始をクリック
inventory_start_element = driver.find_element(By.ID, 'a_lnk_01_item')
inventory_start_element.click()

# 詳細をクリック
detail_element = driver.find_element(By.ID, 'sidAPP_BUTTON')
detail_element.click()

# 要素を取得
element = driver.find_element(By.ID, "g[0].tana_button_item")
# 要素をクリック
element.click()

# プルダウン要素を取得
#! NAMEの数は続いている
max_loop = 1000
try:
    for i in range(max_loop):
        select_element = driver.find_element(By.NAME, f"g[{i}]._cd_youhi_str")
        select = Select(select_element)
        select.select_by_value('1')
        # 15回繰り返したら次へボタンを押す
        if i % 15 == 14:
            # 更新ボタンを押す
            element = driver.find_element(By.ID, "sidUPDATE_A")
            # 要素をクリック
            element.click()
            sleep(2)
            # 次へボタンを取得
            element = driver.find_element(By.ID, "_eventcursornextpage_g")
            # 要素をクリック
            element.click()
            sleep(5)
except:
    print('end')
# ブラウザを閉じる
driver.quit()
