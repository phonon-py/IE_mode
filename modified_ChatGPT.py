from time import sleep
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from config import BINARY_LOCATION, MSEDGE_DRIVER, USER_NAME, USER_PASSWORD, KIKAN_URL

# Edgeのオプション設定
def configure_edge_options():
    """Edgeブラウザのオプションを設定する"""
    edge_options = Options()
    edge_options.use_chromium = True  
    edge_options.binary_location = BINARY_LOCATION
    return edge_options

# WebDriverのインスタンスを作成する関数
def create_webdriver():
    """WebDriverのインスタンスを作成し、返す"""
    service = Service(executable_path=MSEDGE_DRIVER)
    edge_options = configure_edge_options()
    driver = webdriver.Edge(service=service, options=edge_options)
    return driver

# 基幹システムにログイン
def login_kikan(driver, username=USER_NAME, password=USER_PASSWORD):
    """基幹システムにログインする"""
    driver.get(KIKAN_URL)
    driver.find_element(By.ID, 'user').send_keys(username)
    driver.find_element(By.ID, 'pass').send_keys(password)
    driver.find_element(By.ID, "submit").click()

# 新しいウィンドウでの作業を開始
def switch_to_new_window(driver):
    """新しいウィンドウへの切り替えを管理する"""
    current_window = driver.current_window_handle
    WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
    new_window = [h for h in driver.window_handles if h != current_window][0]
    driver.switch_to.window(new_window)

def perform_specific_operations(driver):
    """
    特定のページ操作を実行する。

    この関数は、棚卸し開始から詳細ページに移動し、特定のアクションを行うまでの一連のステップを実装しています。
    ループ内でプルダウン選択を行い、指定された条件で「更新ボタン」と「次へボタン」をクリックします。
    エラーが発生した場合は、その時点で処理を中断し、どのステップで中断されたかを表示します。
    """

    # EXCELの読み込み
    df = pd.read_excel('DataFarame_path')
    
    # 棚卸し開始をクリック
    inventory_start_element = driver.find_element(By.ID, 'a_lnk_01_item')
    inventory_start_element.click()

    # 詳細をクリック
    detail_element = driver.find_element(By.ID, 'sidAPP_BUTTON')
    detail_element.click()

    #! 動作確認したら次のテーブルへ遷移する処理を追加する
    # 要素を取得してクリック
    element = driver.find_element(By.ID, "g[0].tana_button_item")
    element.click()

    # プルダウン要素を操作
    max_loop = len(df)
    for i in range(max_loop):
        try:
            # プルダウン要素を取得
            select_element = driver.find_element(By.NAME, f"g[{i}]._cd_youhi_str")
            select = Select(select_element)
            # dfの要否カラムの値に基づいてプルダウン項目を選択
            value_to_select = str(df.iloc[i]['要否'])  # dfのi行目の要否カラムの値を取得
            select.select_by_value(value_to_select)
            
            # 15回繰り返したら次へボタンを押す
            if i % 15 == 14:
                # 更新ボタンをクリック
                driver.find_element(By.ID, "sidUPDATE_A").click()
                # 更新ボタンのクリック後、次へボタンがクリック可能になるまで待機
                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "_eventcursornextpage_g"))
                )

                # 次へボタンをクリック
                driver.find_element(By.ID, "_eventcursornextpage_g").click()
                # 次へボタンのクリック後、何らかの明確なページ遷移や要素の表示を待機する処理を追加
                # 例: 次ページの特定の要素が表示されるまで待機
                WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.ID, f"g[{i}]._cd_youhi_str"))
                )

        except Exception as e:
            print(f"Process ended or failed at iteration {i}: {e}")
            break  # 要素が見つからない場合など、処理を終了

# メインの処理
def main():
    driver = create_webdriver()
    try:
        login_kikan(driver)
        # その他のページ操作
        switch_to_new_window(driver)
        perform_specific_operations(driver)
    except Exception as e:
        print(f"Error occurred: {e}")
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
