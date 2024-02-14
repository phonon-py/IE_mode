from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options

# Edgeのオプション設定
edge_options = Options()
edge_options.use_chromium = True  # ChromiumベースのEdgeを使用することを指定

# 注意: IEモードを有効にする特定のSeleniumのオプションは存在しないため、
# IEモードを利用する場合はEdgeの設定やグループポリシーを介して行う必要があります。
# ここではその設定は省略しています。

# Edgeのバイナリ位置を指定
edge_options.binary_location = r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'

# Edge WebDriverのパスを指定する
# 注意: ネットワークパスではなく、ローカルファイルシステム上のパスを使用しています。
service = Service(executable_path=r'C:\path\to\your\msedgedriver.exe')

# WebDriverのインスタンスを作成する
driver = webdriver.Edge(service=service, options=edge_options)

# Googleのホームページを開く
driver.get("https://www.google.com")

# ウェブページのタイトルを出力する（確認用）
print(driver.title)

# ブラウザを閉じる
driver.quit()
