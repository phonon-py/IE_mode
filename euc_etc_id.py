from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options

# Edgeのオプション設定
edge_options = Options()
edge_options.use_chromium =True
edge_options.add_experimental_option('ieMode',True)
edge_options.binary_location = 'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'

# Edge WebDriverのパスを指定する
service = Service(executable_path="\\cpiweb\\www\\4930\\ソフトウェア\\MicrosoftEdge\\WebDriver\\120.0.2210.121\\msedgedriver.exe")

# WebDriverのインスタンスを作成する
driver = webdriver.Edge(service=service, options=edge_options)

# Googleのホームページを開く
driver.get("https://www.google.com")

# ウェブページのタイトルを出力する（確認用）
print(driver.title)

# ブラウザを閉じる
driver.quit()
