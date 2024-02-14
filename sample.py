from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options

# WebDriverのパスを指定（rを使用してraw stringとして扱う）
webdriver_path = r'C:\path\to\your\edgedriver.exe'

# Edgeのオプションを設定
options = Options()
options.use_chromium = True  # ChromiumベースのEdgeを使用
options.add_argument("ie.mode=IE11")  # IEモードで起動

# WebDriverサービスを設定
service = Service(executable_path=webdriver_path)

# WebDriverを初期化（EdgeをIEモードで起動）
driver = webdriver.Edge(service=service, options=options)

# Googleのホームページを開く
driver.get('https://www.google.com')

# 作業が完了したら、ブラウザを閉じる
# driver.quit()
