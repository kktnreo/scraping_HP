from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException
import time
import requests
from bs4 import BeautifulSoup
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv
from datetime import datetime
from urllib.parse import urlparse
import re
import os

# .envファイルを読み込む
load_dotenv()

def chrome_driver():
	# Chrome Driverの設定
	chrome_options = Options()
	# chrome_options.add_argument("--headless")  # ヘッドレスモードで実行

	# Webドライバーの起動
	driver = webdriver.Chrome(options=chrome_options)
	driver.maximize_window()
	url = "https://www.business-expo.jp/exhibitor/"
	# 指定サイトにアクセス
	driver.get(url)
	# ページが完全に読み込まれるまで少し待機
	time.sleep(5)

	# 企業のURLを取得 (XPathは適宜修正)
	company_elements = driver.find_elements(By.XPATH, '//*[@id="luxy"]/main/section[2]/div[2]/div[2]/dl[3]/dd/ol/li/a')  # 企業一覧

	# 各企業のURLをリストに保存
	a_tags = [element.get_attribute('href') for element in company_elements]
	

	company_list = []
	url_list = []
	seo_list = []
	for a_tag in a_tags:
		driver.get(a_tag)  # URL文字列を取得してアクセス
		try:
			company_name = driver.find_element(By.XPATH, '//*[@id="luxy"]/main/section/div/div[2]/div[1]/div/h1').text
		except NoSuchElementException:
			company_name = "要素なし"

		try:
			post_code = driver.find_element(By.XPATH, '//*[@id="luxy"]/main/section/div/div[2]/div[1]/div/dl[dt[text()="郵便番号"]]/dd').text
		except NoSuchElementException:
			post_code = "要素なし"

		try:
			address = driver.find_element(By.XPATH, '//*[@id="luxy"]/main/section/div/div[2]/div[1]/div/dl[dt[text()="住所"]]/dd').text
		except NoSuchElementException:
			address = "要素なし"

		try:
			tel = driver.find_element(By.XPATH, '//*[@id="luxy"]/main/section/div/div[2]/div[1]/div/dl[dt[text()="TEL"]]/dd').text
		except NoSuchElementException:
			tel = "要素なし"

		try:
			post = driver.find_element(By.XPATH, '//*[@id="luxy"]/main/section/div/div[2]/div[1]/div/dl[dt[text()="代表者役職名"]]/dd').text
		except NoSuchElementException:
			post = "要素なし"

		try:
			name = driver.find_element(By.XPATH, '//*[@id="luxy"]/main/section/div/div[2]/div[1]/div/dl[dt[text()="代表者氏名"]]/dd').text
		except NoSuchElementException:
			name = "要素なし"

		try:
			department = driver.find_element(By.XPATH, '//*[@id="luxy"]/main/section/div/div[2]/div[1]/div/dl[dt[text()="担当者部課名"]]/dd').text
		except NoSuchElementException:
			department = "要素なし"

		try:
			home_page = driver.find_element(By.XPATH, '//*[@id="luxy"]/main/section/div/div[2]/div[1]/div/dl[dt[text()="ホームページ"]]/dd/a').get_attribute('href')  # URLを取得
			data = {company_name: home_page}
			url_list.append(data)
		except NoSuchElementException:
			home_page = '要素なし'
			data = {company_name: '要素なし'}
			url_list.append(data)

		data_dict = {
			'会社名': company_name,
			'郵便番号': post_code,
			'住所': address,
			'電話番号': tel,
			'役職名': post,
			'代表者氏名': name,
			'部署名': department,
			'HP': home_page
		}
		company_list.append(data_dict)
	# ブラウザを閉じる
	driver.quit()

	# 各企業ページでSEO情報を取得
	for url in url_list:
		for name, link in url.items():
			title, h1_tags, h2_tags, description = get_seo_info(link)
			seo_dict = {
				'会社名': name,
				'タイトル': title,
				'H1': h1_tags,
				'H2': h2_tags,
				'説明文': description
			}
			seo_list.append(seo_dict)

	df = pd.DataFrame(company_list)
	df2 = pd.DataFrame(seo_list)
	
	spread_sheet(df, df2)


def get_seo_info(url):
	if url != '要素なし':
		response = requests.get(url)
		soup = BeautifulSoup(response.content, 'html.parser')

		# SEOタイトルの取得
		title = soup.find('title').get_text() if soup.find('title') else 'No Title'

		# H1タグの取得（複数ある場合）
		h1_tags = soup.find_all('h1')
		# h1_list = [re.sub(r'\s+', ' ', h1.get_text().strip()) for h1 in h1_tags] if h1_tags else ['No H1 Tags']
		h1_list = ', '.join([re.sub(r'\s+', ' ', h1.get_text().strip()) for h1 in h1_tags]) if h1_tags else 'No H1 Tags'

		# H2タグの取得（複数ある場合）
		h2_tags = soup.find_all('h2')
		# h2_list = [re.sub(r'\s+', ' ', h2.get_text().strip()) for h2 in h2_tags] if h2_tags else ['No H2 Tags']
		h2_list = ', '.join([re.sub(r'\s+', ' ', h2.get_text().strip()) for h2 in h2_tags]) if h2_tags else 'No H1 Tags'

		# メタディスクリプションの取得
		description = soup.find('meta', attrs={'name': 'description'})
		description_content = description['content'] if description else 'No Description'

		return title, h1_list, h2_list, description_content
	else:
		title = '要素なし'
		h1_list = '要素なし'
		h2_list = '要素なし'
		description_content = '要素なし'
		return title, h1_list, h2_list, description_content


def spread_sheet(df, df2):
    # 2つのAPIを記述しないとリフレッシュトークンを3600秒毎に発行し続けなければならない
	scope = ['https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/drive']
    # 認証情報設定
    # ダウンロードしたjsonファイル名をクレデンシャル変数に設定。Pythonファイルと同じフォルダに置く。
	credentials = Credentials.from_service_account_file("rakuten-api-421705-fee8480637fa.json", scopes=scope)
    # 共有設定したスプレッドシートキーを格納
	SPREADSHEET_KEY = os.getenv('SPREADSHEET_KEY')
	gc = gspread.authorize(credentials)
	sheet_name = '展示会参加企業'
	sheet = gc.open_by_key(SPREADSHEET_KEY).worksheet(sheet_name)
	sheet_name2 = 'SEO'
	sheet2 = gc.open_by_key(SPREADSHEET_KEY).worksheet(sheet_name2)
	# データフレームの内容をリストに変換して書き込み
	sheet.clear()  # シートを一旦クリア（必要に応じて）
	sheet.update([df.columns.values.tolist()] + df.values.tolist())
	# データフレームの内容をリストに変換して書き込み
	sheet2.clear()  # シートを一旦クリア（必要に応じて）
	sheet2.update([df2.columns.values.tolist()] + df2.values.tolist())


chrome_driver()

