"""
スクレイピング処理モジュール
認証済みのWebDriverを使用してKPIデータを抽出する
"""

import pandas as pd
import time
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import logging

# ログ設定
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class KpiScraper:
    """KPIデータのスクレイピングを行うクラス"""
    
    def __init__(self, driver, config):
        """
        初期化
        
        Args:
            driver: selenium WebDriverインスタンス
            config: configparserオブジェクト
        """
        self.driver = driver
        self.config = config
        self.timeout = config.getint('Browser', 'timeout', fallback=30)
    
    def scrape_table_data(self, url):
        """
        指定URLからテーブルデータを抽出する
        
        Args:
            url (str): スクレイピング対象のURL
            
        Returns:
            pd.DataFrame: 抽出されたテーブルデータ
        """
        try:
            logger.info(f"スクレイピング開始: {url}")
            
            # ページにアクセス
            self.driver.get(url)
            
            # ページの読み込み完了を待機
            WebDriverWait(self.driver, self.timeout).until(
                EC.presence_of_element_located((By.TAG_NAME, "table"))
            )
            
            # 少し待機してJavaScriptの実行を待つ
            time.sleep(3)
            
            # ページソースを取得
            page_source = self.driver.page_source
            
            # BeautifulSoupでHTMLを解析
            soup = BeautifulSoup(page_source, 'html.parser')
            
            # テーブルを検索
            tables = soup.find_all('table')
            
            if not tables:
                logger.warning("テーブルが見つかりませんでした")
                return pd.DataFrame()
            
            # 最初のテーブルを使用（複数ある場合は最大のテーブルを選択）
            target_table = max(tables, key=lambda t: len(t.find_all('tr')))
            
            # pandasでテーブルを読み込み
            df = self._parse_table_to_dataframe(target_table)
            
            logger.info(f"データ抽出完了: {len(df)}行のデータを取得")
            return df
            
        except Exception as e:
            logger.error(f"スクレイピングエラー: {e}")
            return pd.DataFrame()
    
    def _parse_table_to_dataframe(self, table):
        """
        BeautifulSoupのテーブル要素をDataFrameに変換する
        
        Args:
            table: BeautifulSoupのテーブル要素
            
        Returns:
            pd.DataFrame: 変換されたDataFrame
        """
        try:
            # テーブルの行を取得
            rows = table.find_all('tr')
            
            if not rows:
                return pd.DataFrame()
            
            # ヘッダー行を取得
            header_row = rows[0]
            headers = []
            
            for th in header_row.find_all(['th', 'td']):
                header_text = th.get_text(strip=True)
                headers.append(header_text if header_text else f"列{len(headers)+1}")
            
            # データ行を取得
            data_rows = []
            for row in rows[1:]:
                cells = row.find_all(['td', 'th'])
                row_data = []
                
                for cell in cells:
                    cell_text = cell.get_text(strip=True)
                    row_data.append(cell_text)
                
                # 列数を合わせる
                while len(row_data) < len(headers):
                    row_data.append('')
                
                if row_data:  # 空行でない場合のみ追加
                    data_rows.append(row_data[:len(headers)])
            
            # DataFrameを作成
            df = pd.DataFrame(data_rows, columns=headers)
            
            # 空の行を削除
            df = df.dropna(how='all')
            
            return df
            
        except Exception as e:
            logger.error(f"テーブル解析エラー: {e}")
            return pd.DataFrame()
    
    def scrape_kpi_data(self):
        """
        設定ファイルで指定されたURLからKPIデータを抽出する
        
        Returns:
            pd.DataFrame: 抽出されたKPIデータ
        """
        try:
            target_url = self.config.get('Scraper', 'target_url')
            return self.scrape_table_data(target_url)
            
        except Exception as e:
            logger.error(f"KPIデータ抽出エラー: {e}")
            return pd.DataFrame()
    
    def scrape_multiple_tables(self, url):
        """
        複数のテーブルがある場合に全てのテーブルを抽出する
        
        Args:
            url (str): スクレイピング対象のURL
            
        Returns:
            list: DataFrameのリスト
        """
        try:
            logger.info(f"複数テーブルのスクレイピング開始: {url}")
            
            # ページにアクセス
            self.driver.get(url)
            
            # ページの読み込み完了を待機
            WebDriverWait(self.driver, self.timeout).until(
                EC.presence_of_element_located((By.TAG_NAME, "table"))
            )
            
            time.sleep(3)
            
            # ページソースを取得
            page_source = self.driver.page_source
            soup = BeautifulSoup(page_source, 'html.parser')
            
            # 全てのテーブルを取得
            tables = soup.find_all('table')
            
            dataframes = []
            for i, table in enumerate(tables):
                df = self._parse_table_to_dataframe(table)
                if not df.empty:
                    df.name = f"テーブル{i+1}"
                    dataframes.append(df)
            
            logger.info(f"複数テーブル抽出完了: {len(dataframes)}個のテーブルを取得")
            return dataframes
            
        except Exception as e:
            logger.error(f"複数テーブルスクレイピングエラー: {e}")
            return []


def test_scraper():
    """スクレイピングモジュールのテスト関数"""
    import configparser
    from auth import Authenticator
    
    # テスト用設定
    config = configparser.ConfigParser()
    config.read('config.ini')
    
    # 認証してスクレイピングテスト
    auth = Authenticator(config)
    
    try:
        if auth.login():
            driver = auth.get_driver()
            scraper = KpiScraper(driver, config)
            
            # テストURL（GitHubトレンド）でテスト
            test_url = "https://github.com/trending"
            df = scraper.scrape_table_data(test_url)
            
            if not df.empty:
                print("スクレイピングテスト成功")
                print(f"取得データ: {len(df)}行")
                print(df.head())
            else:
                print("スクレイピングテスト失敗: データが取得できませんでした")
                
    except Exception as e:
        print(f"テストエラー: {e}")
    finally:
        auth.close()


if __name__ == "__main__":
    test_scraper()
