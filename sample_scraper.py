"""
サンプルスクレイピングコード
GitHubトレンドページを対象として、ツールの主要機能をデモンストレーションする
"""

import configparser
import pandas as pd
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import logging

# 自作モジュールのインポート
from excel_writer import ExcelWriter
from slack_notifier import MockSlackNotifier

# ログ設定
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class GitHubTrendScraper:
    """GitHubトレンドページのスクレイピングを行うサンプルクラス"""
    
    def __init__(self):
        """初期化"""
        self.driver = None
        self.setup_driver()
    
    def setup_driver(self):
        """WebDriverを設定・起動する"""
        try:
            # Chrome オプションの設定
            chrome_options = Options()
            chrome_options.add_argument('--no-sandbox')
            chrome_options.add_argument('--disable-dev-shm-usage')
            chrome_options.add_argument('--disable-gpu')
            chrome_options.add_argument('--window-size=1920,1080')
            
            # WebDriverManagerを使用してChromeDriverを自動管理
            service = Service(ChromeDriverManager().install())
            
            # WebDriverを起動
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
            self.driver.implicitly_wait(10)
            
            logger.info("WebDriverが正常に起動しました")
            return True
            
        except Exception as e:
            logger.error(f"WebDriverの起動に失敗しました: {e}")
            return False
    
    def scrape_trending_repositories(self):
        """
        GitHubトレンドページからリポジトリ情報を取得する
        
        Returns:
            pd.DataFrame: トレンドリポジトリのデータ
        """
        try:
            logger.info("GitHubトレンドページにアクセスしています...")
            
            # GitHubトレンドページにアクセス
            self.driver.get("https://github.com/trending")
            
            # ページの読み込み完了を待機
            WebDriverWait(self.driver, 30).until(
                EC.presence_of_element_located((By.CLASS_NAME, "Box-row"))
            )
            
            time.sleep(3)  # 追加の待機
            
            # ページソースを取得
            page_source = self.driver.page_source
            soup = BeautifulSoup(page_source, 'html.parser')
            
            # トレンドリポジトリの情報を抽出
            repositories = []
            
            # リポジトリのコンテナを検索
            repo_containers = soup.find_all('article', class_='Box-row')
            
            for container in repo_containers[:20]:  # 上位20件を取得
                try:
                    repo_data = self._extract_repository_data(container)
                    if repo_data:
                        repositories.append(repo_data)
                        
                except Exception as e:
                    logger.warning(f"リポジトリデータ抽出エラー: {e}")
                    continue
            
            # DataFrameに変換
            df = pd.DataFrame(repositories)
            
            logger.info(f"トレンドリポジトリデータ取得完了: {len(df)}件")
            return df
            
        except Exception as e:
            logger.error(f"GitHubトレンドスクレイピングエラー: {e}")
            return pd.DataFrame()
    
    def _extract_repository_data(self, container):
        """
        リポジトリコンテナから個別のデータを抽出する
        
        Args:
            container: BeautifulSoupの要素
            
        Returns:
            dict: リポジトリデータ
        """
        try:
            # リポジトリ名とURL
            title_element = container.find('h2', class_='h3')
            if not title_element:
                return None
            
            repo_link = title_element.find('a')
            if not repo_link:
                return None
            
            repo_name = repo_link.get_text(strip=True).replace('\n', '').replace(' ', '')
            repo_url = f"https://github.com{repo_link.get('href', '')}"
            
            # 説明文
            description_element = container.find('p', class_='col-9')
            description = description_element.get_text(strip=True) if description_element else ""
            
            # プログラミング言語
            language_element = container.find('span', {'itemprop': 'programmingLanguage'})
            language = language_element.get_text(strip=True) if language_element else "不明"
            
            # スター数
            stars_element = container.find('a', href=lambda x: x and '/stargazers' in x)
            stars = self._extract_number(stars_element.get_text(strip=True)) if stars_element else 0
            
            # フォーク数
            forks_element = container.find('a', href=lambda x: x and '/forks' in x)
            forks = self._extract_number(forks_element.get_text(strip=True)) if forks_element else 0
            
            # 今日のスター数
            today_stars_element = container.find('span', class_='d-inline-block')
            today_stars = 0
            if today_stars_element:
                today_stars_text = today_stars_element.get_text(strip=True)
                today_stars = self._extract_number(today_stars_text)
            
            return {
                'リポジトリ名': repo_name,
                'URL': repo_url,
                '説明': description,
                'プログラミング言語': language,
                'スター数': stars,
                'フォーク数': forks,
                '今日のスター数': today_stars,
                '取得日時': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
            
        except Exception as e:
            logger.warning(f"個別リポジトリデータ抽出エラー: {e}")
            return None
    
    def _extract_number(self, text):
        """
        テキストから数値を抽出する（k, M などの単位に対応）
        
        Args:
            text (str): 数値を含むテキスト
            
        Returns:
            int: 抽出された数値
        """
        try:
            # 数字以外の文字を除去
            import re
            
            # k, M などの単位を処理
            text = text.replace(',', '')
            
            if 'k' in text.lower():
                number = float(re.findall(r'[\d.]+', text)[0])
                return int(number * 1000)
            elif 'm' in text.lower():
                number = float(re.findall(r'[\d.]+', text)[0])
                return int(number * 1000000)
            else:
                numbers = re.findall(r'\d+', text)
                return int(numbers[0]) if numbers else 0
                
        except Exception:
            return 0
    
    def close(self):
        """WebDriverを終了する"""
        if self.driver:
            self.driver.quit()
            logger.info("WebDriverを終了しました")


def run_sample_scraping():
    """サンプルスクレイピングを実行する"""
    scraper = None
    
    try:
        print("=" * 60)
        print("GitHubトレンドスクレイピング サンプル実行")
        print("=" * 60)
        
        # スクレイピング実行
        scraper = GitHubTrendScraper()
        df = scraper.scrape_trending_repositories()
        
        if df.empty:
            print("❌ データが取得できませんでした")
            return False
        
        print(f"✅ {len(df)}件のトレンドリポジトリを取得しました")
        
        # データの表示
        print("\n📊 取得データのサンプル:")
        print(df[['リポジトリ名', 'プログラミング言語', 'スター数', '今日のスター数']].head())
        
        # Excel出力
        print("\n📁 Excelファイルを出力しています...")
        
        # サンプル用設定を作成
        config = configparser.ConfigParser()
        config['Excel'] = {
            'output_filename': 'github_trends_{timestamp}.xlsx',
            'output_directory': './output'
        }
        
        excel_writer = ExcelWriter(config)
        excel_filepath = excel_writer.write_to_excel(df, sheet_name='GitHubトレンド')
        
        if excel_filepath:
            print(f"✅ Excelファイルを出力しました: {excel_filepath}")
        else:
            print("❌ Excelファイルの出力に失敗しました")
            return False
        
        # Slack通知（モック）
        print("\n📢 Slack通知をテストしています...")
        
        config['Slack'] = {
            'webhook_url': '',
            'channel': '#general',
            'username': 'GitHub Trend Bot'
        }
        
        notifier = MockSlackNotifier(config)
        
        if notifier.send_success_notification(excel_filepath, len(df)):
            print("✅ Slack通知テスト完了")
        else:
            print("❌ Slack通知テスト失敗")
        
        # サマリー情報
        print("\n📈 データサマリー:")
        print(f"  - 総リポジトリ数: {len(df)}件")
        print(f"  - 最も人気の言語: {df['プログラミング言語'].mode().iloc[0] if not df.empty else '不明'}")
        print(f"  - 平均スター数: {df['スター数'].mean():.0f}")
        print(f"  - 今日の総スター数: {df['今日のスター数'].sum()}")
        
        print("\n🎉 サンプルスクレイピングが正常に完了しました！")
        return True
        
    except Exception as e:
        print(f"\n❌ サンプル実行エラー: {e}")
        logger.error(f"サンプル実行エラー: {e}")
        return False
        
    finally:
        if scraper:
            scraper.close()


def test_individual_modules():
    """個別モジュールのテストを実行する"""
    print("=" * 60)
    print("個別モジュールテスト")
    print("=" * 60)
    
    # Excel出力テスト
    print("\n1. Excel出力モジュールテスト")
    try:
        from excel_writer import test_excel_writer
        test_excel_writer()
        print("✅ Excel出力テスト完了")
    except Exception as e:
        print(f"❌ Excel出力テストエラー: {e}")
    
    # Slack通知テスト
    print("\n2. Slack通知モジュールテスト")
    try:
        from slack_notifier import test_slack_notifier
        test_slack_notifier()
        print("✅ Slack通知テスト完了")
    except Exception as e:
        print(f"❌ Slack通知テストエラー: {e}")


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1 and sys.argv[1] == '--test':
        test_individual_modules()
    else:
        run_sample_scraping()
