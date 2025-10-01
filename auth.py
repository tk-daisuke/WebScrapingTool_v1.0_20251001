"""
認証とCookie管理モジュール
USBセキュリティキーを用いた認証の手動実行とCookieの永続化を管理する
"""

import os
import pickle
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import logging

# ログ設定
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class Authenticator:
    """認証とCookie管理を行うクラス"""
    
    def __init__(self, config):
        """
        初期化
        
        Args:
            config: configparserオブジェクト
        """
        self.config = config
        self.driver = None
        self.cookies_file = "cookies.pkl"
        
    def setup_driver(self):
        """WebDriverを設定・起動する"""
        try:
            # Chrome オプションの設定
            chrome_options = Options()
            
            # ヘッドレスモードの設定
            if self.config.getboolean('Browser', 'headless', fallback=False):
                chrome_options.add_argument('--headless')
            
            # その他のオプション
            chrome_options.add_argument('--no-sandbox')
            chrome_options.add_argument('--disable-dev-shm-usage')
            chrome_options.add_argument('--disable-gpu')
            chrome_options.add_argument('--window-size=1920,1080')
            
            # WebDriverManagerを使用してChromeDriverを自動管理
            service = Service(ChromeDriverManager().install())
            
            # WebDriverを起動
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
            
            # 暗黙的な待機時間を設定
            implicit_wait = self.config.getint('Browser', 'implicit_wait', fallback=10)
            self.driver.implicitly_wait(implicit_wait)
            
            logger.info("WebDriverが正常に起動しました")
            return True
            
        except Exception as e:
            logger.error(f"WebDriverの起動に失敗しました: {e}")
            return False
    
    def load_cookies(self):
        """保存されたCookieを読み込む"""
        try:
            if os.path.exists(self.cookies_file):
                with open(self.cookies_file, 'rb') as f:
                    cookies = pickle.load(f)
                
                # まず対象サイトにアクセス
                login_url = self.config.get('Scraper', 'login_url')
                self.driver.get(login_url)
                
                # Cookieを設定
                for cookie in cookies:
                    try:
                        self.driver.add_cookie(cookie)
                    except Exception as e:
                        logger.warning(f"Cookie設定エラー: {e}")
                
                logger.info("保存されたCookieを読み込みました")
                return True
            else:
                logger.info("保存されたCookieが見つかりません")
                return False
                
        except Exception as e:
            logger.error(f"Cookie読み込みエラー: {e}")
            return False
    
    def save_cookies(self):
        """現在のセッションのCookieを保存する"""
        try:
            cookies = self.driver.get_cookies()
            with open(self.cookies_file, 'wb') as f:
                pickle.dump(cookies, f)
            logger.info("Cookieを保存しました")
            return True
            
        except Exception as e:
            logger.error(f"Cookie保存エラー: {e}")
            return False
    
    def manual_login(self):
        """手動ログインを実行する"""
        try:
            login_url = self.config.get('Scraper', 'login_url')
            self.driver.get(login_url)
            
            print("\n" + "="*60)
            print("手動ログインが必要です")
            print("="*60)
            print("1. ブラウザでログインページが開かれました")
            print("2. ユーザー名とパスワードを入力してください")
            print("3. USBセキュリティキーによる認証を完了してください")
            print("4. ログインが完了したら、このプログラムでEnterキーを押してください")
            print("="*60)
            
            # ユーザーの入力を待つ
            input("ログイン完了後、Enterキーを押してください...")
            
            # ログイン成功の確認（URLの変化やページ要素で判定）
            current_url = self.driver.current_url
            if current_url != login_url:
                logger.info("ログインが成功したと判定されました")
                self.save_cookies()
                return True
            else:
                logger.warning("ログインが完了していない可能性があります")
                return False
                
        except Exception as e:
            logger.error(f"手動ログインエラー: {e}")
            return False
    
    def login(self):
        """
        ログイン処理のメイン関数
        保存されたCookieがあれば使用し、なければ手動ログインを実行
        
        Returns:
            bool: ログイン成功の可否
        """
        try:
            # WebDriverを起動
            if not self.setup_driver():
                return False
            
            # 保存されたCookieでのログインを試行
            if self.load_cookies():
                # ログイン状態の確認
                target_url = self.config.get('Scraper', 'target_url')
                self.driver.get(target_url)
                
                # ページタイトルやURLでログイン状態を確認
                time.sleep(3)  # ページ読み込み待機
                
                if "login" not in self.driver.current_url.lower():
                    logger.info("Cookieによるログインが成功しました")
                    return True
                else:
                    logger.info("Cookieが無効です。手動ログインを実行します")
            
            # 手動ログインを実行
            return self.manual_login()
            
        except Exception as e:
            logger.error(f"ログイン処理エラー: {e}")
            return False
    
    def get_driver(self):
        """WebDriverインスタンスを取得する"""
        return self.driver
    
    def close(self):
        """WebDriverを終了する"""
        if self.driver:
            self.driver.quit()
            logger.info("WebDriverを終了しました")


def test_authenticator():
    """認証モジュールのテスト関数"""
    import configparser
    
    # テスト用設定
    config = configparser.ConfigParser()
    config.read('config.ini')
    
    # 認証テスト
    auth = Authenticator(config)
    
    try:
        if auth.login():
            print("認証テスト成功")
            time.sleep(5)  # 5秒待機
        else:
            print("認証テスト失敗")
    finally:
        auth.close()


if __name__ == "__main__":
    test_authenticator()
