"""
メイン実行モジュール
全ての機能を統合してスクレイピング処理を実行する
"""

import configparser
import os
import sys
import logging
from datetime import datetime
import traceback

# 自作モジュールのインポート
from auth import Authenticator
from scraper import KpiScraper
from excel_writer import ExcelWriter
from slack_notifier import SlackNotifier, MockSlackNotifier
from gui import SettingsGUI

# ログ設定
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('scraping.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


def load_config(config_file='config.ini'):
    """
    設定ファイルを読み込む
    
    Args:
        config_file (str): 設定ファイルのパス
        
    Returns:
        configparser.ConfigParser: 設定オブジェクト
    """
    config = configparser.ConfigParser()
    
    try:
        if os.path.exists(config_file):
            config.read(config_file, encoding='utf-8')
            logger.info(f"設定ファイルを読み込みました: {config_file}")
        else:
            logger.warning(f"設定ファイルが見つかりません: {config_file}")
            # GUIで設定を作成
            gui = SettingsGUI(config_file)
            gui.run()
            config.read(config_file, encoding='utf-8')
            
    except Exception as e:
        logger.error(f"設定ファイル読み込みエラー: {e}")
        raise
    
    return config


def run_scraping_process(config, progress_window=None):
    """
    スクレイピング処理のメイン関数
    
    Args:
        config: configparserオブジェクト
        progress_window: 進行状況表示ウィンドウ（オプション）
        
    Returns:
        bool: 処理成功の可否
    """
    auth = None
    
    try:
        logger.info("スクレイピング処理を開始します")
        
        # 進行状況更新
        if progress_window:
            progress_window['-PROGRESS_TEXT-'].update('認証中...')
            progress_window['-PROGRESS-'].update(10)
            progress_window['-LOG-'].update('認証処理を開始しています...\n', append=True)
        
        # 1. 認証処理
        auth = Authenticator(config)
        if not auth.login():
            raise Exception("認証に失敗しました")
        
        logger.info("認証が完了しました")
        
        # 進行状況更新
        if progress_window:
            progress_window['-PROGRESS_TEXT-'].update('データ取得中...')
            progress_window['-PROGRESS-'].update(30)
            progress_window['-LOG-'].update('KPIデータを取得しています...\n', append=True)
        
        # 2. スクレイピング処理
        driver = auth.get_driver()
        scraper = KpiScraper(driver, config)
        df = scraper.scrape_kpi_data()
        
        if df.empty:
            raise Exception("データが取得できませんでした")
        
        logger.info(f"データ取得完了: {len(df)}行")
        
        # 進行状況更新
        if progress_window:
            progress_window['-PROGRESS_TEXT-'].update('Excel出力中...')
            progress_window['-PROGRESS-'].update(60)
            progress_window['-LOG-'].update(f'{len(df)}行のデータを取得しました。Excelファイルを作成中...\n', append=True)
        
        # 3. Excel出力
        excel_writer = ExcelWriter(config)
        excel_filepath = excel_writer.write_to_excel(df)
        
        if not excel_filepath:
            raise Exception("Excelファイルの出力に失敗しました")
        
        logger.info(f"Excelファイル出力完了: {excel_filepath}")
        
        # 進行状況更新
        if progress_window:
            progress_window['-PROGRESS_TEXT-'].update('Slack通知中...')
            progress_window['-PROGRESS-'].update(80)
            progress_window['-LOG-'].update(f'Excelファイルを出力しました: {os.path.basename(excel_filepath)}\n', append=True)
        
        # 4. Slack通知
        webhook_url = config.get('Slack', 'webhook_url', fallback='')
        
        if webhook_url.strip():
            notifier = SlackNotifier(config)
        else:
            notifier = MockSlackNotifier(config)
            logger.info("Webhook URLが未設定のため、モック通知を使用します")
        
        if notifier.send_success_notification(excel_filepath, len(df)):
            logger.info("Slack通知送信完了")
        else:
            logger.warning("Slack通知の送信に失敗しました")
        
        # 進行状況更新
        if progress_window:
            progress_window['-PROGRESS_TEXT-'].update('完了')
            progress_window['-PROGRESS-'].update(100)
            progress_window['-LOG-'].update('すべての処理が正常に完了しました。\n', append=True)
        
        logger.info("スクレイピング処理が正常に完了しました")
        return True
        
    except Exception as e:
        error_message = f"スクレイピング処理エラー: {e}"
        logger.error(error_message)
        logger.error(traceback.format_exc())
        
        # エラー通知
        try:
            webhook_url = config.get('Slack', 'webhook_url', fallback='')
            
            if webhook_url.strip():
                notifier = SlackNotifier(config)
            else:
                notifier = MockSlackNotifier(config)
            
            notifier.send_error_notification(str(e))
            
        except Exception as notify_error:
            logger.error(f"エラー通知送信失敗: {notify_error}")
        
        # 進行状況更新
        if progress_window:
            progress_window['-PROGRESS_TEXT-'].update('エラー')
            progress_window['-LOG-'].update(f'エラーが発生しました: {e}\n', append=True)
        
        return False
        
    finally:
        # リソースのクリーンアップ
        if auth:
            auth.close()


def main():
    """メイン関数"""
    try:
        print("=" * 60)
        print("ウェブスクレイピングツール")
        print("=" * 60)
        
        # コマンドライン引数の確認
        if len(sys.argv) > 1:
            if sys.argv[1] == '--gui':
                # GUI モードで起動
                gui = SettingsGUI()
                gui.run()
                return
            elif sys.argv[1] == '--config':
                # 設定ファイルのパスを指定
                config_file = sys.argv[2] if len(sys.argv) > 2 else 'config.ini'
            else:
                config_file = 'config.ini'
        else:
            # デフォルトでGUIモードを起動
            gui = SettingsGUI()
            gui.run()
            return
        
        # 設定ファイルを読み込み
        config = load_config(config_file)
        
        # スクレイピング処理を実行
        success = run_scraping_process(config)
        
        if success:
            print("\n✅ スクレイピング処理が正常に完了しました")
        else:
            print("\n❌ スクレイピング処理中にエラーが発生しました")
            sys.exit(1)
        
    except KeyboardInterrupt:
        print("\n\n処理が中断されました")
        sys.exit(1)
        
    except Exception as e:
        print(f"\n❌ 予期しないエラーが発生しました: {e}")
        logger.error(f"メイン処理エラー: {e}")
        logger.error(traceback.format_exc())
        sys.exit(1)


def run_sample():
    """サンプル実行関数（テスト用）"""
    try:
        print("サンプルスクレイピングを実行します...")
        
        # サンプル用設定を作成
        config = configparser.ConfigParser()
        config['Scraper'] = {
            'target_url': 'https://github.com/trending',
            'login_url': 'https://github.com/login',
            'username': '',
            'password': ''
        }
        
        config['Excel'] = {
            'output_filename': 'sample_data_{timestamp}.xlsx',
            'output_directory': './output'
        }
        
        config['Slack'] = {
            'webhook_url': '',
            'channel': '#general',
            'username': 'Sample Bot'
        }
        
        config['Browser'] = {
            'headless': 'False',
            'timeout': '30',
            'implicit_wait': '10'
        }
        
        # サンプル処理を実行
        return run_scraping_process(config)
        
    except Exception as e:
        print(f"サンプル実行エラー: {e}")
        return False


if __name__ == "__main__":
    main()
