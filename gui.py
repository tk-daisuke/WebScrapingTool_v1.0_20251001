"""
GUI設定画面モジュール
PySimpleGUIを使用して設定画面を提供する
"""

import PySimpleGUI as sg
import configparser
import os
import threading
import time
from datetime import datetime
import logging

# ログ設定
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class SettingsGUI:
    """設定用GUIクラス"""
    
    def __init__(self, config_file='config.ini'):
        """
        初期化
        
        Args:
            config_file (str): 設定ファイルのパス
        """
        self.config_file = config_file
        self.config = configparser.ConfigParser()
        self.load_config()
        
        # PySimpleGUIのテーマ設定
        sg.theme('LightBlue3')
        
    def load_config(self):
        """設定ファイルを読み込む"""
        try:
            if os.path.exists(self.config_file):
                self.config.read(self.config_file, encoding='utf-8')
            else:
                # デフォルト設定を作成
                self.create_default_config()
                
        except Exception as e:
            logger.error(f"設定ファイル読み込みエラー: {e}")
            self.create_default_config()
    
    def create_default_config(self):
        """デフォルト設定を作成する"""
        self.config['Scraper'] = {
            'target_url': 'https://example.com/kpi-dashboard',
            'login_url': 'https://example.com/login',
            'username': '',
            'password': ''
        }
        
        self.config['Excel'] = {
            'output_filename': 'KPI_data_{timestamp}.xlsx',
            'output_directory': './output'
        }
        
        self.config['Slack'] = {
            'webhook_url': '',
            'channel': '#general',
            'username': 'Web Scraping Bot'
        }
        
        self.config['Browser'] = {
            'headless': 'False',
            'timeout': '30',
            'implicit_wait': '10'
        }
    
    def save_config(self):
        """設定ファイルを保存する"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                self.config.write(f)
            logger.info("設定ファイルを保存しました")
            return True
            
        except Exception as e:
            logger.error(f"設定ファイル保存エラー: {e}")
            return False
    
    def create_main_window(self):
        """メイン設定ウィンドウを作成する"""
        
        # スクレイピング設定タブ
        scraper_layout = [
            [sg.Text('KPIページURL:', size=(15, 1)), 
             sg.Input(self.config.get('Scraper', 'target_url', fallback=''), key='-TARGET_URL-', size=(50, 1))],
            [sg.Text('ログインページURL:', size=(15, 1)), 
             sg.Input(self.config.get('Scraper', 'login_url', fallback=''), key='-LOGIN_URL-', size=(50, 1))],
            [sg.Text('ユーザー名:', size=(15, 1)), 
             sg.Input(self.config.get('Scraper', 'username', fallback=''), key='-USERNAME-', size=(30, 1))],
            [sg.Text('パスワード:', size=(15, 1)), 
             sg.Input(self.config.get('Scraper', 'password', fallback=''), key='-PASSWORD-', size=(30, 1), password_char='*')],
            [sg.HSeparator()],
            [sg.Text('注意: USBセキュリティキーによる認証は手動で行う必要があります。', text_color='red')]
        ]
        
        # Excel設定タブ
        excel_layout = [
            [sg.Text('出力ファイル名:', size=(15, 1)), 
             sg.Input(self.config.get('Excel', 'output_filename', fallback=''), key='-OUTPUT_FILENAME-', size=(40, 1))],
            [sg.Text('出力ディレクトリ:', size=(15, 1)), 
             sg.Input(self.config.get('Excel', 'output_directory', fallback=''), key='-OUTPUT_DIR-', size=(35, 1)),
             sg.FolderBrowse('参照', target='-OUTPUT_DIR-')],
            [sg.HSeparator()],
            [sg.Text('ファイル名に使用可能な変数:')],
            [sg.Text('  {timestamp} - 実行時刻 (例: 20241001_143000)', font=('Courier', 9))]
        ]
        
        # Slack設定タブ
        slack_layout = [
            [sg.Text('Webhook URL:', size=(15, 1)), 
             sg.Input(self.config.get('Slack', 'webhook_url', fallback=''), key='-WEBHOOK_URL-', size=(50, 1))],
            [sg.Text('チャンネル:', size=(15, 1)), 
             sg.Input(self.config.get('Slack', 'channel', fallback=''), key='-CHANNEL-', size=(20, 1))],
            [sg.Text('ボット名:', size=(15, 1)), 
             sg.Input(self.config.get('Slack', 'username', fallback=''), key='-BOT_USERNAME-', size=(30, 1))],
            [sg.HSeparator()],
            [sg.Button('接続テスト', key='-TEST_SLACK-'), sg.Text('', key='-SLACK_STATUS-', size=(30, 1))]
        ]
        
        # ブラウザ設定タブ
        browser_layout = [
            [sg.Checkbox('ヘッドレスモード', 
                        default=self.config.getboolean('Browser', 'headless', fallback=False), 
                        key='-HEADLESS-')],
            [sg.Text('タイムアウト(秒):', size=(15, 1)), 
             sg.Input(self.config.get('Browser', 'timeout', fallback='30'), key='-TIMEOUT-', size=(10, 1))],
            [sg.Text('暗黙的待機(秒):', size=(15, 1)), 
             sg.Input(self.config.get('Browser', 'implicit_wait', fallback='10'), key='-IMPLICIT_WAIT-', size=(10, 1))],
            [sg.HSeparator()],
            [sg.Text('ヘッドレスモード: ブラウザウィンドウを表示せずに実行')]
        ]
        
        # タブレイアウト
        tab_group = [
            [sg.Tab('スクレイピング', scraper_layout, key='-TAB_SCRAPER-')],
            [sg.Tab('Excel出力', excel_layout, key='-TAB_EXCEL-')],
            [sg.Tab('Slack通知', slack_layout, key='-TAB_SLACK-')],
            [sg.Tab('ブラウザ', browser_layout, key='-TAB_BROWSER-')]
        ]
        
        # メインレイアウト
        layout = [
            [sg.Text('ウェブスクレイピングツール設定', font=('Arial', 16, 'bold'))],
            [sg.HSeparator()],
            [sg.TabGroup(tab_group, enable_events=True)],
            [sg.HSeparator()],
            [sg.Button('保存', key='-SAVE-', size=(10, 1)), 
             sg.Button('キャンセル', key='-CANCEL-', size=(10, 1)),
             sg.Button('実行', key='-RUN-', size=(10, 1), button_color=('white', 'green')),
             sg.Text('', key='-STATUS-', size=(30, 1))]
        ]
        
        return sg.Window('ウェブスクレイピングツール設定', layout, finalize=True, resizable=True)
    
    def create_progress_window(self):
        """進行状況表示ウィンドウを作成する"""
        layout = [
            [sg.Text('スクレイピング実行中...', font=('Arial', 12))],
            [sg.ProgressBar(100, orientation='h', size=(40, 20), key='-PROGRESS-')],
            [sg.Text('', key='-PROGRESS_TEXT-', size=(50, 1))],
            [sg.Multiline('', key='-LOG-', size=(60, 10), disabled=True, autoscroll=True)],
            [sg.Button('キャンセル', key='-CANCEL_RUN-')]
        ]
        
        return sg.Window('実行中', layout, finalize=True, modal=True)
    
    def update_config_from_values(self, values):
        """GUIの値から設定を更新する"""
        try:
            # スクレイピング設定
            self.config.set('Scraper', 'target_url', values['-TARGET_URL-'])
            self.config.set('Scraper', 'login_url', values['-LOGIN_URL-'])
            self.config.set('Scraper', 'username', values['-USERNAME-'])
            self.config.set('Scraper', 'password', values['-PASSWORD-'])
            
            # Excel設定
            self.config.set('Excel', 'output_filename', values['-OUTPUT_FILENAME-'])
            self.config.set('Excel', 'output_directory', values['-OUTPUT_DIR-'])
            
            # Slack設定
            self.config.set('Slack', 'webhook_url', values['-WEBHOOK_URL-'])
            self.config.set('Slack', 'channel', values['-CHANNEL-'])
            self.config.set('Slack', 'username', values['-BOT_USERNAME-'])
            
            # ブラウザ設定
            self.config.set('Browser', 'headless', str(values['-HEADLESS-']))
            self.config.set('Browser', 'timeout', values['-TIMEOUT-'])
            self.config.set('Browser', 'implicit_wait', values['-IMPLICIT_WAIT-'])
            
            return True
            
        except Exception as e:
            logger.error(f"設定更新エラー: {e}")
            return False
    
    def test_slack_connection(self, values):
        """Slack接続をテストする"""
        try:
            # 一時的に設定を更新
            temp_config = configparser.ConfigParser()
            temp_config.read_dict(self.config)
            temp_config.set('Slack', 'webhook_url', values['-WEBHOOK_URL-'])
            temp_config.set('Slack', 'channel', values['-CHANNEL-'])
            temp_config.set('Slack', 'username', values['-BOT_USERNAME-'])
            
            # Slack通知テスト
            from slack_notifier import SlackNotifier, MockSlackNotifier
            
            if values['-WEBHOOK_URL-'].strip():
                notifier = SlackNotifier(temp_config)
            else:
                notifier = MockSlackNotifier(temp_config)
            
            return notifier.test_connection()
            
        except Exception as e:
            logger.error(f"Slack接続テストエラー: {e}")
            return False
    
    def run_scraping(self, progress_window):
        """スクレイピングを実行する（別スレッドで実行）"""
        try:
            from main import run_scraping_process
            
            # 進行状況の更新
            progress_window['-PROGRESS_TEXT-'].update('認証中...')
            progress_window['-PROGRESS-'].update(20)
            progress_window['-LOG-'].update('認証処理を開始しています...\n', append=True)
            
            # メイン処理を実行
            result = run_scraping_process(self.config, progress_window)
            
            if result:
                progress_window['-PROGRESS_TEXT-'].update('完了')
                progress_window['-PROGRESS-'].update(100)
                progress_window['-LOG-'].update('スクレイピングが正常に完了しました。\n', append=True)
            else:
                progress_window['-PROGRESS_TEXT-'].update('エラー')
                progress_window['-LOG-'].update('スクレイピング中にエラーが発生しました。\n', append=True)
            
        except Exception as e:
            logger.error(f"スクレイピング実行エラー: {e}")
            progress_window['-PROGRESS_TEXT-'].update('エラー')
            progress_window['-LOG-'].update(f'エラー: {e}\n', append=True)
    
    def run(self):
        """GUIのメインループを実行する"""
        window = self.create_main_window()
        
        while True:
            event, values = window.read()
            
            if event in (sg.WIN_CLOSED, '-CANCEL-'):
                break
            
            elif event == '-SAVE-':
                if self.update_config_from_values(values):
                    if self.save_config():
                        window['-STATUS-'].update('設定を保存しました', text_color='green')
                    else:
                        window['-STATUS-'].update('保存に失敗しました', text_color='red')
                else:
                    window['-STATUS-'].update('設定の更新に失敗しました', text_color='red')
            
            elif event == '-TEST_SLACK-':
                window['-SLACK_STATUS-'].update('テスト中...', text_color='blue')
                window.refresh()
                
                if self.test_slack_connection(values):
                    window['-SLACK_STATUS-'].update('接続成功', text_color='green')
                else:
                    window['-SLACK_STATUS-'].update('接続失敗', text_color='red')
            
            elif event == '-RUN-':
                # 設定を保存してからスクレイピングを実行
                if self.update_config_from_values(values):
                    self.save_config()
                    
                    # 進行状況ウィンドウを表示
                    progress_window = self.create_progress_window()
                    
                    # 別スレッドでスクレイピングを実行
                    thread = threading.Thread(target=self.run_scraping, args=(progress_window,))
                    thread.daemon = True
                    thread.start()
                    
                    # 進行状況ウィンドウのイベントループ
                    while True:
                        prog_event, prog_values = progress_window.read(timeout=100)
                        
                        if prog_event in (sg.WIN_CLOSED, '-CANCEL_RUN-'):
                            break
                        
                        if not thread.is_alive():
                            time.sleep(1)  # 最終メッセージを表示するため少し待機
                            break
                    
                    progress_window.close()
        
        window.close()


def test_gui():
    """GUI設定画面のテスト関数"""
    try:
        gui = SettingsGUI()
        gui.run()
        
    except Exception as e:
        print(f"GUIテストエラー: {e}")


if __name__ == "__main__":
    test_gui()
