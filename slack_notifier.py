"""
Slack通知モジュール
Webhookを使用してSlackにメッセージとファイルを送信する
"""

import requests
import os
import json
from datetime import datetime
import logging

# ログ設定
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class SlackNotifier:
    """Slack通知を行うクラス"""
    
    def __init__(self, config):
        """
        初期化
        
        Args:
            config: configparserオブジェクト
        """
        self.config = config
        self.webhook_url = config.get('Slack', 'webhook_url', fallback='')
        self.channel = config.get('Slack', 'channel', fallback='#general')
        self.username = config.get('Slack', 'username', fallback='Web Scraping Bot')
        
    def send_message(self, message, color='good'):
        """
        Slackにテキストメッセージを送信する
        
        Args:
            message (str): 送信するメッセージ
            color (str): メッセージの色 ('good', 'warning', 'danger')
            
        Returns:
            bool: 送信成功の可否
        """
        try:
            if not self.webhook_url:
                logger.warning("Webhook URLが設定されていません")
                return False
            
            # メッセージペイロードを作成
            payload = {
                "channel": self.channel,
                "username": self.username,
                "attachments": [
                    {
                        "color": color,
                        "text": message,
                        "ts": datetime.now().timestamp()
                    }
                ]
            }
            
            # Slackに送信
            response = requests.post(
                self.webhook_url,
                data=json.dumps(payload),
                headers={'Content-Type': 'application/json'},
                timeout=30
            )
            
            if response.status_code == 200:
                logger.info("Slackメッセージ送信成功")
                return True
            else:
                logger.error(f"Slackメッセージ送信失敗: {response.status_code} - {response.text}")
                return False
                
        except Exception as e:
            logger.error(f"Slackメッセージ送信エラー: {e}")
            return False
    
    def send_success_notification(self, excel_filepath, record_count):
        """
        スクレイピング成功通知を送信する
        
        Args:
            excel_filepath (str): 生成されたExcelファイルのパス
            record_count (int): 取得したレコード数
            
        Returns:
            bool: 送信成功の可否
        """
        try:
            filename = os.path.basename(excel_filepath) if excel_filepath else "ファイル未生成"
            
            message = f"""
📊 **KPIデータ取得完了**

✅ **ステータス**: 成功
📈 **取得レコード数**: {record_count}件
📁 **出力ファイル**: {filename}
🕐 **実行時刻**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

KPIデータの取得とExcel出力が正常に完了しました。
            """.strip()
            
            return self.send_message(message, 'good')
            
        except Exception as e:
            logger.error(f"成功通知送信エラー: {e}")
            return False
    
    def send_error_notification(self, error_message):
        """
        エラー通知を送信する
        
        Args:
            error_message (str): エラーメッセージ
            
        Returns:
            bool: 送信成功の可否
        """
        try:
            message = f"""
❌ **KPIデータ取得エラー**

🚨 **ステータス**: 失敗
📝 **エラー内容**: {error_message}
🕐 **実行時刻**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

KPIデータの取得中にエラーが発生しました。
システム管理者にお問い合わせください。
            """.strip()
            
            return self.send_message(message, 'danger')
            
        except Exception as e:
            logger.error(f"エラー通知送信エラー: {e}")
            return False
    
    def send_file_with_message(self, filepath, message):
        """
        ファイル付きメッセージを送信する（Slack API使用）
        
        Note: この機能を使用するにはSlack APIトークンが必要です
        現在の実装ではWebhookのみを使用しているため、
        ファイルアップロードは別途実装が必要です
        
        Args:
            filepath (str): 送信するファイルのパス
            message (str): 添付メッセージ
            
        Returns:
            bool: 送信成功の可否
        """
        try:
            # 現在はWebhookのみの実装のため、ファイル情報をメッセージに含める
            if os.path.exists(filepath):
                file_size = os.path.getsize(filepath)
                file_size_mb = round(file_size / (1024 * 1024), 2)
                
                enhanced_message = f"""
{message}

📎 **添付ファイル情報**
- ファイル名: {os.path.basename(filepath)}
- ファイルサイズ: {file_size_mb} MB
- ファイルパス: {filepath}
                """.strip()
                
                return self.send_message(enhanced_message)
            else:
                logger.warning(f"ファイルが見つかりません: {filepath}")
                return False
                
        except Exception as e:
            logger.error(f"ファイル付きメッセージ送信エラー: {e}")
            return False
    
    def send_daily_summary(self, summary_data):
        """
        日次サマリー通知を送信する
        
        Args:
            summary_data (dict): サマリーデータ
            
        Returns:
            bool: 送信成功の可否
        """
        try:
            message = f"""
📊 **日次KPIサマリー**

📅 **日付**: {datetime.now().strftime('%Y年%m月%d日')}
            """
            
            # サマリーデータを追加
            for key, value in summary_data.items():
                message += f"\n📈 **{key}**: {value}"
            
            message += f"\n\n🕐 **レポート生成時刻**: {datetime.now().strftime('%H:%M:%S')}"
            
            return self.send_message(message, 'good')
            
        except Exception as e:
            logger.error(f"日次サマリー送信エラー: {e}")
            return False
    
    def test_connection(self):
        """
        Slack接続テストを実行する
        
        Returns:
            bool: 接続成功の可否
        """
        try:
            test_message = f"""
🔧 **接続テスト**

Slack通知機能のテストメッセージです。
🕐 **テスト実行時刻**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

このメッセージが表示されれば、Slack通知が正常に動作しています。
            """.strip()
            
            return self.send_message(test_message, 'warning')
            
        except Exception as e:
            logger.error(f"接続テストエラー: {e}")
            return False


class MockSlackNotifier(SlackNotifier):
    """テスト用のモックSlack通知クラス"""
    
    def __init__(self, config):
        super().__init__(config)
        self.sent_messages = []
    
    def send_message(self, message, color='good'):
        """
        メッセージをログに出力する（実際には送信しない）
        
        Args:
            message (str): 送信するメッセージ
            color (str): メッセージの色
            
        Returns:
            bool: 常にTrue
        """
        log_message = f"[MOCK SLACK] Color: {color}\nMessage: {message}"
        logger.info(log_message)
        
        self.sent_messages.append({
            'message': message,
            'color': color,
            'timestamp': datetime.now()
        })
        
        return True
    
    def get_sent_messages(self):
        """送信されたメッセージの履歴を取得する"""
        return self.sent_messages


def test_slack_notifier():
    """Slack通知モジュールのテスト関数"""
    import configparser
    
    # テスト用設定
    config = configparser.ConfigParser()
    config.read('config.ini')
    
    # Webhook URLが設定されていない場合はモックを使用
    webhook_url = config.get('Slack', 'webhook_url', fallback='')
    
    if webhook_url:
        notifier = SlackNotifier(config)
        print("実際のSlackに通知テストを送信します...")
    else:
        notifier = MockSlackNotifier(config)
        print("モック通知でテストを実行します...")
    
    try:
        # 接続テスト
        if notifier.test_connection():
            print("✅ 接続テスト成功")
        else:
            print("❌ 接続テスト失敗")
        
        # 成功通知テスト
        if notifier.send_success_notification("test_file.xlsx", 100):
            print("✅ 成功通知テスト成功")
        else:
            print("❌ 成功通知テスト失敗")
        
        # エラー通知テスト
        if notifier.send_error_notification("テストエラーメッセージ"):
            print("✅ エラー通知テスト成功")
        else:
            print("❌ エラー通知テスト失敗")
        
        # サマリー通知テスト
        summary_data = {
            "総売上": "¥1,000,000",
            "訪問者数": "5,000人",
            "コンバージョン率": "2.5%"
        }
        
        if notifier.send_daily_summary(summary_data):
            print("✅ サマリー通知テスト成功")
        else:
            print("❌ サマリー通知テスト失敗")
        
        # モックの場合は送信履歴を表示
        if isinstance(notifier, MockSlackNotifier):
            print(f"\n📝 送信メッセージ履歴: {len(notifier.get_sent_messages())}件")
            
    except Exception as e:
        print(f"テストエラー: {e}")


if __name__ == "__main__":
    test_slack_notifier()
