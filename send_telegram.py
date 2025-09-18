import requests
import os
import sys

def send_telegram_message(chat_id, message):
    """Send message to Telegram"""
    TELEGRAM_TOKEN = os.getenv('TELEGRAM_BOT_TOKEN')
    
    if not TELEGRAM_TOKEN:
        print("TELEGRAM_BOT_TOKEN not set")
        return False

    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    data = {
        'chat_id': chat_id,
        'text': message,
        'parse_mode': 'HTML'
    }

    try:
        response = requests.post(url, json=data)
        if response.status_code == 200:
            return True
        else:
            print(f"Failed to send message: {response.text}")
            return False
    except Exception as e:
        print(f"Error sending Telegram message: {e}")
        return False

def send_telegram_document(chat_id, caption, document_path):
    """Send document to Telegram"""
    TELEGRAM_TOKEN = os.getenv('TELEGRAM_BOT_TOKEN')
    
    if not TELEGRAM_TOKEN:
        print("TELEGRAM_BOT_TOKEN not set")
        return False

    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendDocument"
    
    with open(document_path, 'rb') as file:
        files = {'document': file}
        data = {
            'chat_id': chat_id,
            'caption': caption,
            'parse_mode': 'HTML'
        }
        
        try:
            response = requests.post(url, data=data, files=files)
            if response.status_code == 200:
                return True
            else:
                print(f"Failed to send document: {response.text}")
                return False
        except Exception as e:
            print(f"Error sending Telegram document: {e}")
            return False

def main():
    if len(sys.argv) < 3:
        print("Usage: python send_telegram.py <chat_id> <search_parameter>")
        sys.exit(1)
    
    chat_id = sys.argv[1]
    search_parameter = sys.argv[2]
    
    # Send completion message
    message = f"✅ <b>Search Completed</b>\n\n🔍 <b>Search Parameter:</b> {search_parameter}\n📊 <b>Results:</b> See attached file"
    send_telegram_message(chat_id, message)
    
    # Send results file
    send_telegram_document(chat_id, f"📊 Search results for: {search_parameter}", "results.xlsx")

if __name__ == "__main__":
    main()