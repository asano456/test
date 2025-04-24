import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr

# ==== 設定 ====
EXCEL_PATH = 'list.xlsx'
SHEET_NAME = 'Sheet1'
HTML_TEMPLATE_PATH = '20240422_骨再建プレート.html'

# お名前.com のSMTP情報（必要に応じて調整）
SMTP_SERVER = 'smtp01.gmoserver.jp'
SMTP_PORT = 587

# 差出人情報
EMAIL_ADDRESS = 'asano@next21.info'
EMAIL_PASSWORD = 'Kotaro4351'  # ← 本物のパスワードに置き換える
DISPLAY_NAME = 'ネクスト21'

# ==== HTML読み込み ====
with open(HTML_TEMPLATE_PATH, 'r', encoding='utf-8') as f:
    html_template = f.read()

# ==== Excel読み込み ====
df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)
df = df.dropna(subset=['Email'])  # Emailが空の行は除外

# ==== SMTP接続（STARTTLS使用） ====
server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
server.starttls()
server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)

# ==== メール送信 ====
for _, row in df.iterrows():
    to_email = row['Email']
    name = row['Name'] if 'Name' in row else ''

    personalized_body = html_template.replace('{{name}}', str(name))

    msg = MIMEMultipart('alternative')
    msg['Subject'] = "【ご案内】骨再建プレートのご紹介"
    msg['From'] = formataddr((DISPLAY_NAME, EMAIL_ADDRESS))
    msg['To'] = to_email
    msg.attach(MIMEText(personalized_body, 'html'))

    try:
        server.sendmail(EMAIL_ADDRESS, to_email, msg.as_string())
        print(f"✅ 送信成功: {to_email}")
    except Exception as e:
        print(f"❌ 送信失敗: {to_email} - {e}")

server.quit()
