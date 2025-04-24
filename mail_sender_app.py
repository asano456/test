import streamlit as st
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr
from email.mime.base import MIMEBase
from email import encoders
import datetime
import csv
import os
import matplotlib.pyplot as plt
from io import BytesIO
import time

st.set_page_config(page_title="ネクスト21 メール送信ツール", layout="centered")

st.markdown("""
<style>
h1 {
    color: #1f77b4;
    font-family: 'Segoe UI', sans-serif;
}
.block-container {
    padding-top: 1.5rem;
    padding-bottom: 1.5rem;
}
footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

if os.path.exists("logo.png"):
    st.image("logo.png", width=120)
else:
    st.warning("⚠️ logo.png が見つかりませんでした。")
st.title("📨 ネクスト21 メール一括送信ツール")

# ファイルアップロード
st.markdown("### 1. ファイルのアップロード")
excel_file = st.file_uploader("📎 メールリスト（Excel）", type=['xlsx'])
html_file = st.file_uploader("📎 HTMLテンプレート", type=['html'])

if excel_file:
    st.success(f"📘 Excelファイル読込完了：{excel_file.name}")
if html_file:
    st.success(f"📝 HTMLテンプレート読込完了：{html_file.name}")

# 差出人・設定
st.markdown("### 2. 差出人情報とメール設定")
attachment_file = st.file_uploader("📎 添付ファイル（任意）", type=None)
from_name = st.text_input("表示名", value="ネクスト21")
from_email = st.text_input("送信元メールアドレス", value="asano@next21.info")
from_pass = st.text_input("メールパスワード（非表示）", type="password")
subject = st.text_input("件名", value="【ご案内】骨再建プレートのご紹介")

if html_file:
    html_template_raw = html_file.read().decode('utf-8')
    html_template_display = html_template_raw.replace('{{name}}', '〇〇様')
    st.markdown("### 📄 メール本文プレビュー")
    st.components.v1.html(html_template_display, height=400, scrolling=True)

import hashlib  # ← 差し込みID用のハッシュ化

# メイン処理
if excel_file and html_file:
    df = pd.read_excel(excel_file)
    id_map = {row['Email'].lower(): hashlib.sha256(row['Email'].encode()).hexdigest() for _, row in df.iterrows() if 'Email' in row}
    if 'Email' not in df.columns:
        st.error("Excelに 'Email' 列が必要です。")
    else:
        # テスト送信
        st.markdown("### 🔧 テスト送信（1人だけ）")
        test_send_mode = st.radio("テスト送信方法", ["即時送信", "時刻指定送信"], horizontal=True)
        if test_send_mode == "時刻指定送信":
            test_send_date = st.date_input("テスト送信日", value=datetime.date.today(), key="test_date")
            test_send_time = st.time_input("テスト送信時刻", value=datetime.datetime.now().time(), key="test_time")

        if st.button("🧪 テスト送信"):
            if test_send_mode == "時刻指定送信":
                now = datetime.datetime.now()
                send_datetime = datetime.datetime.combine(test_send_date, test_send_time)
                if send_datetime < now:
                    send_datetime += datetime.timedelta(days=1)
                wait_seconds = (send_datetime - now).total_seconds()
                st.info(f"⏳ {send_datetime.strftime('%Y-%m-%d %H:%M:%S')} に送信予定（{int(wait_seconds)}秒後）")
                time.sleep(wait_seconds)
            try:
                server = smtplib.SMTP("smtp01.gmoserver.jp", 587)
                server.starttls()
                server.login(from_email, from_pass)

                first_row = df.iloc[0]
                to_email = first_row["Email"]
                name = str(first_row["Name"]) if "Name" in first_row else ""
                email_hash = hashlib.sha256(to_email.encode()).hexdigest()
                body = html_template_raw.replace("{{name}}", name).replace("{{id}}", email_hash)

                msg = MIMEMultipart()
                msg["Subject"] = subject
                msg["From"] = formataddr((from_name, from_email))
                msg["To"] = to_email
                msg.attach(MIMEText(body, "html"))

                if attachment_file is not None:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(attachment_file.read())
                    encoders.encode_base64(part)
                    part.add_header("Content-Disposition", f'attachment; filename="{attachment_file.name}"')
                    msg.attach(part)

                server.sendmail(from_email, to_email, msg.as_string())
                server.quit()
                st.success(f"✅ テスト送信成功: {to_email}")
            except Exception as e:
                st.error(f"❌ テスト送信失敗: {str(e)}")

        st.markdown("---")
        st.markdown("### 📤 一括送信 & ログ出力")
        send_mode = st.radio("送信モードを選択：", ["即時送信", "時刻指定送信"], horizontal=True)
        if send_mode == "時刻指定送信":
            send_date = st.date_input("送信日を指定：", value=datetime.date.today())
            send_time = st.time_input("送信時刻を指定：", value=datetime.datetime.now().time())

        if st.button("📩 メール一括送信開始"):
            if not st.confirm("⚠️ 本当に全員に送信してよろしいですか？ この操作は取り消せません。"):
                st.warning("送信をキャンセルしました。")
                st.stop()

            if send_mode == "時刻指定送信":
                now = datetime.datetime.now()
                send_datetime = datetime.datetime.combine(send_date, send_time)
                if send_datetime < now:
                    send_datetime += datetime.timedelta(days=1)
                wait_seconds = (send_datetime - now).total_seconds()
                st.info(f"⏳ {send_datetime.strftime('%Y-%m-%d %H:%M:%S')} に送信します（{int(wait_seconds)}秒後）")
                time.sleep(wait_seconds)

            sent_count = 0
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            log_filename = f"mail_log_{timestamp}.csv"
            with open(log_filename, mode='w', newline='', encoding='utf-8') as logfile:
                writer = csv.writer(logfile)
                writer.writerow(['Email', 'Name', 'Status', 'Error'])
                error_logs = []

                try:
                    server = smtplib.SMTP("smtp01.gmoserver.jp", 587)
                    server.starttls()
                    server.login(from_email, from_pass)

                    progress = st.progress(0)
                    total = len(df)

                    for i, (_, row) in enumerate(df.iterrows()):
                        to_email = row['Email']
                        name = str(row['Name']) if 'Name' in row else ''
                        email_hash = hashlib.sha256(to_email.encode()).hexdigest()
                        company = str(row['Company']) if 'Company' in row else ''
                        body = html_template_raw.replace('{{name}}', name).replace('{{id}}', email_hash).replace('{{company}}', company)

                        msg = MIMEMultipart()
                        msg['Subject'] = subject
                        msg['From'] = formataddr((from_name, from_email))
                        msg['To'] = to_email
                        msg.attach(MIMEText(body, 'html'))

                        if attachment_file is not None:
                            part = MIMEBase('application', 'octet-stream')
                            part.set_payload(attachment_file.read())
                            encoders.encode_base64(part)
                            part.add_header('Content-Disposition', f'attachment; filename="{attachment_file.name}"')
                            msg.attach(part)

                        try:
                            server.sendmail(from_email, to_email, msg.as_string())
                            writer.writerow([to_email, name, 'Success', ''])
                            st.success(f"✅ 送信成功: {to_email}")
                            sent_count += 1
                        except Exception as e:
                            error_message = str(e)
                            writer.writerow([to_email, name, 'Failed', error_message])
                            error_logs.append(f"{to_email} → エラー: {error_message}")
                            st.error(f"❌ 送信失敗: {to_email} - {error_message}")

                        progress.progress((i + 1) / total)

                    server.quit()

                    st.success(f"🎉 完了：{sent_count} 件のメールを送信しました")

                    st.markdown("### 📊 送信件数グラフ")
                    success_count = sent_count
                    fail_count = total - sent_count
                    fig, ax = plt.subplots()
                    ax.bar(['成功', '失敗'], [success_count, fail_count])
                    ax.set_ylabel('件数')
                    ax.set_title('送信結果')
                    st.pyplot(fig)

                    if error_logs:
                        st.markdown("### ❗ エラー送信リスト")
                        for log in error_logs:
                            st.code(log, language='text')

                    with open(log_filename, 'rb') as f:
                        st.download_button("📥 ログをダウンロード", f, file_name=log_filename)

                except Exception as e:
                    st.error(f"❌ SMTP接続エラー: {str(e)}")

        st.markdown("---")
        st.markdown("### 📥 開封ログをExcelに反映")
        open_log_file = st.file_uploader("📎 トラッキングログ（CSV形式）", type=["csv"], key="open_log")
        if open_log_file and st.button("✅ 開封ログを反映してExcel生成"):
            try:
                log_df = pd.read_csv(open_log_file, names=["timestamp", "email"])
                opened_emails = set(log_df["email"].str.lower())
                df["Opened"] = df["Email"].str.lower().apply(lambda x: "Yes" if id_map.get(x) in opened_emails else "")
                output = BytesIO()
                df.to_excel(output, index=False)
                output.seek(0)
                st.success("✅ 開封ログを反映したExcelが生成されました")
                st.download_button("📥 開封反映済みExcelをダウンロード", output, file_name="list_with_opened.xlsx")
            except Exception as e:
                st.error(f"❌ エラーが発生しました: {str(e)}")
