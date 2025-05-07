from flask import Flask, render_template, request, redirect, url_for, flash, send_file, Response, jsonify
from werkzeug.utils import secure_filename
import os
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr
from email.mime.base import MIMEBase
from email import encoders
import hashlib
import datetime
import csv
import socket
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
plt.rcParams['font.family'] = 'MS Gothic'
import io
import base64
import requests
import time

app = Flask(__name__)
app.secret_key = 'your_secret_key_here'
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/generate-test-from-excel', methods=['POST'])
def generate_test_from_excel():
    excel_file = request.files.get('excel')
    if not excel_file:
        return jsonify({"error": "Excelファイルが見つかりません"}), 400
    try:
        df = pd.read_excel(excel_file)
        if df.empty:
            return jsonify({"error": "Excelファイルにデータがありません"}), 400
        return jsonify(df.iloc[0].to_dict())
    except Exception as e:
        return jsonify({"error": f"読み込みエラー: {str(e)}"}), 500

@app.route('/download-open-log')
def download_open_log():
    log_path = os.path.join(UPLOAD_FOLDER, 'open_log.csv')
    if not os.path.exists(log_path):
        return "ログが存在しません。", 404
    return send_file(log_path, as_attachment=True)

@app.route('/export-open-summary')
def export_open_summary():
    master_log_path = os.path.join(UPLOAD_FOLDER, 'mail_log_master.csv')
    open_log_path = os.path.join(UPLOAD_FOLDER, 'open_log.csv')

    if not os.path.exists(master_log_path) or not os.path.exists(open_log_path):
        return "ログが不足しています。", 404

    df_sent = pd.read_csv(master_log_path)
    df_open = pd.read_csv(open_log_path)

    df_open['Timestamp'] = pd.to_datetime(df_open['Timestamp'], errors='coerce')
    df_open.dropna(subset=['Timestamp'], inplace=True)
    df_open['Date'] = df_open['Timestamp'].dt.date

    id_to_name = dict(zip(df_sent['ID'], df_sent['Name']))
    df_open['Name'] = df_open['ID'].map(id_to_name)

    summary = df_open.groupby(['Date', 'ID', 'Name']).size().reset_index(name='Count')
    export_path = os.path.join(UPLOAD_FOLDER, f'open_summary_{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.csv')
    summary.to_csv(export_path, index=False, encoding='utf-8-sig')
    return send_file(export_path, as_attachment=True)

@app.route('/log-view')
def view_log_table():
    log_path = os.path.join(UPLOAD_FOLDER, 'mail_log_master.csv')
    logs = []
    if os.path.exists(log_path):
        df = pd.read_csv(log_path)
        logs = df.to_dict(orient='records')
    return render_template('logs.html', logs=logs)

@app.route('/log-history')
def log_history():
    log_files = [f for f in os.listdir(UPLOAD_FOLDER) if f.startswith('mail_log_') and f.endswith('.csv')]
    log_files.sort(reverse=True)
    return render_template('log_history.html', files=log_files)

@app.route('/log-history/<filename>')
def view_log_file(filename):
    log_path = os.path.join(UPLOAD_FOLDER, filename)
    logs = []
    if os.path.exists(log_path):
        df = pd.read_csv(log_path)
        logs = df.to_dict(orient='records')
    else:
        flash("指定されたログファイルが存在しません。", 'danger')
        return redirect(url_for('log_history'))
    return render_template('logs.html', logs=logs, filename=filename)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        try:
            smtp_host = "smtp.next21.info"
            socket.gethostbyname(smtp_host)

            excel = request.files['excel']
            html = request.files['html']
            attach = request.files.get('attach')

            excel_path = os.path.join(UPLOAD_FOLDER, secure_filename(excel.filename))
            html_path = os.path.join(UPLOAD_FOLDER, secure_filename(html.filename))
            attach_path = os.path.join(UPLOAD_FOLDER, secure_filename(attach.filename)) if attach else None

            excel.save(excel_path)
            html.save(html_path)
            if attach:
                attach.save(attach_path)

            from_name = request.form['from_name']
            from_email = request.form['from_email']
            from_pass = request.form['from_pass']
            subject = request.form['subject']

            with open(html_path, 'r', encoding='utf-8') as f:
                html_template = f.read()

            df = pd.read_excel(excel_path)
            log = []

            server = smtplib.SMTP(smtp_host, 587)
            server.starttls()
            server.login(from_email, from_pass)

            for _, row in df.iterrows():
                to_email = row['Email']
                name = str(row['Name']) if 'Name' in row else ''
                company = str(row['Company']) if 'Company' in row else ''
                email_hash = hashlib.sha256(to_email.encode()).hexdigest()

                body = html_template.replace("{{name}}", name).replace("{{id}}", email_hash).replace("{{company}}", company)

                msg = MIMEMultipart()
                msg['Subject'] = subject
                msg['From'] = formataddr((from_name, from_email))
                msg['To'] = to_email
                msg['Message-ID'] = f"<{hashlib.sha256((to_email + str(datetime.datetime.utcnow())).encode()).hexdigest()}@next21.info>"
                msg.add_header("List-Unsubscribe", f"<mailto:{from_email}>")
                msg.attach(MIMEText(body, 'html'))

                if attach_path:
                    with open(attach_path, 'rb') as a:
                        part = MIMEBase('application', 'octet-stream')
                        part.set_payload(a.read())
                        encoders.encode_base64(part)
                        part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(attach_path)}"')
                        msg.attach(part)

                try:
                    server.sendmail(from_email, to_email, msg.as_string())
                    log.append((to_email, name, email_hash, subject, 'Success', ''))
                except Exception as e:
                    log.append((to_email, name, email_hash, subject, 'Failed', str(e)))

                time.sleep(1)

            server.quit()

            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            log_path = os.path.join(UPLOAD_FOLDER, f'mail_log_{timestamp}.csv')
            with open(log_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(['Email', 'Name', 'ID', 'Subject', 'Status', 'Error'])
                writer.writerows(log)

            master_path = os.path.join(UPLOAD_FOLDER, 'mail_log_master.csv')
            master_exists = os.path.exists(master_path)
            with open(master_path, 'a', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                if not master_exists:
                    writer.writerow(['Email', 'Name', 'ID', 'Subject', 'Status', 'Error'])
                writer.writerows(log)

            flash(f"送信完了：{len(log)}件のログを作成しました。", 'success')
            return send_file(log_path, as_attachment=True)

        except Exception as e:
            flash(f"送信エラー：{str(e)}", 'danger')

    return render_template('index.html')

@app.route('/open.gif')
def track_open():
    user_id = request.args.get('id')
    timestamp = datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
    ip = request.remote_addr

    log_path = os.path.join(UPLOAD_FOLDER, 'open_log.csv')
    file_exists = os.path.exists(log_path)
    with open(log_path, 'a', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(['ID', 'Timestamp', 'IP'])
        writer.writerow([user_id, timestamp, ip])

    return Response(status=204)

@app.route('/analyze', methods=['GET'])
def analyze():
    try:
        render_log_url = "https://flask-mail-app-opcz.onrender.com/download-open-log"
        response = requests.get(render_log_url)
        if response.status_code == 200:
            backup_path = os.path.join(UPLOAD_FOLDER, f"open_log_backup_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
            target_path = os.path.join(UPLOAD_FOLDER, 'open_log.csv')
            if os.path.exists(target_path):
                os.rename(target_path, backup_path)
            with open(target_path, 'wb') as f:
                f.write(response.content)
    except Exception as e:
        flash(f"Renderログ取得失敗: {str(e)}", "danger")

    master_log_path = os.path.join(UPLOAD_FOLDER, 'mail_log_master.csv')
    open_log_path = os.path.join(UPLOAD_FOLDER, 'open_log.csv')

    if not os.path.exists(master_log_path) or not os.path.exists(open_log_path):
        flash("ログファイルが不足しています。", 'danger')
        return redirect('/')

    df_sent = pd.read_csv(master_log_path)

    # 過去のバックアップログも取得
    backup_logs = sorted(
        [f for f in os.listdir(UPLOAD_FOLDER) if f.startswith('open_log_backup_') and f.endswith('.csv')],
        reverse=True
    )
    options = [('open_log.csv', '最新ログ')]
    for file in backup_logs:
        label = file.replace('open_log_backup_', '').replace('.csv', '')
        options.append((file, label))

    selected_file = request.args.get('log_file', 'open_log.csv')
    selected_log_path = os.path.join(UPLOAD_FOLDER, selected_file)

    if not os.path.exists(selected_log_path):
        flash(f"選択されたログファイルが見つかりません: {selected_file}", 'danger')
        return redirect('/')

    df_open = pd.read_csv(selected_log_path)

    if df_open.empty or df_open.shape[0] == 0 or 'Timestamp' not in df_open.columns:
        flash("開封ログがまだ記録されていません。", "warning")
        return render_template("analyze.html", graph="", opened_by={}, impression_rate=0,
                               total_sent=df_sent['ID'].nunique(), unique_opens=0,
                               pie_chart="", dm_summary=[], open_log=[], last_updated="データなし",
                               log_options=options, selected_file=selected_file)

    df_open['Timestamp'] = pd.to_datetime(df_open['Timestamp'], errors='coerce')
    df_open.dropna(subset=['Timestamp'], inplace=True)
    if df_open.empty:
        flash("有効な開封ログがありません。", "warning")
        return render_template("analyze.html", graph="", opened_by={}, impression_rate=0,
                               total_sent=df_sent['ID'].nunique(), unique_opens=0,
                               pie_chart="", dm_summary=[], open_log=[], last_updated="データなし",
                               log_options=options, selected_file=selected_file)

    df_open['Date'] = df_open['Timestamp'].dt.date
    daily_counts = df_open.groupby('Date').size()

    id_to_name = dict(zip(df_sent['ID'], df_sent['Name']))
    df_open['Name'] = df_open['ID'].map(id_to_name).fillna('不明')
    opened_by = df_open['Name'].value_counts()

    total_sent = df_sent['ID'].nunique()
    unique_opens = df_open['ID'].nunique()
    impression_rate = unique_opens / total_sent * 100 if total_sent else 0

    dm_summary = []
    if 'Subject' in df_sent.columns:
        subjects = df_sent['Subject'].unique()
        for subj in subjects:
            df_sent_sub = df_sent[df_sent['Subject'] == subj]
            sent_ids = df_sent_sub['ID'].unique()
            open_ids = df_open[df_open['ID'].isin(sent_ids)]['ID'].nunique()
            total_ids = len(sent_ids)
            rate = round(open_ids / total_ids * 100, 2) if total_ids else 0
            dm_summary.append({'Subject': subj, 'Opened': open_ids, 'Total': total_ids, 'Rate': rate})

    graph_base64 = ""
    if not daily_counts.empty:
        plt.figure(figsize=(10, 5))
        daily_counts.plot(kind='bar', title='📊 日別の開封数')
        plt.xlabel('日付')
        plt.ylabel('開封数')
        buf_bar = io.BytesIO()
        plt.tight_layout()
        plt.savefig(buf_bar, format='png')
        buf_bar.seek(0)
        graph_base64 = base64.b64encode(buf_bar.read()).decode('utf-8')
        plt.close()

    labels = ['開封', '未開封']
    sizes = [unique_opens, max(0, total_sent - unique_opens)]
    pie_base64 = ""
    if sum(sizes) > 0:
        plt.figure(figsize=(5, 5))
        plt.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90, colors=['#4CAF50', '#FFC107'])
        plt.axis('equal')
        buf_pie = io.BytesIO()
        plt.savefig(buf_pie, format='png')
        buf_pie.seek(0)
        pie_base64 = base64.b64encode(buf_pie.read()).decode('utf-8')
        plt.close()

    last_updated = "データなし"
    if not df_open['Timestamp'].isna().all():
        last_updated = df_open['Timestamp'].max().strftime('%Y-%m-%d %H:%M:%S')

    return render_template('analyze.html', graph=graph_base64,
                           opened_by=opened_by.to_dict(),
                           impression_rate=round(impression_rate, 2),
                           total_sent=total_sent,
                           unique_opens=unique_opens,
                           pie_chart=pie_base64,
                           dm_summary=dm_summary,
                           open_log=df_open.to_dict(orient='records'),
                           last_updated=last_updated,
                           log_options=options,
                           selected_file=selected_file)

if __name__ == '__main__':
    app.run(debug=True)
