import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd
import numpy as np
import glob
from dotenv import load_dotenv
import os
import math
import json
import sub
from datetime import datetime
ENV = './.env' 
load_dotenv(dotenv_path=ENV)
UPLOAD_FOLDER = 'uploads/upload_month'
SIGNATURE_FOLDER = 'static/year_signatures'


def find_unsign(SIGNATURE_FOLDER):
    #'橋村桃園大江店', 'A14926', '陳妮潔', '門市正職人員', 20251113, '國定假日', 'angus.chang@kingza.com.tw']
    data=sub.docxuser_manager_mail()
    df = pd.DataFrame(data)
    df['單位主管信箱'] = df['單位主管信箱'].replace('', np.nan)
    #{'單位名稱': '雞三和漢神巨蛋店', '員工編號': 'A15562', '員工姓名': '李芳', '身份別': '門市正職人員', '單位主管': 'A12655', '單位主管信箱': ''}
    
  #  filtered_df = df[['單位名稱', '員工編號', '員工姓名', '身份別', '單位主管信箱']].values.tolist()
    
    EMID = df['員工編號'].unique().tolist()

    # 建立未簽名名單 
    unsigned_data = []
    for emp_id in EMID:
        emp_rows = df[df['員工編號'] == emp_id].reset_index(drop=True)
        for i, row in emp_rows.iterrows():
            img_path = os.path.join(SIGNATURE_FOLDER, f'{emp_id}.png')
            if not os.path.exists(img_path):
                unsigned_data.append([
                    row['單位名稱'], row['員工編號'], row['員工姓名'],
                    row['身份別'],row['單位主管信箱']
                ])
    return unsigned_data
def ChangeToHTML(unsigned_data):
    # 轉成 HTML 表格
    headers = ['單位名稱', '員工編號', '員工姓名', '身份別']
    html_table = '<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse;">'
    # 表頭
    html_table += '<tr>' + ''.join([f'<th>{h}</th>' for h in headers]) + '</tr>'
    # 資料
    for row in unsigned_data:
        html_table += '<tr>' + ''.join([f'<td>{cell}</td>' for cell in row]) + '</tr>'
    html_table += '</table>'

    # 郵件主體
    
    body_html = f"""
    <html>
    <head>
    <meta charset="utf-8">
    </head>
    <body>
    <p>以下是尚未簽核的名單，再請盡快通知並完成簽核，謝謝</p>
    <p><a href="http://hrsignin.kingza.com.tw/login">🔗 點我前往簽核系統</a></p>
    <p>帳號 : 員工工號</p>
    <p>密碼 : 身分證後九碼</p>
    {html_table}
    </body>
    </html>
    """
    return body_html
def Send_EMAIL(unsigned_data,email):
    # 郵件內容設定
    sender_email = os.getenv('MAIL')
    receiver_email=email
    password = os.getenv('MAIL_PW')
    ad_year = datetime.now().year
    # 2. 換算成民國年
    roc_year = ad_year - 1911
    subject = f"{roc_year}年度國定假日調移同意書未簽核名單"
    
    body_html=ChangeToHTML(unsigned_data)


    # 建立郵件物件
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject
    # 郵件主體
    message.attach(MIMEText(body_html, "html"))



    try:
        # 建立與 Gmail SMTP 伺服器的連線 (使用 SSL)
        with smtplib.SMTP_SSL("mail.kingza.com.tw", 465) as server:
            if not (isinstance(email, float) and math.isnan(email)):
                server.login(sender_email, password)
                server.sendmail(sender_email, receiver_email, message.as_string())
                print("郵件寄送成功！")

    except Exception as e:
        print(f"發生錯誤：{e}")

with open('email.json', 'r', encoding='utf-8') as f:
        store_email_data = json.load(f)


unsigned_data=find_unsign(SIGNATURE_FOLDER)

chief_groups={}
store_groups={}
EMAIL=[]
storeEMAIL=[]
for row in unsigned_data:
    store_email = row[0]
    # 如果門市還沒出現，就建立一個新的 list
    for item in store_email_data:
        if item['name'] == row[0]:
            store_email=item['email']
            break
        else : 
            store_email='error'
    if store_email not in store_groups:
        storeEMAIL.append(store_email)
        store_groups[store_email] = []
    # 把該筆 row 加進去
    store_groups[store_email].append(row)

# ################發給區經理###########################    
    chief_email = row[-1]
    # 如果主管還沒出現，就建立一個新的 list
    if chief_email not in chief_groups:
        EMAIL.append(chief_email)
        chief_groups[chief_email] = []
    # 把該筆 row 加進去
    chief_groups[chief_email].append(row)
for i in range(len(storeEMAIL)): #根據門市發信
    if storeEMAIL[i] !='error':
        rows_for_send = [r[:-1] for r in store_groups[storeEMAIL[i]]]
        print(rows_for_send)
        Send_EMAIL(rows_for_send,storeEMAIL[i])
for i in range(len(EMAIL)): #根據部門主管發信
    if not pd.isna(EMAIL[i]):
        rows_for_send = [r[:-1] for r in chief_groups[EMAIL[i]]]
        Send_EMAIL(rows_for_send,EMAIL[i])
rows_for_send_all = [r[:-1] for r in unsigned_data]
Send_EMAIL(rows_for_send_all,'hr@kingza.com.tw')#HR信箱
