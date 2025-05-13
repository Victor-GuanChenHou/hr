from flask import Flask, render_template, request, jsonify, send_file, session, redirect, url_for, flash
import pandas as pd
import base64
import glob
import os
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from io import BytesIO
import shutil
app = Flask(__name__)
app.secret_key = 'your_secret_key_here'

UPLOAD_FOLDER = 'uploads'
SIGNATURE_FOLDER = 'static/signatures'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['SIGNATURE_FOLDER'] = SIGNATURE_FOLDER

# 建立資料夾
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(SIGNATURE_FOLDER, exist_ok=True)

USER_DB = {
    'a03003': 'a03003',
    'a03004': 'a03004',
    'admin': 'admin'
}

@app.route('/icon')
def icon():
    return send_file('./templates/kingza.ico')

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        if username in USER_DB and USER_DB[username] == password:
            session['username'] = username
            return redirect(url_for('index'))
        else:
            return render_template('login.html', error='登入失敗，帳號或密碼錯誤')
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.pop('username', None)
    return redirect(url_for('login'))

@app.route('/')
def index():
    if 'username' not in session:
        return redirect(url_for('login'))

    username = session['username']

    if username == 'admin':
        return render_template('admin.html')

    if username == 'a03003':
        AffiliatedUnit = '杏子台北車站微風店'
    else:
        AffiliatedUnit = ''

    # 讀取最新 Excel
    files = glob.glob(os.path.join(UPLOAD_FOLDER, '*.xlsx'))
    if not files:
        raise FileNotFoundError('找不到任何 Excel 檔案')

    latest_file = max(files, key=os.path.getmtime)
    df = pd.read_excel(latest_file)

    # 過濾國定假日
    filtered_df = df[(df['單位名稱'] == AffiliatedUnit) & (df['班別'] == '國定假日')]
    filtered_df = filtered_df.reset_index(drop=True)

    display_data = []
    for idx, row in filtered_df.iterrows():
        item = row[['單位名稱', '員工編號', '員工姓名', '身份別', '日期', '班別']].to_dict()
        signature_file = os.path.join(SIGNATURE_FOLDER, username, f'row_{idx}.png')
        if os.path.exists(signature_file):
            item['signature'] = f'/static/signatures/{username}/row_{idx}.png'
        else:
            item['signature'] = ''  # 沒有簽名
        display_data.append(item)

    return render_template('index.html', tables=display_data)

@app.route('/upload_original_data', methods=['POST'])
def upload_original_data():
    if 'file' not in request.files:
        return jsonify({"error": "沒有檔案部分！"}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "沒有選擇檔案！"}), 400

    file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(file_path)
    return redirect(url_for('login'))


@app.route('/sign', methods=['POST'])
def sign():
    if 'username' not in session:
        return jsonify({'status': 'fail', 'message': '未登入'}), 401

    username = session['username']
    data = request.json
    row_index = data['row']
    img_data = data['signature'].split(',')[1]  # 移除 base64 開頭
    img_bytes = base64.b64decode(img_data)

    # 建立使用者資料夾
    user_folder = os.path.join(app.config['SIGNATURE_FOLDER'], username)
    if not os.path.exists(user_folder):
        os.makedirs(user_folder)

    filename = f'row_{row_index}.png'
    file_path = os.path.join(user_folder, filename)

    with open(file_path, 'wb') as f:
        f.write(img_bytes)

    # 回傳圖片路徑給前端（用於顯示）
    return jsonify({'status': 'success', 'file': f'/static/signatures/{username}/{filename}'})

@app.route('/saveall', methods=['POST'])
def saveall():
    username = session.get('username')
    if username == 'a03003':
        AffiliatedUnit = '杏子台北車站微風店'
    else:
        AffiliatedUnit = ''

    files = glob.glob(os.path.join(UPLOAD_FOLDER, '*.xlsx'))
    if not files:
        raise FileNotFoundError('找不到任何 Excel 檔案')

    latest_file = max(files, key=os.path.getmtime)
    df = pd.read_excel(latest_file)

    filtered_df = df[(df['單位名稱'] == AffiliatedUnit) & (df['班別'] == '國定假日')]
    filtered_df = filtered_df.reset_index()

    # 建立 signature_path 欄位
    

    for _, row in filtered_df.iterrows():
        index = row['index']
        sign_filename = f'row_{index}.png'
        sign_path = os.path.join(SIGNATURE_FOLDER,username, sign_filename)
        if os.path.exists(sign_path):
            df.at[index, 'signature_path'] = f'/static/signatures/{username}/{sign_filename}'

    df.drop(columns=['index'], inplace=True)
    df.to_excel(latest_file, index=False)

    
    return jsonify({'status': 'success', 'message': '所有簽名已儲存！'})

@app.route('/download_latest_excel', methods=['GET'])
def download_latest_excel():
    username = 'a03003'  # 或從 session 中取
    if username == 'a03003':
        AffiliatedUnit = '杏子台北車站微風店'
    else:
        AffiliatedUnit = ''

    files = glob.glob(os.path.join(UPLOAD_FOLDER, '*.xlsx'))
    if not files:
        raise FileNotFoundError('找不到任何 Excel 檔案')
    latest_file = max(files, key=os.path.getmtime)

    df = pd.read_excel(latest_file)
    filtered_df = df[(df['單位名稱'] == AffiliatedUnit) & (df['班別'] == '國定假日')].reset_index(drop=True)

    # 建立新 Excel
    wb = Workbook()
    ws = wb.active
    ws.title = '簽名表'

    # 加入表頭
    headers = ['單位名稱', '員工編號', '員工姓名', '身份別', '日期', '班別', '簽名']
    ws.append(headers)

    # 加入每一列資料 + 插入簽名圖片
    for i, row in filtered_df.iterrows():
        row_data = [row['單位名稱'], row['員工編號'], row['員工姓名'],
                    row['身份別'], row['日期'], row['班別'], '']  # 最後一格預留給圖片
        ws.append(row_data)

        # 插入圖片
        img_path = os.path.join(SIGNATURE_FOLDER, username, f'row_{i}.png')
        if os.path.exists(img_path):
            img = ExcelImage(img_path)
            img.width, img.height = 100, 50
            ws.add_image(img, f'G{i+2}')  # G 欄，從第 2 行開始
            ws.row_dimensions[i+2].height = 40

    # 儲存 Excel
    if not os.path.exists('temp'):
        os.makedirs('temp')
    output_path = os.path.join('temp', f'{username}_signed_filtered.xlsx')
    wb.save(output_path)

    return send_file(output_path, as_attachment=True)
@app.route('/deletimage',methods=['POST'])
def deletimage():
    if 'username' not in session:
        return redirect(url_for('login'))

    #username = session['username']
    username='a03003'
    user_signature_folder = os.path.join(app.config['SIGNATURE_FOLDER'],username)

    if os.path.exists(user_signature_folder):
        for file in os.listdir(user_signature_folder):
            file_path = os.path.join(user_signature_folder, file)
            if os.path.isfile(file_path):
                os.remove(file_path)

    
    return jsonify({'status': 'success', 'message': '簽名檔已全部清除'})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=327)
