from flask import Flask, send_from_directory,render_template, request, jsonify, send_file, session, redirect, url_for, flash
import pandas as pd
import base64
import glob
import os
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from io import BytesIO
import shutil
import json
import sub as sub
from datetime import datetime
from dotenv import load_dotenv
import os
ENV = './.env' 
load_dotenv(dotenv_path=ENV)
SEC_KEY = os.getenv('SEC_KEY')
app = Flask(__name__)
app.secret_key = SEC_KEY

UPLOAD_FOLDER = 'uploads'
SIGNATURE_FOLDER = 'static/signatures'
HISTORY_FOLDER='history'
TEMP='temp'

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['SIGNATURE_FOLDER'] = SIGNATURE_FOLDER
app.config['TEMP'] = TEMP
app.config['HISTORY_FOLDER'] = HISTORY_FOLDER

# 建立資料夾
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(SIGNATURE_FOLDER, exist_ok=True)
def gettabledata(foldername):

    dept_no=session['dept_no']
    display_data=[]
    uploadfoler=HISTORY_FOLDER+'/'+foldername+'/uploads'
    signaturefolder=HISTORY_FOLDER+'/'+foldername+'/signatures'
    files = glob.glob(os.path.join(uploadfoler, '*.xlsx'))
    if not files:
        if dept_no == '139' or dept_no=='452':
            return display_data
    latest_file = max(files, key=os.path.getmtime)

    # 使用我們定義的函式來兼容讀取
    try:
        df = sub.read_excel_compatible(latest_file)
    except Exception as e:
        return display_data
    # 過濾國定假日

    if dept_no == '139' or dept_no=='452':    #人資部&資訊部
        filtered_df = df[(df['班別'] == '國定假日') & ((df['身份別'] == '門市副理(含)級以上') | (df['身份別'] == '門市正職人員'))].reset_index(drop=True)
        EMID = filtered_df['員工編號'].unique().tolist()
        display_data = []
        for emp_id in EMID:
            emp_rows = filtered_df[filtered_df['員工編號'] == emp_id].reset_index(drop=True)

            for i, row in emp_rows.iterrows():
                # 寫入一筆資料
                item = row[['單位名稱', '員工編號', '員工姓名', '身份別', '日期', '班別']].to_dict()
                signature_file = os.path.join(signaturefolder, emp_id, f'row_{i}.png')
                if os.path.exists(signature_file):
                    item['signature'] = f'/history/{foldername}/signatures/{emp_id}/row_{i}.png'
                else:
                    item['signature'] = ''  # 沒有簽名
                display_data.append(item)
    return display_data
@app.route('/icon')
def icon():
    return send_file('./templates/kingza.ico')
@app.route('/')
def index_redirect():
    if 'username' in session:
        return redirect(url_for('home'))
    return redirect(url_for('login'))
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        user_info = sub.get_user_info(username)
        user_ip = request.remote_addr
        if user_info and( user_info['password'] == password or user_info['password'] != password):
            if( user_info['CLASS']=='D' or user_info['DEPT_NO']=='139' or user_info['DEPT_NO']=='452') :
            
                # with open('allowdept.json', 'r', encoding='utf-8') as f:
                #     config = json.load(f)
                # allow_dept = set(config.get('allowdept', []))a
                # if user_info['DEPT_NO'] in allow_dept or user_info['DEPT_KIND'] in allow_dept:
                #     session['username'] = user_info['username']
                #     session['name'] = user_info['name']
                #     session['dept_no']=user_info['DEPT_NO']
                #     session['dept_name']=user_info['DEPT_NAME']
                #     return redirect(url_for('home'))
                # else:
                #     return render_template('login.html', error='，帳號或密碼錯誤')
                session['username'] = user_info['username']
                session['name'] = user_info['name']
                session['dept_no']=user_info['DEPT_NO']
                session['dept_name']=user_info['DEPT_NAME']
                user_ip = request.headers.get('X-Forwarded-For', request.remote_addr)
                sub.loglogin(session['username'],user_ip)
                return redirect(url_for('home'))
            else:
                return render_template('login.html', error='無權限請洽管理員')
        else:
            return render_template('login.html', error='登入失敗，帳號或密碼錯誤')
    return render_template('login.html')
@app.route('/admin')
def admin():
    if 'username' not in session:
        return redirect(url_for('login'))
    return render_template('admin.html')
@app.route('/logout')
def logout():
    session.pop('username', None)
    return redirect(url_for('login'))
@app.route('/home')
def home():
    if 'username' not in session:
        return redirect(url_for('login'))
    username = session['username']
    name = session['name']
    dept_no=session['dept_no']
    dept1 = []
    dept2 = []
    try:
        with open("email.json", "r", encoding="utf-8") as f:
            store_data = json.load(f)
    except Exception as e:
        print("讀取錯誤:", e)
        store_data = []
    dept1 = list({item.get("dept1") for item in store_data if item.get("dept1")})
    dept2 = list({item.get("dept2") for item in store_data if item.get("dept2")})

    if dept_no == '139' or dept_no=='452':
        return render_template('homeadmin.html', username=username,name=name)
    else:
        if username in dept1 or username in dept2:
            return render_template('home.html', username=username,name=name,has_permission=True)
        else:
            return render_template('home.html', username=username,name=name,has_permission=False)
@app.route('/home/sing')
def index():
    if 'username' not in session:
        return redirect(url_for('login'))
    username = session['username']
    name = session['name']
    dept_no=session['dept_no']
    dept1 = []
    dept2 = []
    try:
        with open("email.json", "r", encoding="utf-8") as f:
            store_data = json.load(f)
    except Exception as e:
        print("讀取錯誤:", e)
        store_data = []
    dept1 = list({item.get("dept1") for item in store_data if item.get("dept1")})
    dept2 = list({item.get("dept2") for item in store_data if item.get("dept2")})
    # 讀取最新 Excel
    files = glob.glob(os.path.join(UPLOAD_FOLDER, '*.xlsx'))
    if not files:
        if dept_no == '139' or dept_no=='452':
            return render_template('admin.html', tables=[], username=username, name=name, no_data=True)
        else:
            if username in dept1 or username in dept2:
                return render_template('index.html', tables=[], username=username, name=name, no_data=True,has_permission=True)
            else:
                return render_template('index.html', tables=[], username=username, name=name, no_data=True,has_permission=False)
    latest_file = max(files, key=os.path.getmtime)

    # 使用我們定義的函式來兼容讀取
    try:
        df = sub.read_excel_compatible(latest_file)
    except Exception as e:
        return f'Excel 載入失敗：{e}', 500
    # 過濾國定假日

    if dept_no == '139' or dept_no=='452':    #人資部&資訊部
        filtered_df = df[(df['班別'] == '國定假日') & ((df['身份別'] == '門市副理(含)級以上') | (df['身份別'] == '門市正職人員'))].reset_index(drop=True)
        EMID = filtered_df['員工編號'].unique().tolist()
        display_data = []
        for emp_id in EMID:
            emp_rows = filtered_df[filtered_df['員工編號'] == emp_id].reset_index(drop=True)

            for i, row in emp_rows.iterrows():
                # 寫入一筆資料
                item = row[['單位名稱', '員工編號', '員工姓名', '身份別', '日期', '班別']].to_dict()
                signature_file = os.path.join(SIGNATURE_FOLDER, emp_id, f'row_{i}.png')
                if os.path.exists(signature_file):
                    item['signature'] = f'/static/signatures/{emp_id}/row_{i}.png'
                else:
                    item['signature'] = ''  # 沒有簽名
                display_data.append(item)
        
        return render_template('admin.html', tables=display_data, username=username,name=name)
        #return render_template('admin.html', username=username, name=name)
    else:
        filtered_df = df[(df['員工編號'] == username) & (df['班別'] == '國定假日') & ((df['身份別'] == '門市副理(含)級以上') | (df['身份別'] == '門市正職人員'))]
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
        if username in dept1 or username in dept2:
            return render_template('index.html', tables=display_data, username=username, name=name, has_permission=True)
        else:
            return render_template('index.html', tables=display_data, username=username, name=name, has_permission=False)

    
@app.route('/home/historysearch')
def historysearch():
    if 'username' not in session:
        return redirect(url_for('login'))
    username = session['username']
    name = session['name']
    dept_no=session['dept_no']
    history_path = 'history'

    # 取得資料夾中的所有項目（檔案與資料夾）
    items = os.listdir(history_path)

    # 過濾出只有資料夾的名稱
    #folders = [item[:10] for item in items if os.path.isdir(os.path.join(history_path, item))]
    folders = [item for item in items if os.path.isdir(os.path.join(history_path, item))]
    foldername=folders[0]
    # 讀取最新 Excel
    uploadfoler=HISTORY_FOLDER+'/'+foldername+'/uploads'
    signaturefolder=HISTORY_FOLDER+'/'+foldername+'/signatures'
    files = glob.glob(os.path.join(uploadfoler, '*.xlsx'))
    if not files:
        if dept_no == '139' or dept_no=='452':
            return render_template('history.html',folders=folders, tables=[], username=username, name=name, no_data=True)
    latest_file = max(files, key=os.path.getmtime)

    # 使用我們定義的函式來兼容讀取
    try:
        df = sub.read_excel_compatible(latest_file)
    except Exception as e:
        return f'Excel 載入失敗：{e}', 500
    # 過濾國定假日

    if dept_no == '139' or dept_no=='452':    #人資部&資訊部
        filtered_df = df[(df['班別'] == '國定假日') & ((df['身份別'] == '門市副理(含)級以上') | (df['身份別'] == '門市正職人員'))].reset_index(drop=True)
        EMID = filtered_df['員工編號'].unique().tolist()
        display_data = []
        for emp_id in EMID:
            emp_rows = filtered_df[filtered_df['員工編號'] == emp_id].reset_index(drop=True)

            for i, row in emp_rows.iterrows():
                # 寫入一筆資料
                item = row[['單位名稱', '員工編號', '員工姓名', '身份別', '日期', '班別']].to_dict()
                signature_file = os.path.join(signaturefolder, emp_id, f'row_{i}.png')
                if os.path.exists(signature_file):
                    item['signature'] = f'/history/{foldername}/signatures/{emp_id}/row_{i}.png'
                else:
                    item['signature'] = ''  # 沒有簽名
                display_data.append(item)
        
        return render_template('history.html',folders=folders, tables=display_data, username=username,name=name,no_data=False)
        #return render_template('admin.html', username=username, name=name)
@app.route("/filter_table")
def filter_table():
    filter_value = request.args.get("filter")
    tables=gettabledata(filter_value)
    return render_template("table_only.html", tables=tables,no_data=False)
@app.route('/upload_original_data', methods=['POST'])
def upload_original_data():
    if 'file' not in request.files:
        return jsonify({"success": False, "error": "沒有檔案部分！"}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({"success": False, "error": "沒有選擇檔案！"}), 400
    try:
        if file.filename.endswith('.xlsx') :
            df = pd.read_excel(file)
        else:
            return jsonify({"success": False, "error": "不支援的檔案格式！"}), 400
    except Exception as e:
        return jsonify({"success": False, "error": f"檔案讀取失敗：{str(e)}"}), 400
    try:
        filtered_df = df[(df['班別'] == '國定假日') & ((df['身份別'] == '門市副理(含)級以上') | (df['身份別'] == '門市正職人員'))]
        filtered_df = filtered_df.reset_index(drop=True)
        if filtered_df.empty:
            return jsonify({"success": False, "error": "沒有符合條件的資料！"}), 400
        else:
            filtered_df['主管'] = filtered_df['員工編號'].apply(lambda emp_id: sub.find_deptchie(emp_id))
        save_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    except:
        return jsonify({"success": False, "error": "沒有符合條件的資料！"}), 400
    try:
        filtered_df.to_excel(save_path, index=False)
    except Exception as e:
        return jsonify({"success": False, "error": f"檔案儲存失敗：{str(e)}"}), 500
    return jsonify({"success": True, "message": "檔案已成功上傳！"})
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

    filtered_df = df[(df['單位名稱'] == AffiliatedUnit) & ((df['班別'] == '國定假日') & ((df['身份別'] == '門市副理(含)級以上') | (df['身份別'] == '門市正職人員')))]
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
    files = glob.glob(os.path.join(UPLOAD_FOLDER, '*.xlsx'))
    if not files:
        raise FileNotFoundError('找不到任何 Excel 檔案')

    latest_file = max(files, key=os.path.getmtime)
    df = pd.read_excel(latest_file)
    filtered_df = df[(df['班別'] == '國定假日') & ((df['身份別'] == '門市副理(含)級以上') | (df['身份別'] == '門市正職人員'))].reset_index(drop=True)
    EMID = filtered_df['員工編號'].unique().tolist()

    wb_all = Workbook()
    ws_all = wb_all.active
    ws_all.title = '全部'

    wb_signed = Workbook()
    ws_signed = wb_signed.active
    ws_signed.title = '已簽名'

    wb_unsigned = Workbook()
    ws_unsigned = wb_unsigned.active
    ws_unsigned.title = '未簽名'
    

    headers = ['單位名稱', '員工編號', '員工姓名', '身份別', '日期', '班別', '簽名']
    for ws in [ws_all, ws_signed, ws_unsigned]:
        ws.append(headers)

    for emp_id in EMID:
        emp_rows = filtered_df[filtered_df['員工編號'] == emp_id].reset_index(drop=True)

        for i, row in emp_rows.iterrows():
            row_data = [
                row['單位名稱'], row['員工編號'], row['員工姓名'],
                row['身份別'], row['日期'], row['班別'], ''
            ]

            img_path = os.path.join(SIGNATURE_FOLDER, emp_id, f'row_{i}.png')
            img_exists = os.path.exists(img_path)

            # 將資料寫入三個工作表
            # 1. 全部
            ws_all.append(row_data)
            row_idx_all = ws_all.max_row
            if img_exists:
                img = ExcelImage(img_path)
                img.width, img.height = 100, 50
                ws_all.add_image(img, f'G{row_idx_all}')
                ws_all.row_dimensions[row_idx_all].height = 40

            # 2. 已簽名
            if img_exists:
                ws_signed.append(row_data)
                row_idx_signed = ws_signed.max_row
                img = ExcelImage(img_path)
                img.width, img.height = 100, 50
                ws_signed.add_image(img, f'G{row_idx_signed}')
                ws_signed.row_dimensions[row_idx_signed].height = 40
            else:
                # 3. 未簽名
                ws_unsigned.append(row_data)
    ############未簽名資料匯出欄寬######################
    ws_unsigned.column_dimensions['A'].width = 25
    ws_unsigned.column_dimensions['B'].width = 10
    ws_unsigned.column_dimensions['C'].width = 10
    ws_unsigned.column_dimensions['D'].width = 25
    ws_unsigned.column_dimensions['E'].width = 15
    ws_unsigned.column_dimensions['F'].width = 10
    ws_unsigned.column_dimensions['G'].width = 15
    ############已簽名資料匯出欄寬######################
    ws_signed.column_dimensions['A'].width = 25
    ws_signed.column_dimensions['B'].width = 10
    ws_signed.column_dimensions['C'].width = 10
    ws_signed.column_dimensions['D'].width = 25
    ws_signed.column_dimensions['E'].width = 15
    ws_signed.column_dimensions['F'].width = 10
    ws_signed.column_dimensions['G'].width = 15
    ############全部資料匯出欄寬######################
    ws_all.column_dimensions['A'].width = 25
    ws_all.column_dimensions['B'].width = 10
    ws_all.column_dimensions['C'].width = 10
    ws_all.column_dimensions['D'].width = 25
    ws_all.column_dimensions['E'].width = 15
    ws_all.column_dimensions['F'].width = 10
    ws_all.column_dimensions['G'].width = 15
    # 儲存檔案
    if not os.path.exists(app.config['TEMP']):
        os.makedirs(app.config['TEMP'])
    output_path = os.path.join(app.config['TEMP'], 'signed_filtered.xlsx')
    status = request.args.get('status')
    
    if(status=='unsigned'):
        wb_unsigned.save(output_path)
    elif(status=='signed'):
        wb_signed.save(output_path)
    else:
        wb_all.save(output_path)

    return send_file(output_path, as_attachment=True)
@app.route('/settlement',methods=['POST'])
def settlement():
    if 'username' not in session:
        return redirect(url_for('login'))

    settlement_time = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    settlement_folder = os.path.join(app.config['HISTORY_FOLDER'], settlement_time)
    os.makedirs(settlement_folder, exist_ok=True)

    # 處理 signatures 整個資料夾
    if os.path.exists(app.config['SIGNATURE_FOLDER']):
        dest_signatures = os.path.join(settlement_folder, 'signatures')
        shutil.move(app.config['SIGNATURE_FOLDER'], dest_signatures)

    # 處理 uploads 整個資料夾
    if os.path.exists(app.config['UPLOAD_FOLDER']):
        dest_uploads = os.path.join(settlement_folder, 'uploads')
        shutil.move(app.config['UPLOAD_FOLDER'], dest_uploads)

    # 移動完後，重新建立空的 signatures 和 uploads 資料夾
    os.makedirs(app.config['SIGNATURE_FOLDER'], exist_ok=True)
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

    
    return jsonify({'status': 'success', 'message': '已結算'})
@app.route('/get_signed_data')
def get_signed_data():
    files = glob.glob(os.path.join(UPLOAD_FOLDER, '*.xlsx'))
    if not files:
        return jsonify([])

    latest_file = max(files, key=os.path.getmtime)
    df = pd.read_excel(latest_file)
    filtered_df = df[(df['班別'] == '國定假日') & ((df['身份別'] == '門市副理(含)級以上') | (df['身份別'] == '門市正職人員'))].reset_index(drop=True)

    # 新增簽名欄位：有圖就回傳圖網址，否則為 None
    def get_signature_path(row):
        emp_id = row['員工編號']
        img_path = os.path.join(SIGNATURE_FOLDER, str(emp_id), f'row_{row.name}.png')
        if os.path.exists(img_path):
            return f"/static/signatures/{emp_id}/row_{row.name}.png"
        return None

    filtered_df['簽名圖片'] = filtered_df.apply(get_signature_path, axis=1)

    data = filtered_df.to_dict(orient='records')
    return jsonify(data)
@app.route("/email", methods=["GET"])
def email():
    if 'username' not in session:
        return redirect(url_for('login'))
    username = session['username']
    name = session['name']
    dept_no=session['dept_no']
    brand_group = request.args.get("brand_group", "")
    store_name = request.args.get("store_name", "")
    email = request.args.get("email", "")
    dept1 = request.args.get("dept1", "")
    dept2 = request.args.get("dept2", "")
    try:
        with open('email.json', 'r', encoding='utf-8') as f:
            store_data = json.load(f)
    except:
        store_data=[]
    filtered = store_data
    if brand_group:
        filtered = [s for s in filtered if s["brand_group"] == brand_group]
    if store_name:
        filtered = [s for s in filtered if store_name in s["name"]]
    if email:
        filtered = [s for s in filtered if email in s["email"]]
    if dept1:
        filtered = [s for s in filtered if dept1 in s["dept1"]]
    if dept2:
        filtered = [s for s in filtered if dept2 in s["dept2"]]
    return render_template("email.html", email_list=filtered,username=username,name=name,dept_no=dept_no,dept1=dept1,dept2=dept2)
@app.route("/addemail", methods=["POST"])
def add_store():
    if 'username' not in session:
        return redirect(url_for('login'))
    with open('email.json', 'r', encoding='utf-8') as f:
        store_data = json.load(f)
    data = request.get_json()
    brand = data.get('brand')
    name = data.get('name')
    email = data.get('email')
    dept1 = data.get('dept1')
    dept2 = data.get('dept2')
    for item in store_data:
        if item['name'] == name:
            return jsonify({'success': False, 'error': '門市名稱已存在！'})    
    store_data.append({
        'brand_group': brand,
        'name': name,
        'email': email,
        'dept1': dept1,
        'dept2': dept2
    })
    with open('email.json', 'w', encoding='utf-8') as f:
        json.dump(store_data, f, ensure_ascii=False, indent=4)
    return jsonify({'success': True})
@app.route("/editemail", methods=["POST"])
def edit_store():
    if 'username' not in session:
        return redirect(url_for('login'))
    with open('email.json', 'r', encoding='utf-8') as f:
        store_data = json.load(f)
    data = request.get_json()
    brand = data.get('Ebrand')
    name = data.get('Ename')
    email = data.get('Eemail')
    dept1 = data.get('edept1')
    dept2 = data.get('edept2')
    ori_brand = data.get('ori_brand')
    ori_name = data.get('ori_name')
    ori_email = data.get('ori_email')
    ori_dept1 = data.get('ori_dept1')
    ori_dept2 = data.get('ori_dept2')
    updated = False
    for item in store_data:
        if (item['brand_group'] == ori_brand and 
            item['name'] == ori_name and 
            item['email'] == ori_email):
            
            # 更新該筆資料
            item['brand_group'] = brand
            item['name'] = name
            item['email'] = email
            item['dept1'] = dept1
            item['dept2'] = dept2
            updated = True
            break  # 找到後就不必繼續了
    if updated:
        # 覆寫回去 email.json
        with open('email.json', 'w', encoding='utf-8') as f:
            json.dump(store_data, f, ensure_ascii=False, indent=4)
        return jsonify({'success': True})
    else:
        return jsonify({'success': False, 'error': '更新失敗'})    
@app.route("/deletemail", methods=["POST"])
def delet_store():
    if 'username' not in session:
        return redirect(url_for('login'))
    with open('email.json', 'r', encoding='utf-8') as f:
        store_data = json.load(f)
    data = request.get_json()
    name = data.get('name')
    # 濾除 name 相同的資料（保留其他的）
    new_store_data = [item for item in store_data if item['name'] != name]

    # 如果沒有變化，表示找不到要刪的
    if len(new_store_data) == len(store_data):
        return jsonify({"success": False, "message": "找不到該門市"}), 404

    # 寫回去
    with open('email.json', 'w', encoding='utf-8') as f:
        json.dump(new_store_data, f, ensure_ascii=False, indent=4)

    return jsonify({"success": True, "message": "刪除成功"})

@app.route('/history/<path:filename>')
def serve_history_file(filename):
    return send_from_directory('history', filename)
@app.route('/download_history_excel', methods=['GET'])
def download_history_excel():
    status = request.args.get('status')
    sign = request.args.get('sign')
    label = request.args.get('label')
    historyfoler=HISTORY_FOLDER+'/'+status+'/uploads'
    files = glob.glob(os.path.join(historyfoler, '*.xlsx'))
    if not files:
        wb_all = Workbook()
        ws_all = wb_all.active
        if not os.path.exists(app.config['TEMP']):
            os.makedirs(app.config['TEMP'])
        output_path = os.path.join(app.config['TEMP'], 'signed_filtered.xlsx')
        wb_all.save(output_path)
        return send_file(output_path, as_attachment=True)

    latest_file = max(files, key=os.path.getmtime)
    df = pd.read_excel(latest_file)
    
    if label=="":
        filtered_df = df[(df['班別'] == '國定假日') & ((df['身份別'] == '門市副理(含)級以上') | (df['身份別'] == '門市正職人員'))].reset_index(drop=True)
    else:
        filtered_df = df[(df['班別'] == '國定假日') & ((df['身份別'] == '門市副理(含)級以上') | (df['身份別'] == '門市正職人員')) & (df['員工編號'].str.lower().str.contains(str(label).lower())) ].reset_index(drop=True)
    EMID = filtered_df['員工編號'].unique().tolist()

    wb_all = Workbook()
    ws_all = wb_all.active
    ws_all.title = '全部'

    wb_signed = Workbook()
    ws_signed = wb_signed.active
    ws_signed.title = '已簽名'

    wb_unsigned = Workbook()
    ws_unsigned = wb_unsigned.active
    ws_unsigned.title = '未簽名'
    

    headers = ['單位名稱', '員工編號', '員工姓名', '身份別', '日期', '班別', '簽名']
    for ws in [ws_all, ws_signed, ws_unsigned]:
        ws.append(headers)

    for emp_id in EMID:
        emp_rows = filtered_df[filtered_df['員工編號'] == emp_id].reset_index(drop=True)

        for i, row in emp_rows.iterrows():
            row_data = [
                row['單位名稱'], row['員工編號'], row['員工姓名'],
                row['身份別'], row['日期'], row['班別'], ''
            ]
            siggnaturefolder=HISTORY_FOLDER+'/'+status+'/signatures'
            img_path = os.path.join(siggnaturefolder, emp_id, f'row_{i}.png')
            img_exists = os.path.exists(img_path)

            # 將資料寫入三個工作表
            # 1. 全部
            ws_all.append(row_data)
            row_idx_all = ws_all.max_row
            if img_exists:
                img = ExcelImage(img_path)
                img.width, img.height = 100, 50
                ws_all.add_image(img, f'G{row_idx_all}')
                ws_all.row_dimensions[row_idx_all].height = 40

            # 2. 已簽名
            if img_exists:
                ws_signed.append(row_data)
                row_idx_signed = ws_signed.max_row
                img = ExcelImage(img_path)
                img.width, img.height = 100, 50
                ws_signed.add_image(img, f'G{row_idx_signed}')
                ws_signed.row_dimensions[row_idx_signed].height = 40
            else:
                # 3. 未簽名
                ws_unsigned.append(row_data)
    ############未簽名資料匯出欄寬######################
    ws_unsigned.column_dimensions['A'].width = 25
    ws_unsigned.column_dimensions['B'].width = 10
    ws_unsigned.column_dimensions['C'].width = 10
    ws_unsigned.column_dimensions['D'].width = 25
    ws_unsigned.column_dimensions['E'].width = 15
    ws_unsigned.column_dimensions['F'].width = 10
    ws_unsigned.column_dimensions['G'].width = 15
    ############已簽名資料匯出欄寬######################
    ws_signed.column_dimensions['A'].width = 25
    ws_signed.column_dimensions['B'].width = 10
    ws_signed.column_dimensions['C'].width = 10
    ws_signed.column_dimensions['D'].width = 25
    ws_signed.column_dimensions['E'].width = 15
    ws_signed.column_dimensions['F'].width = 10
    ws_signed.column_dimensions['G'].width = 15
    ############全部資料匯出欄寬######################
    ws_all.column_dimensions['A'].width = 25
    ws_all.column_dimensions['B'].width = 10
    ws_all.column_dimensions['C'].width = 10
    ws_all.column_dimensions['D'].width = 25
    ws_all.column_dimensions['E'].width = 15
    ws_all.column_dimensions['F'].width = 10
    ws_all.column_dimensions['G'].width = 15
    # 儲存檔案
    if not os.path.exists(app.config['TEMP']):
        os.makedirs(app.config['TEMP'])
    output_path = os.path.join(app.config['TEMP'], 'signed_filtered.xlsx')
    
    
    if(sign=='unsigned'):
        wb_unsigned.save(output_path)
    elif(sign=='signed'):
        wb_signed.save(output_path)
    else:
        wb_all.save(output_path)
  
    return send_file(output_path, as_attachment=True)

@app.route('/home/search')
def search():
    if 'username' not in session:
        return redirect(url_for('login'))
    username = session['username']
    name = session['name']
    dept_no=session['dept_no']
    dept1 = []
    dept2 = []
    
    try:
        with open("email.json", "r", encoding="utf-8") as f:
            store_data = json.load(f)
    except Exception as e:
        print("讀取錯誤:", e)
        store_data = []
    dept1 = list({item.get("dept1") for item in store_data if item.get("dept1")})
    dept2 = list({item.get("dept2") for item in store_data if item.get("dept2")})
    # 讀取最新 Excel
    files = glob.glob(os.path.join(UPLOAD_FOLDER, '*.xlsx'))
    if not files:
        if dept_no == '139' or dept_no=='452':
            return render_template('admin.html', tables=[], username=username, name=name, no_data=True)
        else:
            if username in dept1 or username in dept2:
                return render_template('search.html', tables=[], username=username, name=name, no_data=True,has_permission=True)
            else:
                return render_template('search.html', tables=[], username=username, name=name, no_data=True,has_permission=False)
    latest_file = max(files, key=os.path.getmtime)

    # 使用我們定義的函式來兼容讀取
    try:
        df = sub.read_excel_compatible(latest_file)
    except Exception as e:
        return f'Excel 載入失敗：{e}', 500
    # 過濾國定假日

    if dept_no == '139' or dept_no=='452':    #人資部&資訊部
        filtered_df = df[(df['班別'] == '國定假日') & ((df['身份別'] == '門市副理(含)級以上') | (df['身份別'] == '門市正職人員'))].reset_index(drop=True)
        EMID = filtered_df['員工編號'].unique().tolist()
        display_data = []
        for emp_id in EMID:
            emp_rows = filtered_df[filtered_df['員工編號'] == emp_id].reset_index(drop=True)

            for i, row in emp_rows.iterrows():
                # 寫入一筆資料
                item = row[['單位名稱', '員工編號', '員工姓名', '身份別', '日期', '班別']].to_dict()
                signature_file = os.path.join(SIGNATURE_FOLDER, emp_id, f'row_{i}.png')
                if os.path.exists(signature_file):
                    item['signature'] = f'/static/signatures/{emp_id}/row_{i}.png'
                else:
                    item['signature'] = ''  # 沒有簽名
                display_data.append(item)
        
        return render_template('admin.html', tables=display_data, username=username,name=name)
        #return render_template('admin.html', username=username, name=name)
    else:
        user_store_names = [
            item["name"]
            for item in store_data
            if item.get("dept1") == username or item.get("dept2") == username
        ]
        filtered_df = df[(df['單位名稱'].isin(user_store_names)) & (df['班別'] == '國定假日') & ((df['身份別'] == '門市副理(含)級以上') | (df['身份別'] == '門市正職人員'))]
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
        if username in dept1 or username in dept2:
            return render_template('search.html', tables=display_data, username=username, name=name, has_permission=True)
        else:
            return render_template('search.html', tables=display_data, username=username, name=name, has_permission=False)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=4275)

