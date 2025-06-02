import pyodbc
import pandas as pd
from openpyxl import load_workbook
from dotenv import load_dotenv
from datetime import datetime
import os
ENV = './.env' 
LOG_DIR='logs'
load_dotenv(dotenv_path=ENV)
def get_user_info(username):
    load_dotenv()
    HRDB_host = os.getenv('HRDB_host')
    HRDB_password = os.getenv('HRDB_password')
    HRDB_uid=os.getenv('HRDB_uid')
    HRDB_name=os.getenv('HRDB_name')
    conn = pyodbc.connect(
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={HRDB_host};"
        f"DATABASE={HRDB_name};"
        f"UID={HRDB_uid};"
        f"PWD={HRDB_password};"
        "Trusted_Connection=no;"
    )
    cursor = conn.cursor()
    cursor.execute("SELECT EMPID, UIDENTID, HECNAME ,DEPT_NO FROM HRM.dbo.HRUSER WHERE EMPID = ?", (username,))
    row = cursor.fetchone()
    

    if row:
        cursor.execute("SELECT DEP_NAME,DEP_KIND FROM HRM.dbo.HRUSER_DEPT_BAS WHERE DEP_NO = ?", (row[3],))
        dep_row = cursor.fetchone()
        conn.close()
        if dep_row:
            return {'username': row[0], 'password': row[1], 'name': row[2], 'DEPT_NO':row[3],'DEPT_NAME':dep_row[0],'DEPT_KIND':dep_row[1]}
        else:
            return {'username': row[0], 'password': row[1], 'name': row[2], 'DEPT_NO':row[3],'DEPT_NAME':'NOT FOUND','DEPT_KIND':'NOT FOUND'}
    else:
        return None






def read_excel_compatible(filepath):
    try:
        df = pd.read_excel(filepath, engine='openpyxl')
        return df
    except Exception as e:
        print(f"pandas 讀取失敗，改用 openpyxl 處理: {e}")
        try:
            wb = load_workbook(filepath, data_only=True)

            # ✅ 確認是否有工作表
            if not wb.sheetnames:
                raise ValueError("Excel 檔案中沒有任何工作表")

            sheet = wb.active
            if sheet is None:
                raise ValueError("無法取得有效的工作表")

            rows = list(sheet.iter_rows(values_only=True))
            if not rows or len(rows) < 2:
                raise ValueError("Excel 資料為空，或無有效資料")

            headers = list(rows[0])
            data = [dict(zip(headers, row)) for row in rows[1:] if any(row)]
            return pd.DataFrame(data)

        except Exception as e2:
            raise RuntimeError(f"openpyxl 解析也失敗: {e2}")
def loglogin(username,ip):
    if not os.path.exists(LOG_DIR):
        os.makedirs(LOG_DIR)
    today = datetime.today().strftime('%Y-%m-%d')
    log_file = os.path.join(LOG_DIR, f'{today}.log')
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    log_line = f'{now} | 使用者 {username} 從 {ip} 登入\n'
    
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(log_line)
def find_deptchie(username):
    load_dotenv()
    HRDB_host = os.getenv('HRDB_host')
    HRDB_password = os.getenv('HRDB_password')
    HRDB_uid=os.getenv('HRDB_uid')
    HRDB_name=os.getenv('HRDB_name')
    conn = pyodbc.connect(
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={HRDB_host};"
        f"DATABASE={HRDB_name};"
        f"UID={HRDB_uid};"
        f"PWD={HRDB_password};"
        "Trusted_Connection=no;"
    )
    cursor = conn.cursor()
    cursor.execute("""
        SELECT CHIEF.EMAIL
        FROM HRM.dbo.HRUSER EMP
        JOIN HRM.dbo.HRUSER_DEPT_BAS DEPT
            ON EMP.DEPT_NO = DEPT.DEP_NO
        JOIN HRM.dbo.HRUSER CHIEF
            ON DEPT.DEP_CHIEF = CHIEF.EMPID
        WHERE EMP.EMPID = ?
    """, (username,))
    result = cursor.fetchone()
    cursor.close()
    conn.close()

    if result:
        return result[0]
    else:
        return None
