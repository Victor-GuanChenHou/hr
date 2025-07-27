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
    # SUBSTRING(UIDENTID, 2, LEN(UIDENTID) - 1) AS UIDENTID 身分證後九碼
    cursor.execute("SELECT EMPID, SUBSTRING(UIDENTID, 2, LEN(UIDENTID) - 1) AS UIDENTID, HECNAME ,DEPT_NO ,CLASS FROM HRM.dbo.HRUSER WHERE EMPID = ?", (username,))
    row = cursor.fetchone()
    

    if row:
        cursor.execute("SELECT DEP_NAME,DEP_KIND FROM HRM.dbo.HRUSER_DEPT_BAS WHERE DEP_NO = ?", (row[3],))
        dep_row = cursor.fetchone()
        conn.close()
        if dep_row:
            return {'username': row[0], 'password': row[1], 'name': row[2], 'DEPT_NO':row[3],'CLASS':row[4],'DEPT_NAME':dep_row[0],'DEPT_KIND':dep_row[1]}
        else:
            return {'username': row[0], 'password': row[1], 'name': row[2], 'DEPT_NO':row[3],'CLASS':row[4],'DEPT_NAME':'NOT FOUND','DEPT_KIND':'NOT FOUND'}
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
    cursor.execute("SELECT DEPT_NO FROM HRM.dbo.HRUSER WHERE EMPID = ?", (username,))
    result = cursor.fetchone()
    if result[0]=='193':#判斷是否是央廚(央廚特別處理不抓上司MAIL)
        result='dcz01@kingza.com.tw'
        return result
    else:
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
def get_dep_order(dep_name):
    if dep_name == '杏子豬排營運部':
        return 1
    elif dep_name.startswith('杏子'):
        return 2
    elif dep_name == '段純貞營運部':
        return 3
    elif dep_name.startswith('段純貞'):
        return 4
    elif dep_name.startswith('王將營運'):
        return 5
    elif dep_name.startswith('王將'):
        return 6
    elif dep_name.startswith('京都勝牛營運部'):
        return 7
    elif dep_name.startswith('勝牛'):
        return 8
    elif dep_name.startswith('橋村營運'):
        return 9
    elif dep_name.startswith('橋村'):
        return 10
    else:
        return 99
def docxuser():
    # 取得連線參數
    HRDB_host = os.getenv('HRDB_host')
    HRDB_password = os.getenv('HRDB_password')
    HRDB_uid = os.getenv('HRDB_uid')
    HRDB_name = os.getenv('HRDB_name')

    # 建立資料庫連線
    conn = pyodbc.connect(
        'DRIVER={ODBC Driver 17 for SQL Server};'
        f'SERVER={HRDB_host};'
        f'DATABASE={HRDB_name};'
        f'UID={HRDB_uid};'
        f'PWD={HRDB_password};'
    )
    cursor = conn.cursor()
    # SQL 查詢：取得在職 Class D 員工對應單位與身份別
    query_classd = """
    SELECT 
        D.DEP_NAME AS 單位名稱,
        U.EMPID AS 員工編號,
        U.HECNAME AS 員工姓名,
        T.UTNAME AS 身份別
    FROM HRM.dbo.HRUSER U
    LEFT JOIN HRM.dbo.HRUSER_DEPT_BAS D ON U.DEPT_NO = D.DEP_NO
    LEFT JOIN HRM.dbo.USERTYPE T ON U.UTYPE = T.UTYPE
    WHERE U.STATE = 'A' AND U.Class = 'D'
    """
    cursor.execute(query_classd)
    columns = [column[0] for column in cursor.description]

    # 取得所有資料
    rows = cursor.fetchall()

    # 轉成 DataFrame
    df = pd.DataFrame.from_records(rows, columns=columns)
    # 執行查詢
    #df = pd.read_sql(query_classd, conn)

    # 關閉連線
    conn.close()
     # ➕ 加入排序欄位
    df['dep_order'] = df['單位名稱'].apply(get_dep_order)

    # 🔽 排序
    df = df.sort_values(by=['dep_order', '單位名稱', '員工編號'])

    # 移除排序用欄位再轉成 list[dict]
    df = df.drop(columns=['dep_order'])

    return df.to_dict(orient='records')

