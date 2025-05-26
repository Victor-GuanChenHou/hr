import pyodbc
import pandas as pd
from openpyxl import load_workbook
from dotenv import load_dotenv
import os
ENV = './.env' 
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
print(get_user_info('A14176'))