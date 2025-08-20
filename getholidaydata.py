import pyodbc
import pandas as pd
import os
from dotenv import load_dotenv
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings("ignore", category=UserWarning)
ENV = './.env' 
load_dotenv(dotenv_path=ENV)

load_dotenv()
HRDB_host = os.getenv('HRDB_host')
HRDB_password = os.getenv('HRDB_password')
HRDB_uid=os.getenv('HRDB_uid')
HRDB_name=os.getenv('HRDB_name')
# 資料庫連線設定
conn = pyodbc.connect(
    'DRIVER={ODBC Driver 17 for SQL Server};'
    f"SERVER={HRDB_host};"
    f"DATABASE={HRDB_name};"
    f"UID={HRDB_uid};"
    f"PWD={HRDB_password};"
)
#排序
def get_dep_order(dep_name):
    if not isinstance(dep_name, str):
        return 99
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
    elif dep_name.startswith('雞三和營運部'):
        return 11
    elif dep_name.startswith('雞三和'):
        return 12
    else:
        return 99
# 取得今天的日期
today = datetime.today()

# 找出「本月的第一天」再減一天，就會是「上個月的最後一天」
first_day_this_month = today.replace(day=1)
last_day_last_month = first_day_this_month - timedelta(days=1)

# 轉換成 YYMM 格式
yymm_last_month = last_day_last_month.strftime("%Y%m")
# 1. 查詢 CLASSDA
query_classda = f"""
SELECT CPNYID, CLASSDA, EMPID, CLASS
FROM HRM.dbo.CLASSDA
WHERE CLASSDA LIKE'{yymm_last_month}%' AND CLASS='H'
"""
df_classda = pd.read_sql(query_classda, conn)

if df_classda.empty:
    print("查無 CLASSDA 資料")
else:
    # 2. 查詢 HRUSER 中的 EMPID 對應 DEPT_NO, HECNAME, UTYPE
    emp_ids = tuple(df_classda['EMPID'].unique())
    query_hruser = f"""
    SELECT EMPID, DEPT_NO, HECNAME, UTYPE ,STATE ,UIDENTID
    FROM HRM.dbo.HRUSER_{yymm_last_month}
    WHERE EMPID IN {emp_ids}   AND (UTYPE='F' OR UTYPE='H')
    """
    df_hruser = pd.read_sql(query_hruser, conn)
    emp_ids_c = tuple(df_hruser[df_hruser['STATE'] == 'C']['EMPID'].unique())
    query_hruser = f"""
    SELECT EMPID, DEPT_NO, HECNAME, UTYPE ,STATE ,UIDENTID
    FROM HRM.dbo.HRUSER_{yymm_last_month}
    WHERE EMPID IN {emp_ids}   AND (UTYPE='F' OR UTYPE='H')
    """
    df_hruser = pd.read_sql(query_hruser, conn)
    UIDENTID_c = tuple(df_hruser[df_hruser['STATE'] == 'C']['UIDENTID'].unique())

    
    if UIDENTID_c:
        #查 HRUSER（不加月份）找 UIDENTID 有 STATE='A' 的
        query_current = f"""
        SELECT EMPID, DEPT_NO, HECNAME, UTYPE ,STATE ,UIDENTID
        FROM HRM.dbo.HRUSER
        WHERE UIDENTID IN {UIDENTID_c} AND STATE='A' 
        """
        df_current = pd.read_sql(query_current, conn)
        
        # 過濾 C，只保留有對應 A 的
        valid_uid = set(df_current['UIDENTID'])
        df_hruser = df_hruser[~((df_hruser['STATE'] == 'C') & (~df_hruser['UIDENTID'].isin(valid_uid)))]


   
    
    active_emp_ids = df_hruser['EMPID'].unique()
    df_classda = df_classda[df_classda['EMPID'].isin(active_emp_ids)]
    # 3. 合併 CLASSDA + HRUSER
    df_merged = pd.merge(df_classda, df_hruser, on='EMPID', how='left')

    # 4. 查詢 HRUSER_DEPT_BAS 取得 DEP_NAME 和 DEP_CHIEF
    dept_nos = tuple(df_merged['DEPT_NO'].dropna().unique())
    query_dept = f"""
    SELECT DEP_NO, DEP_NAME, DEP_CHIEF
    FROM HRM.dbo.HRUSER_DEPT_BAS
    WHERE DEP_NO IN {dept_nos}
    """
    df_dept = pd.read_sql(query_dept, conn)
    df_merged['DEPT_NO'] = df_merged['DEPT_NO'].astype(str)
    df_dept['DEP_NO'] = df_dept['DEP_NO'].astype(str)

    # 合併取得 DEP_NAME 和 DEP_CHIEF
    df_merged = pd.merge(df_merged, df_dept, left_on='DEPT_NO', right_on='DEP_NO', how='left')

    # 查詢主管 EMAIL
    chief_ids = tuple(df_merged['DEP_CHIEF'].dropna().unique())
    query_chief = f"""
    SELECT EMPID, EMAIL
    FROM HRM.dbo.HRUSER
    WHERE EMPID IN {chief_ids}
    """
    df_chief = pd.read_sql(query_chief, conn)
   # df_chief.rename(columns={'EMPID': 'DEP_CHIEF', 'EMAIL': '主管'}, inplace=True)

    # 合併主管 EMAIL
    df_dept = pd.merge(df_dept, df_chief, left_on='DEP_CHIEF', right_on='EMPID', how='left')
    df_dept = df_dept.rename(columns={'EMAIL': '主管'})

    # 5-3. 特殊處理「段純貞中央工廠」的主管 Email
    df_dept.loc[df_dept['DEP_NAME'] == '段純貞中央工廠', '主管'] = 'dcz01@kingza.com.tw'

    # 合併進 df_merged
    df_merged = pd.merge(df_merged, df_dept[['DEP_NO', '主管']], on='DEP_NO', how='left')
        
    
    # 5. 查詢 USERTYPE 取得身份別名稱
    utypes = tuple(df_merged['UTYPE'].dropna().unique())
    query_utype = f"""
    SELECT UTYPE, UTNAME
    FROM HRM.dbo.USERTYPE
    WHERE UTYPE IN {utypes}
    """
    df_utype = pd.read_sql(query_utype, conn)
    df_merged = pd.merge(df_merged, df_utype, on='UTYPE', how='left')
     # 6. 查詢 CLASSSET 取得班別名稱 CLNAME
    classes = tuple(df_merged['CLASS'].dropna().unique())
    query_classset = f"""
    SELECT CLASS, CLNAME
    FROM HRM.dbo.CLASSSET
    WHERE CLASS = 'H'
    """
    df_classset = pd.read_sql(query_classset, conn)

    df_merged = pd.merge(df_merged, df_classset, on='CLASS', how='left')
    # 7. 最終欄位輸出
    df_result = df_merged[['DEP_NAME', 'EMPID', 'HECNAME', 'CLASSDA','CLNAME', 'UTNAME', '主管']]

    # 改欄位名稱為中文
    df_result = df_result.rename(columns={
        'DEP_NAME': '單位名稱',
        'EMPID': '員工編號',
        'HECNAME': '員工姓名',
        'CLASSDA': '日期',
        'CLNAME' :'班別',
        'UTNAME': '身份別',
        '主管': '主管'
    })

    # 排序
    df_result['DEP_ORDER'] = df_result['單位名稱'].apply(get_dep_order)
    df_result = df_result.sort_values(by=['DEP_ORDER', '單位名稱', '員工編號'])
    df_result = df_result.drop(columns=['DEP_ORDER'])

    # 印出結果
    output_filename = './uploads/upload_month/'+yymm_last_month+'_國定假日.xlsx'  # 你可以動態產生名稱，或固定名稱

    df_result.to_excel(output_filename, index=False)

    print(f"已成功輸出 Excel 檔案：{output_filename}")
