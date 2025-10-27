from dotenv import load_dotenv
from datetime import datetime
import os
import sys
ENV = './.env' 
load_dotenv(dotenv_path=ENV)
import paramiko
HR_FOLDER='hr'
HISTORY_FOLDER = 'history'
TMP_FOLDER ='static/year_signed_docs'
def shift_letter(word):
    result = []  # 用列表暫存新字母
    for c in word:
        if c.isalpha():
            base = ord('A') if c.isupper() else ord('a')
            c = chr((ord(c) - base + 1) % 26 + base)
        result.append(c)  # 加入列表
    return ''.join(result)
def insert_kz(ext):
    data="kz"
    j=0
    new_ext=''
    for i in range(len(ext)):
        new_ext=new_ext+data[j%2]+ext[i]
        j+=1
    new_ext=shift_letter(new_ext)
    return new_ext

def mkdir_p(sftp, remote_directory):

    if remote_directory.startswith('./'):
        remote_directory = remote_directory[2:]
    if remote_directory.startswith('../'):
        remote_directory = remote_directory[3:]
    dirs = remote_directory.split('/')
    path = ''  

    for dir in dirs:
        if dir == '':
            continue
        if path == '':
            path = dir
        else:
            path = f"{path}/{dir}"

        try:
            sftp.stat(path)  
        except FileNotFoundError:
            sftp.mkdir(path)
def upload(LOC_FOLDER, NAS_FOLDER, sftp):
    if os.path.isdir(LOC_FOLDER):  # 是資料夾，遞迴
        for item in os.listdir(LOC_FOLDER):
            local_path = os.path.join(LOC_FOLDER, item)
            upload(local_path, NAS_FOLDER, sftp)
    elif os.path.isfile(LOC_FOLDER):  # 是檔案，處理上傳
        relative_path = os.path.relpath(LOC_FOLDER, start=HR_FOLDER)
        remote_file = os.path.join(NAS_FOLDER, relative_path).replace("\\", "/")
        if remote_file.startswith('../'):
            remote_file = remote_file[1:]
        base, ext = os.path.splitext(remote_file)
        ext = ext.lstrip('.') 
        new_ext = insert_kz(ext)
        remote_file = f"{base}.{new_ext}"

        remote_folder = os.path.dirname(remote_file)
        

        # # 建立遠端資料夾
        mkdir_p(sftp, remote_folder)

        # 判斷遠端檔案是否存在，避免重複上傳
        try:
            sftp.stat(remote_file)#存在，跳過
        except FileNotFoundError:
            sftp.put(LOC_FOLDER, remote_file) #不存在，上傳
    else:
        print(f"忽略非檔案資料夾: {LOC_FOLDER}")
# # SFTP 連線資料
try:
    # 連線資訊
    host = os.getenv('HRFTP_host')
    port = os.getenv('HRFTP_port')
    username = os.getenv('HRFTP_uid')
    password = os.getenv('HRFTP_password')

    # 創建SSH客戶端
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(host, port, username, password)
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(f'{now} | 連線成功\n')
except:
    sys.exit()
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(f'{now} | 連線失敗\n')
# 創建SFTP客戶端

sftp = ssh.open_sftp()
# sftp.chdir("hrm_signature")
sftp.chdir("hrm_signature")
NAS_filelist = sftp.listdir()
NAS_FOLDER = ""
print("🔍 遠端起始目錄:", sftp.getcwd())
# LOC_file_list = os.listdir(HISTORY_FOLDER)
# 上傳history檔案
try:
    upload(HISTORY_FOLDER,NAS_FOLDER,sftp)
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(f'{now} | 上傳歷史資料成功\n')
except:
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(f'{now} | 上傳歷史資料失敗\n')
# # 上傳tmp檔案
try:
    upload(TMP_FOLDER,NAS_FOLDER,sftp)
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(f'{now} | 上傳歷史資料成功\n')
except:
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(f'{now} | 上傳歷史資料失敗\n')

# 關閉
sftp.close()
ssh.close()
