import os
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
from functools import wraps
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import io
import requests
from bs4 import BeautifulSoup
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
import pytz


# --- Google Sheets 連線設定 ---
# 為了讓登入功能與後續的紀錄功能共用，將連線部分移至全域
worksheet = None
try:
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file("google_credentials.json", scopes=scopes)
    gc = gspread.authorize(creds)
    # !!請將 "測試資料" 換成您的 Google Sheet 檔案名稱!!
    sh = gc.open("測試資料")
    # !!請將 "users" 換成您要操作的工作表名稱!!
    worksheet = sh.worksheet("users")
    print("Google Sheet 連線成功。")
except Exception as e:
    print(f"連線 Google Sheet 失敗，部分功能可能無法使用: {e}")


# --- Flask 應用程式設定 ---
app = Flask(__name__)
app.secret_key = os.urandom(24)

# 註冊中文字體
pdfmetrics.registerFont(TTFont('NotoSansTC', 'NotoSansTC-Regular.ttf'))

# --- Decorators ---
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'username' not in session:
            flash('請先登入以繼續操作', 'warning')
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

# --- 路由 (Routes) ---
@app.route('/login', methods=['GET', 'POST'])
def login():
    """處理使用者登入"""
    if request.method == 'POST':
        # 從 Excel 讀取使用者資料 (保留原始邏輯)
        try:
            df = pd.read_excel('users.xlsx')
        except FileNotFoundError:
            flash('錯誤：找不到 users.xlsx 檔案。', 'danger')
            return render_template('login.html')

        # 取得表單資料
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '').strip()

        # 在 DataFrame 中尋找使用者
        # 【修正點】: 將 DataFrame 中的 phone 欄位和使用者輸入的 password 都轉為字串比較
        user_row = df[(df['id_card'] == username) & (df['phone'].astype(str).str.strip() == password)]

        if not user_row.empty:
            session['username'] = username
            return redirect(url_for('index'))
        else:
            flash('無效的帳號或密碼。', 'danger')
            return render_template('login.html')

    return render_template('login.html')

@app.route('/logout')
def logout():
    session.pop('username', None)
    flash('您已經成功登出', 'success')
    return redirect(url_for('login'))

@app.route('/')
@login_required
def index():
    df = pd.read_excel('users.xlsx')
    user_data = df[df['id_card'] == session['username']].iloc[0]
    return render_template('index.html', user=user_data.to_dict())

@app.route('/generate_cert', methods=['POST'])
@login_required
def generate_cert():
    df = pd.read_excel('users.xlsx')
    user_data = df[df['id_card'] == session['username']].iloc[0].to_dict()

    name = user_data.get('name', 'N/A')
    user_number = user_data.get('user_number', 'N/A')
    url = request.form.get('url')

    if not url:
        flash('請輸入有效的網址', 'danger')
        return redirect(url_for('index'))

    try:
        response = requests.get(url)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # 這裡保留您原有的爬蟲邏輯
        title_element = soup.select_one('h1.title')
        title = title_element.text.strip() if title_element else "無法讀取標題"
        
        distance_element = soup.select_one('li[title="距離"] strong')
        distance = distance_element.text.strip() if distance_element else "無法讀取距離"
        
        time_element = soup.select_one('li[title="經過時間"] strong')
        moving_time = time_element.text.strip() if time_element else "無法讀取時間"
        
        date_element = soup.select_one('time.timestamp')
        date_time = date_element.text.strip().split(',')[1].strip() if date_element else "無法讀取日期"

    except requests.RequestException as e:
        flash(f'爬取資料失敗: {e}', 'danger')
        return redirect(url_for('index'))

    # 生成 PDF
    buffer = io.BytesIO()
    p = canvas.Canvas(buffer, pagesize=letter)
    p.setFont('NotoSansTC', 12)
    
    p.drawString(100, 750, f"姓名: {name}")
    p.drawString(100, 730, f"選手編號: {user_number}")
    p.drawString(100, 710, f"活動名稱: {title}")
    p.drawString(100, 690, f"距離: {distance}")
    p.drawString(100, 670, f"花費時間: {moving_time}")
    p.drawString(100, 650, f"日期: {date_time}")
    p.save()
    
    buffer.seek(0)

    # 【新增功能】: 寫入 Google Sheet
    if worksheet:
        try:
            cell = worksheet.find(session['username'])
            if cell:
                user_row = cell.row
                headers = worksheet.row_values(1)
                
                LAST_PRINT_COL_NAME = 'last_print'
                URL_COL_NAME = 'uploaded_url'

                if URL_COL_NAME not in headers:
                    worksheet.update_cell(1, len(headers) + 1, URL_COL_NAME)
                    headers.append(URL_COL_NAME)
                
                last_print_col_index = headers.index(LAST_PRINT_COL_NAME) + 1
                url_col_index = headers.index(URL_COL_NAME) + 1
                
                taiwan_tz = pytz.timezone('Asia/Taipei')
                current_time = datetime.now(taiwan_tz).strftime('%Y-%m-%d %H:%M:%S')
                
                worksheet.update_cell(user_row, last_print_col_index, current_time)
                worksheet.update_cell(user_row, url_col_index, url)
                flash('證書生成成功，並已記錄時間與網址！', 'success')
            else:
                flash('在 Google Sheet 中找不到您的資料，無法記錄。', 'warning')

        except Exception as e:
            flash(f'記錄至 Google Sheet 時發生錯誤: {e}', 'warning')
    else:
        flash('未連接至 Google Sheet，無法進行記錄。', 'warning')


    return send_file(buffer, as_attachment=True, download_name=f'{user_number}_cert.pdf', mimetype='application/pdf')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)

