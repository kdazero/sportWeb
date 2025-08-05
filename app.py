import os
import re
import sys
import json
import requests
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, session, Response, flash
from bs4 import BeautifulSoup
from datetime import datetime
from io import BytesIO
from urllib.parse import quote
import threading

# --- Google Sheets & ReportLab ---
import gspread
from google.oauth2.service_account import Credentials
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import cm

# ==============================================================================
# Flask 應用程式設定
# ==============================================================================
app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', os.urandom(24))

# --- 路徑設定 ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FONT_PATH = os.path.join(BASE_DIR, 'NotoSansTC-Regular.ttf')

# --- 字型註冊 ---
if os.path.exists(FONT_PATH):
    pdfmetrics.registerFont(TTFont('NotoSansTC', FONT_PATH))
    FONT_AVAILABLE = True
else:
    print("警告：中文字型 'NotoSansTC-Regular.ttf' 未找到。PDF 中的中文可能無法正常顯示。", file=sys.stderr)
    FONT_AVAILABLE = False

# --- Google Sheets 設定 ---
try:
    # Render 會將 JSON 內容作為一個字串存入環境變數
    # 我們需要將其解析回字典
    creds_json_str = os.environ.get('GOOGLE_CREDENTIALS_JSON')
    if not creds_json_str:
        raise ValueError("環境變數 'GOOGLE_CREDENTIALS_JSON' 未設定。")
    
    creds_info = json.loads(creds_json_str)
    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
    gc = gspread.authorize(creds)
    
    # 從環境變數讀取 Google Sheet 的名稱
    SHEET_NAME = os.environ.get('GOOGLE_SHEET_NAME')
    if not SHEET_NAME:
        raise ValueError("環境變數 'GOOGLE_SHEET_NAME' 未設定。")
        
    spreadsheet = gc.open(SHEET_NAME)
    # 假設您的工作表名稱是 'users'
    worksheet = spreadsheet.worksheet('users') 
    print("成功連接至 Google Sheets。")
    GSPREAD_AVAILABLE = True
except Exception as e:
    print(f"錯誤：無法連接至 Google Sheets。請檢查您的環境變數和憑證設定。 {e}", file=sys.stderr)
    GSPREAD_AVAILABLE = False
    
# --- 全域變數 ---
excel_lock = threading.Lock()

# ==============================================================================
# 輔助函式
# ==============================================================================
def get_user_data():
    """從 Google Sheets 讀取使用者資料並轉換為 DataFrame。"""
    if not GSPREAD_AVAILABLE:
        return None
    try:
        with excel_lock:
            records = worksheet.get_all_records()
        return pd.DataFrame(records)
    except Exception as e:
        print(f"讀取 Google Sheet 時發生錯誤：{e}", file=sys.stderr)
        return None

def seconds_to_hms(seconds):
    """將秒數轉換為 時:分:秒 的格式。"""
    if seconds is None:
        return "N/A"
    hours, remainder = divmod(seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return f"{int(hours):02}:{int(minutes):02}:{int(seconds):02}"

def update_user_log_timestamp(username):
    """在 Google Sheet 中更新使用者的證書產生時間戳記。"""
    if not GSPREAD_AVAILABLE:
        return
    with excel_lock:
        try:
            # 找到對應的使用者所在的列
            cell = worksheet.find(username, in_column=1) # 假設 username 在 A 欄
            if not cell:
                print(f"紀錄失敗：在 Google Sheet 中找不到使用者 {username}。", file=sys.stderr)
                return

            # 找到 'last_certificate_timestamp' 所在的欄
            headers = worksheet.row_values(1)
            try:
                col_index = headers.index('last_certificate_timestamp') + 1
            except ValueError:
                # 如果欄位不存在，則在最後一欄新增
                col_index = len(headers) + 1
                worksheet.update_cell(1, col_index, 'last_certificate_timestamp')

            # 更新該儲存格
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            worksheet.update_cell(cell.row, col_index, timestamp)
            print(f"紀錄成功：使用者 {username} 的證書產生時間已更新為 {timestamp}。")

        except Exception as e:
            print(f"更新 Google Sheet 時發生錯誤：{e}", file=sys.stderr)

# ==============================================================================
# 路由 (Routes) - (此部分與前一版本幾乎相同，僅錯誤處理訊息稍作調整)
# ==============================================================================
@app.route('/', methods=['GET', 'POST'])
def index():
    if 'username' not in session:
        return redirect(url_for('login'))
    activities_data = session.get('activities_data', [])
    return render_template('index.html', activities=activities_data)

@app.route('/login', methods=['GET', 'POST'])
def login():
    if not GSPREAD_AVAILABLE:
        flash('後端資料庫服務異常，請聯繫管理員。', 'danger')
        return render_template('login.html')

    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        
        users_df = get_user_data()
        if users_df is None:
            flash('伺服器錯誤，無法讀取使用者資料。', 'danger')
            return render_template('login.html')

        user = users_df[users_df['username'] == username]
        if not user.empty and str(user.iloc[0]['password']) == str(password):
            session['username'] = username
            session['garmin_url'] = user.iloc[0]['garmin_url']
            return redirect(url_for('fetch_activities'))
        else:
            flash('帳號或密碼錯誤。', 'danger')

    return render_template('login.html')

@app.route('/fetch_activities')
def fetch_activities():
    if 'garmin_url' not in session:
        flash('請先登入。', 'warning')
        return redirect(url_for('login'))

    garmin_url = session['garmin_url']
    try:
        response = requests.get(garmin_url)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        
        script_tag = soup.find('script', string=re.compile('VIEWER_USER_PREFERENCES'))
        if not script_tag:
            flash('在頁面中找不到活動資料。請確認您的個人資料頁面是公開的。', 'danger')
            return redirect(url_for('index'))

        json_str = script_tag.string
        json_data_match = re.search(r'=\s*(\{.*?\});', json_str, re.DOTALL)
        if not json_data_match:
            flash('無法解析活動資料。', 'danger')
            return redirect(url_for('index'))

        data = json.loads(json_data_match.group(1))
        activities = data.get('activities', [])
        session['activities_data'] = activities
        flash('成功抓取活動資料！', 'success')

    except requests.exceptions.RequestException as e:
        flash(f'抓取資料失敗：{e}', 'danger')
    except json.JSONDecodeError:
        flash('解析 JSON 資料失敗。', 'danger')
    except Exception as e:
        flash(f'發生未知錯誤：{e}', 'danger')
        
    return redirect(url_for('index'))

@app.route('/generate_pdf/<activity_id>')
def generate_pdf(activity_id):
    if 'username' not in session:
        return redirect(url_for('login'))

    activities = session.get('activities_data', [])
    activity = next((act for act in activities if str(act['activityId']) == activity_id), None)

    if not activity:
        return "找不到活動", 404

    activity_name = activity.get('activityName', 'N/A')
    distance_meters = activity.get('distance', 0)
    distance_km = round(distance_meters / 1000, 2) if distance_meters else 0
    moving_time_seconds = activity.get('movingTime', 0)
    moving_time_hms = seconds_to_hms(moving_time_seconds)
    
    update_user_log_timestamp(session.get('username'))

    buffer = BytesIO()
    p = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter

    if FONT_AVAILABLE:
        p.setFont('NotoSansTC', 36)
    else:
        p.setFont('Helvetica', 36)

    p.drawCentredString(width / 2.0, height - 100, "完賽證明")

    if FONT_AVAILABLE:
        p.setFont('NotoSansTC', 18)
    else:
        p.setFont('Helvetica', 18)

    p.drawCentredString(width / 2.0, height - 200, f"恭喜 {session.get('username', '參賽者')}")
    p.drawCentredString(width / 2.0, height - 250, "成功挑戰")
    p.drawCentredString(width / 2.0, height - 300, f"{activity_name}")
    p.drawCentredString(width / 2.0, height - 350, f"距離：{distance_km} 公里")
    p.drawCentredString(width / 2.0, height - 400, f"成績：{moving_time_hms}")

    p.showPage()
    p.save()
    buffer.seek(0)
    
    safe_activity_name = quote(activity_name.replace(" ", "_"))
    filename = f"certificate_{session.get('username', '')}_{safe_activity_name}.pdf"

    return Response(
        buffer,
        mimetype='application/pdf',
        headers={'Content-Disposition': f'attachment;filename={filename}'}
    )

@app.route('/logout')
def logout():
    session.clear()
    flash('您已成功登出。', 'info')
    return redirect(url_for('login'))

# ==============================================================================
# 應用程式啟動
# ==============================================================================
if __name__ == '__main__':
    # 在本機測試時，您可能需要手動設定環境變數
    # 例如：os.environ['SECRET_KEY'] = 'your_local_secret_key'
    # ...等等
    app.run(debug=True)
