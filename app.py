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
    print("診斷：開始設定 Google Sheets 連線...")
    
    creds_json_str = os.environ.get('GOOGLE_CREDENTIALS_JSON')
    if not creds_json_str:
        raise ValueError("環境變數 'GOOGLE_CREDENTIALS_JSON' 未設定。")
    
    print("診斷：成功讀取 GOOGLE_CREDENTIALS_JSON。")
    creds_info = json.loads(creds_json_str)
    print("診斷：成功解析 JSON 憑證。")
    
    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive.file'
    ]
    
    creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
    print("診斷：成功建立 Credentials 物件。")
    
    gc = gspread.authorize(creds)
    print("診斷：成功通過 gspread 授權。")
    
    # --- 改用 GOOGLE_SHEET_KEY 來開啟 ---
    SHEET_KEY = os.environ.get('GOOGLE_SHEET_KEY')
    if not SHEET_KEY:
        raise ValueError("環境變數 'GOOGLE_SHEET_KEY' 未設定。")
    
    print(f"診斷：準備透過 KEY 開啟試算表...")
    spreadsheet = gc.open_by_key(SHEET_KEY)
    print("診斷：成功開啟試算表檔案。")
    
    worksheet = spreadsheet.worksheet('users') 
    print("診斷：成功開啟 'users' 工作表。")
    
    print("成功連接至 Google Sheets。")
    GSPREAD_AVAILABLE = True

except Exception as e:
    # --- 增加更詳細的錯誤日誌 ---
    print("錯誤：在連接 Google Sheets 的過程中發生預期外的錯誤。", file=sys.stderr)
    print(f"錯誤類型: {type(e)}", file=sys.stderr)
    print(f"錯誤詳細資訊: {repr(e)}", file=sys.stderr)
    GSPREAD_AVAILABLE = False
    
# --- 全域變數 ---
gsheet_lock = threading.Lock()

# ==============================================================================
# 輔助函式
# ==============================================================================
def get_user_data():
    """從 Google Sheets 讀取使用者資料並轉換為 DataFrame。"""
    if not GSPREAD_AVAILABLE:
        return None
    try:
        with gsheet_lock:
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

def update_user_log_timestamp(user_id_card):
    """在 Google Sheet 中更新使用者的證書產生時間戳記。"""
    if not GSPREAD_AVAILABLE:
        return
    with gsheet_lock:
        try:
            # 找到對應的使用者所在的列 (根據 id_card)
            cell = worksheet.find(user_id_card, in_column=2) # 假設 id_card 在 B 欄
            if not cell:
                print(f"紀錄失敗：在 Google Sheet 中找不到使用者 {user_id_card}。", file=sys.stderr)
                return

            # 找到 'last_print' 所在的欄
            headers = worksheet.row_values(1)
            try:
                col_index = headers.index('last_print') + 1
            except ValueError:
                # 如果欄位不存在，則在最後一欄新增
                col_index = len(headers) + 1
                worksheet.update_cell(1, col_index, 'last_print')

            # 更新該儲存格
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            worksheet.update_cell(cell.row, col_index, timestamp)
            print(f"紀錄成功：使用者 {user_id_card} 的證書產生時間已更新為 {timestamp}。")

        except Exception as e:
            print(f"更新 Google Sheet 時發生錯誤：{e}", file=sys.stderr)

# ==============================================================================
# 路由 (Routes)
# ==============================================================================
@app.route('/', methods=['GET', 'POST'])
def index():
    if 'user_id_card' not in session:
        return redirect(url_for('login'))
    activities_data = session.get('activities_data', [])
    return render_template('index.html', activities=activities_data)

@app.route('/login', methods=['GET', 'POST'])
def login():
    if not GSPREAD_AVAILABLE:
        flash('後端資料庫服務異常，請聯繫管理員。', 'danger')
        return render_template('login.html')

    if request.method == 'POST':
        id_card = request.form['id_card']
        phone = request.form['phone']
        
        users_df = get_user_data()
        if users_df is None:
            flash('伺服器錯誤，無法讀取使用者資料。', 'danger')
            return render_template('login.html')

        # 確保 phone 欄位是字串格式以進行比較
        users_df['phone'] = users_df['phone'].astype(str)
        user = users_df[users_df['id_card'] == id_card]

        if not user.empty and user.iloc[0]['phone'] == phone:
            # 登入成功，將所需資訊存入 session
            session['user_id_card'] = id_card
            session['user_name'] = user.iloc[0]['name']
            session['user_number'] = user.iloc[0]['user_number']
            session['garmin_url'] = user.iloc[0]['garmin_url'] # 假設您的 sheet 有 garmin_url 欄位
            return redirect(url_for('fetch_activities'))
        else:
            flash('身分證號或手機號碼錯誤。', 'danger')

    return render_template('login.html')

@app.route('/fetch_activities')
def fetch_activities():
    if 'garmin_url' not in session:
        flash('請先登入。', 'warning')
        return redirect(url_for('login'))

    garmin_url = session['garmin_url']
    if not garmin_url or not isinstance(garmin_url, str) or not garmin_url.startswith('http'):
        flash('此帳號未設定有效的 Garmin Connect 網址。', 'danger')
        return redirect(url_for('index'))
        
    try:
        response = requests.get(garmin_url)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        
        script_tag = soup.find('script', string=re.compile('VIEWER_USER_PREFERENCES'))
        if not script_tag:
            flash('在頁面中找不到活動資料。請確認您的 Garmin Connect 個人資料頁面是公開的。', 'danger')
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
        flash(f'抓取 Garmin 資料失敗：{e}', 'danger')
    except json.JSONDecodeError:
        flash('解析 Garmin JSON 資料失敗。', 'danger')
    except Exception as e:
        flash(f'發生未知錯誤：{e}', 'danger')
        
    return redirect(url_for('index'))

@app.route('/generate_pdf/<activity_id>')
def generate_pdf(activity_id):
    if 'user_id_card' not in session:
        return redirect(url_for('login'))

    activities = session.get('activities_data', [])
    activity = next((act for act in activities if str(act['activityId']) == activity_id), None)

    if not activity:
        return "找不到活動", 404

    # --- 更新使用者紀錄 ---
    update_user_log_timestamp(session.get('user_id_card'))

    # --- 提取資料 ---
    activity_name = activity.get('activityName', 'N/A')
    distance_meters = activity.get('distance', 0)
    distance_km = round(distance_meters / 1000, 2) if distance_meters else 0
    moving_time_seconds = activity.get('movingTime', 0)
    moving_time_hms = seconds_to_hms(moving_time_seconds)
    
    # --- 產生 PDF ---
    buffer = BytesIO()
    p = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter

    # --- 證書內容 ---
    if FONT_AVAILABLE:
        p.setFont('NotoSansTC', 36)
    else:
        p.setFont('Helvetica', 36)
    p.drawCentredString(width / 2.0, height - 100, "完賽證明")

    if FONT_AVAILABLE:
        p.setFont('NotoSansTC', 18)
    else:
        p.setFont('Helvetica', 18)

    user_name = session.get('user_name', '')
    user_number = session.get('user_number', '')

    p.drawCentredString(width / 2.0, height - 200, f"恭喜 {user_name} (編號: {user_number})")
    p.drawCentredString(width / 2.0, height - 250, "成功挑戰")
    p.drawCentredString(width / 2.0, height - 300, f"{activity_name}")
    p.drawCentredString(width / 2.0, height - 350, f"距離：{distance_km} 公里")
    p.drawCentredString(width / 2.0, height - 400, f"成績：{moving_time_hms}")

    p.showPage()
    p.save()
    buffer.seek(0)
    
    # --- 設定檔名 ---
    safe_activity_name = quote(activity_name.replace(" ", "_"))
    filename = f"certificate_{user_name}_{safe_activity_name}.pdf"

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
    app.run(debug=True)
