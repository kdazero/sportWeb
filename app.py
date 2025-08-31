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
    creds_json_str = os.environ.get('GOOGLE_CREDENTIALS_JSON')
    if not creds_json_str:
        raise ValueError("環境變數 'GOOGLE_CREDENTIALS_JSON' 未設定。")
    creds_info = json.loads(creds_json_str)
    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive.file'
    ]
    creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
    gc = gspread.authorize(creds)
    SHEET_KEY = os.environ.get('GOOGLE_SHEET_KEY')
    if not SHEET_KEY:
        raise ValueError("環境變數 'GOOGLE_SHEET_KEY' 未設定。")
    spreadsheet = gc.open_by_key(SHEET_KEY)
    worksheet = spreadsheet.worksheet('users') 
    print("成功連接至 Google Sheets。")
    GSPREAD_AVAILABLE = True
except Exception as e:
    print("錯誤：在連接 Google Sheets 的過程中發生預期外的錯誤。", file=sys.stderr)
    print(f"錯誤類型: {type(e)}", file=sys.stderr)
    print(f"錯誤詳細資訊: {repr(e)}", file=sys.stderr)
    GSPREAD_AVAILABLE = False
    
# --- 全域變數 ---
gsheet_lock = threading.Lock()

# ==============================================================================
# 資料爬取與處理 (依照本地成功版本邏輯，並標準化輸出)
# ==============================================================================
def get_strava_data(url):
    """從 Strava 活動頁面抓取資料，並返回標準化格式的活動列表。"""
    try:
        response = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'})
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')

        activity_name_tag = soup.find('h1', class_='title')
        activity_name = activity_name_tag.get_text(strip=True) if activity_name_tag else '未知活動'
        
        details_div = soup.find('div', class_='details-container')
        if not details_div:
            raise ValueError("在頁面中找不到活動詳細資料區塊。")

        distance_str = "0"
        time_str = "0h0m0s"
        
        # 尋找包含 Distance 和 Moving Time 的區塊
        for item in details_div.find_all('div', class_='detail'):
            label = item.find('div', class_='label').get_text(strip=True)
            value = item.find('div', class_='value').get_text(strip=True)
            if 'Distance' in label:
                distance_str = value
            elif 'Moving Time' in label:
                time_str = value

        # 處理距離
        distance_km = 0.0
        cleaned_distance_str = distance_str.replace(',', '')
        if 'km' in cleaned_distance_str:
            distance_km = float(re.findall(r"[\d\.]+", cleaned_distance_str)[0])
        elif 'm' in cleaned_distance_str:
            distance_km = float(re.findall(r"[\d\.]+", cleaned_distance_str)[0]) / 1000

        # 處理時間
        time_seconds = 0
        if 'h' in time_str:
            time_seconds += int(time_str.split('h')[0]) * 3600
            time_str = time_str.split('h')[1] if 'h' in time_str else ''
        if 'm' in time_str:
            time_seconds += int(time_str.split('m')[0]) * 60
            time_str = time_str.split('m')[1] if 'm' in time_str else ''
        if 's' in time_str:
            time_seconds += int(time_str.split('s')[0])
            
        # 返回標準化格式的活動列表 (即使只有一個)
        return [{
            'id': 0, # Strava 單頁活動給予一個固定的 ID
            'name': activity_name,
            'distance_km': round(distance_km, 2),
            'time_seconds': time_seconds
        }]
    except Exception as e:
        print(f"抓取 Strava 資料時發生錯誤: {e}", file=sys.stderr)
        return None

def get_garmin_data(url):
    """從 Garmin Connect 頁面抓取資料，並返回標準化格式的活動列表。"""
    try:
        response = requests.get(url)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        
        script_tag = soup.find('script', string=re.compile('VIEWER_USER_PREFERENCES'))
        if not script_tag:
            raise ValueError('在頁面中找不到活動資料腳本。')

        json_str = script_tag.string
        json_data_match = re.search(r'=\s*(\{.*?\});', json_str, re.DOTALL)
        if not json_data_match:
            raise ValueError('無法解析活動資料 JSON。')

        data = json.loads(json_data_match.group(1))
        activities_raw = data.get('activities', [])
        
        activities_standardized = []
        for act in activities_raw:
            activities_standardized.append({
                'id': act.get('activityId'),
                'name': act.get('activityName', 'N/A'),
                'distance_km': round(act.get('distance', 0) / 1000, 2),
                'time_seconds': act.get('movingTime', 0)
            })
        return activities_standardized
    except Exception as e:
        print(f"抓取 Garmin 資料時發生錯誤: {e}", file=sys.stderr)
        return None

# ==============================================================================
# 輔助函式
# ==============================================================================
def get_user_data():
    if not GSPREAD_AVAILABLE: return None
    try:
        with gsheet_lock: records = worksheet.get_all_records()
        return pd.DataFrame(records) if records else pd.DataFrame()
    except Exception as e:
        print(f"讀取 Google Sheet 時發生錯誤：{e}", file=sys.stderr)
        return None

def seconds_to_hms(seconds):
    if seconds is None: return "N/A"
    seconds = int(seconds)
    hours, remainder = divmod(seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return f"{hours:02}:{minutes:02}:{seconds:02}"

def update_user_log(user_id_card, successful_url):
    if not GSPREAD_AVAILABLE: return
    with gsheet_lock:
        try:
            cleaned_id_card = str(user_id_card).strip().upper()
            id_card_column = [str(val).strip().upper() for val in worksheet.col_values(2)]
            try:
                match_row = id_card_column.index(cleaned_id_card) + 1
            except ValueError:
                print(f"紀錄失敗：在 Google Sheet 中找不到使用者 {cleaned_id_card}。", file=sys.stderr)
                return

            headers = worksheet.row_values(1)
            
            time_col_index = headers.index('last_time') + 1 if 'last_time' in headers else len(headers) + 1
            if 'last_time' not in headers: worksheet.update_cell(1, time_col_index, 'last_time'); headers.append('last_time')
            worksheet.update_cell(match_row, time_col_index, datetime.now().strftime('%Y-%m-%d %H:%M:%S'))

            link_col_index = headers.index('last_link') + 1 if 'last_link' in headers else len(headers) + 1
            if 'last_link' not in headers: worksheet.update_cell(1, link_col_index, 'last_link')
            worksheet.update_cell(match_row, link_col_index, str(successful_url) if successful_url else "")

            print(f"紀錄成功：使用者 {cleaned_id_card} 的 last_time 和 last_link 已更新。")
        except Exception as e:
            print(f"更新 Google Sheet 時發生錯誤：{e}", file=sys.stderr)

# ==============================================================================
# 路由 (Routes)
# ==============================================================================
@app.route('/', methods=['GET', 'POST'])
def index():
    if 'user_id_card' not in session:
        return redirect(url_for('login'))

    if request.method == 'POST':
        session.pop('activities_data', None)
        url = request.form.get('activity_url')
        url_type = request.form.get('url_type')
        activities = None

        if url:
            if url_type == 'garmin':
                activities = get_garmin_data(url)
            elif url_type == 'strava':
                activities = get_strava_data(url)
            else:
                flash('請選擇有效的平台。', 'warning')
        else:
            flash('請輸入活動網址。', 'warning')
        
        if activities is not None:
            for act in activities:
                act['time_hms'] = seconds_to_hms(act.get('time_seconds', 0))
            session['activities_data'] = activities
            session['last_successful_url'] = url
            flash('成功抓取活動資料！', 'success')
        elif url:
            flash('抓取或解析資料失敗，請確認網址是否正確且頁面為公開。', 'danger')
            
        return redirect(url_for('index'))

    activities = session.get('activities_data', [])
    return render_template('index.html', activities=activities)

@app.route('/login', methods=['GET', 'POST'])
def login():
    if 'user_id_card' in session: return redirect(url_for('index'))
    if not GSPREAD_AVAILABLE:
        flash('後端資料庫服務異常，請聯繫管理員。', 'danger')
        return render_template('login.html')

    if request.method == 'POST':
        id_card_input = request.form.get('id_card', '').strip().upper()
        phone_input = request.form.get('phone', '').strip()
        
        users_df = get_user_data()
        if users_df is None or users_df.empty:
            flash('無法讀取或資料庫無使用者資料。', 'danger')
            return render_template('login.html')
        
        if not all(col in users_df.columns for col in ['id_card', 'phone']):
            flash('資料庫欄位設定錯誤，請聯繫管理員。', 'danger')
            return render_template('login.html')

        df_id_card_col = users_df['id_card'].astype(str).str.strip().str.upper()
        df_phone_col = users_df['phone'].astype(str).str.strip()
        phone_input_no_zero = phone_input[1:] if phone_input.startswith('0') else phone_input
        
        user = users_df[
            (df_id_card_col == id_card_input) & 
            ((df_phone_col == phone_input) | (df_phone_col == phone_input_no_zero))
        ]

        if not user.empty:
            user_info = user.iloc[0]
            session['user_id_card'] = user_info['id_card']
            session['user_name'] = user_info['name']
            session['user_number'] = user_info['user_number']
            return redirect(url_for('index'))
        else:
            flash('身分證號或手機號碼錯誤。', 'danger')
            return render_template('login.html')

    return render_template('login.html')

@app.route('/generate_pdf/<int:activity_index>')
def generate_pdf(activity_index):
    if 'user_id_card' not in session:
        return redirect(url_for('login'))

    activities = session.get('activities_data', [])
    if not activities or activity_index >= len(activities):
        flash('找不到活動資料或憑證已過期，請重新抓取。', 'danger')
        return redirect(url_for('index'))

    activity = activities[activity_index]
    successful_url = session.get('last_successful_url')
    update_user_log(session.get('user_id_card'), successful_url)

    buffer = BytesIO()
    p = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter

    p.setFont('NotoSansTC' if FONT_AVAILABLE else 'Helvetica', 36)
    p.drawCentredString(width / 2.0, height - 100, "完賽證明")

    p.setFont('NotoSansTC' if FONT_AVAILABLE else 'Helvetica', 18)
    user_name = session.get('user_name', '')
    user_number = session.get('user_number', '')

    p.drawCentredString(width / 2.0, height - 200, f"恭喜 {user_name} (編號: {user_number})")
    p.drawCentredString(width / 2.0, height - 250, "成功挑戰")
    p.drawCentredString(width / 2.0, height - 300, f"{activity.get('name', 'N/A')}")
    p.drawCentredString(width / 2.0, height - 350, f"距離：{activity.get('distance_km', 0)} 公里")
    p.drawCentredString(width / 2.0, height - 400, f"成績：{activity.get('time_hms', 'N/A')}")

    p.showPage()
    p.save()
    buffer.seek(0)
    
    safe_activity_name = quote(activity.get('name', 'activity').replace(" ", "_"))
    filename = f"certificate_{user_name}_{safe_activity_name}.pdf"

    return Response(buffer, mimetype='application/pdf',
                    headers={'Content-Disposition': f'attachment;filename="{filename}"'})

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

