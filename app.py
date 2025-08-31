import os
import re
import sys
import json
import requests
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, session, Response, flash
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
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
# 輔助函式
# ==============================================================================
def seconds_to_hms(seconds):
    """將秒數轉換為 時:分:秒 的格式。"""
    if seconds is None: return "N/A"
    seconds = int(seconds)
    hours, remainder = divmod(seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return f"{hours:02}:{minutes:02}:{seconds:02}"

def hms_to_seconds(t):
    """將 '1h 23m 45s' 或 'hh:mm:ss' 或 'mm:ss' 格式的時間字串轉換為總秒數"""
    h, m, s = 0, 0, 0
    if not isinstance(t, str):
        return 0
    
    if 'h' in t or 'm' in t or 's' in t: # 處理 Strava 的 '1h 23m 45s' 格式
        if 'h' in t:
            h = int(t.split('h')[0])
            t = t.split('h')[1].strip()
        if 'm' in t:
            m = int(t.split('m')[0])
            t = t.split('m')[1].strip()
        if 's' in t:
            s = int(t.split('s')[0])
        return h * 3600 + m * 60 + s
    elif ':' in t: # 處理 Garmin meta tag 的 'hh:mm:ss' 或 'mm:ss' 格式
        parts = list(map(int, t.split(':')))
        if len(parts) == 3: # hh:mm:ss
            return parts[0] * 3600 + parts[1] * 60 + parts[2]
        elif len(parts) == 2: # mm:ss
            return parts[0] * 60 + parts[1]
    return 0


def calculate_pace(distance_km, total_seconds):
    """計算配速 (分鐘/公里)"""
    if distance_km == 0 or total_seconds == 0:
        return "0'00\" /km"
    pace_seconds_per_km = total_seconds / distance_km
    pace_minutes = int(pace_seconds_per_km // 60)
    pace_seconds = int(pace_seconds_per_km % 60)
    return f"{pace_minutes}'{pace_seconds:02d}\" /km"

# ==============================================================================
# 資料爬取與處理
# ==============================================================================
def get_strava_data(url):
    """從 Strava 活動頁面爬取資料"""
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        response = requests.get(url, headers=headers)
        response.raise_for_status()

        soup = BeautifulSoup(response.text, 'html.parser')
        stats = soup.find_all('div', class_='Stat_statValue__lmw2H')
        
        if len(stats) < 3:
            return {'error': '無法在頁面上找到足夠的統計數據，請確認網址是否為公開活動。'}

        distance_str = stats[0].text.replace('km', '').strip()
        time_str = stats[1].text
        elevation_str = stats[2].text.replace('m', '').replace(',', '').strip()

        distance = float(distance_str)
        total_seconds = hms_to_seconds(time_str)
        elevation = int(elevation_str)
        avg_pace = calculate_pace(distance, total_seconds)

        return {
            'distance': f"{distance:.2f} km",
            'time': time_str,
            'elevation_gain': f"{elevation} m",
            'avg_pace': avg_pace,
            'source': 'Strava'
        }
    except Exception as e:
        return {'error': f'爬取 Strava 資料時發生錯誤: {e}'}

def get_garmin_data(url):
    """
    從 Garmin Connect 公開活動頁面的 meta 標籤直接爬取資料。
    這個方法最為穩定，因為 meta 標籤是設計給爬蟲讀取的。
    """
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36',
            'Accept-Language': 'en-US,en;q=0.9',
        }
        response = requests.get(url, headers=headers)
        response.raise_for_status()

        soup = BeautifulSoup(response.text, 'html.parser')
        
        # 找到 property 為 'og:description' 的 meta 標籤
        meta_tag = soup.find('meta', property='og:description')
        
        if not meta_tag or not meta_tag.get('content'):
            return {'error': '無法在 Garmin 頁面中找到活動的 meta 資訊。請確認活動為公開，或網址正確。'}
            
        content = meta_tag.get('content')
        # content 格式: "Distance 6.07 km | Time 36:20 | Pace 5:59 /km | Elevation 7 m"

        # 使用正規表示式從 content 字串中提取各項數據
        dist_match = re.search(r'Distance ([\d\.]+) km', content)
        time_match = re.search(r'Time ([\d:]+)', content)
        elev_match = re.search(r'Elevation ([\d]+) m', content)

        if not (dist_match and time_match and elev_match):
            return {'error': '從 meta 資訊中解析數據失敗，可能是 Garmin 更改了格式。'}

        distance_km = float(dist_match.group(1))
        time_str = time_match.group(1) # e.g., "36:20" or "1:05:33"
        elevation_m = int(elev_match.group(1))
        
        # 將時間字串轉換為總秒數
        total_seconds = hms_to_seconds(time_str)

        # 重新計算配速以確保格式統一
        avg_pace = calculate_pace(distance_km, total_seconds)
        
        # 格式化總時間為 hh:mm:ss
        td = timedelta(seconds=total_seconds)
        hours, remainder = divmod(td.seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        formatted_time = f"{int(hours):02}:{int(minutes):02}:{int(seconds):02}"

        return {
            'distance': f"{distance_km:.2f} km",
            'time': formatted_time,
            'elevation_gain': f"{elevation_m} m",
            'avg_pace': avg_pace,
            'source': 'Garmin Connect'
        }
        
    except Exception as e:
        return {'error': f'處理 Garmin 資料時發生未預期的錯誤: {e}'}

def get_user_data():
    if not GSPREAD_AVAILABLE: return None
    try:
        with gsheet_lock: records = worksheet.get_all_records()
        return pd.DataFrame(records) if records else pd.DataFrame()
    except Exception as e:
        print(f"讀取 Google Sheet 時發生錯誤：{e}", file=sys.stderr)
        return None

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
@app.route('/', methods=['GET'])
def index():
    if 'user_id_card' not in session:
        return redirect(url_for('login'))
    return render_template('index.html')

@app.route('/process_activity', methods=['POST'])
def process_activity():
    if 'user_id_card' not in session:
        flash('使用者未登入，請先登入。', 'warning')
        return redirect(url_for('login'))

    url = request.form.get('activity_url')
    url_type = request.form.get('url_type')
    
    if not url:
        flash('請輸入活動網址。', 'warning')
        return redirect(url_for('index'))

    activities = None
    if url_type == 'garmin':
        activities = get_garmin_data(url)
    elif url_type == 'strava':
        activities = get_strava_data(url)
    else:
        flash('請選擇有效的平台。', 'warning')
        return redirect(url_for('index'))
    
    if activities:
        # 直接選取第一個活動來產生證書
        activity = activities[0]
        
        # 更新後台紀錄
        update_user_log(session.get('user_id_card'), url)

        # 產生 PDF
        buffer = BytesIO()
        p = canvas.Canvas(buffer, pagesize=letter)
        width, height = letter

        p.setFont('NotoSansTC' if FONT_AVAILABLE else 'Helvetica', 36)
        p.drawCentredString(width / 2.0, height - 100, "完賽證明")

        p.setFont('NotoSansTC' if FONT_AVAILABLE else 'Helvetica', 18)
        user_name = session.get('user_name', '')
        user_number = session.get('user_number', '')

        p.drawCentredString(width / 2.0, height - 200, f"恭喜 {user_name} (編號: {user_number})")
        p.drawCentredString(width / 2.0, height - 250, "挑戰成功")
        p.drawCentredString(width / 2.0, height - 350, f"總距離：{activity.get('distance', 'N/A')}")
        p.drawCentredString(width / 2.0, height - 400, f"總時長：{activity.get('time', 'N/A')}")
        p.drawCentredString(width / 2.0, height - 450, f"最高海拔：{activity.get('elevation_gain', 'N/A')}")
        p.drawCentredString(width / 2.0, height - 500, f"平均配速：{activity.get('avg_pace', 'N/A')}")

        p.showPage()
        p.save()
        buffer.seek(0)
        
        safe_activity_name = quote(activity.get('activity_name', 'activity').replace(" ", "_"))
        filename = f"certificate_{user_name}_{safe_activity_name}.pdf"

        return Response(buffer, mimetype='application/pdf',
                        headers={'Content-Disposition': f'attachment;filename="{filename}"'})
    else:
        flash('抓取或解析資料失敗，請確認網址是否正確且頁面為公開。', 'danger')
        return redirect(url_for('index'))

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

