import os
import re
import sys
import requests
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, session, Response, flash
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
from io import BytesIO

# --- ReportLab ---
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
# TTFont is no longer needed as we are not using custom fonts
# from reportlab.pdfbase.ttfonts import TTFont 
from reportlab.lib.units import cm

# ==============================================================================
# Flask 應用程式設定
# ==============================================================================
app = Flask(__name__)
# 請務必更換為您自己的密鑰
app.secret_key = os.urandom(24) 

# --- 路徑設定 ---
# 取得 app.py 所在的目錄，確保路徑在 Render 上是正確的
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# The user's file is a CSV, even if it has .xlsx in the name.
USERS_CSV_PATH = os.path.join(BASE_DIR, 'users.xlsx')

# --- 字型註冊 ---
# 根據您的要求，已移除對 'NotoSansTC-Regular.ttf' 字型檔案的依賴與檢查。
# PDF 將使用 ReportLab 的內建字型，這可能導致中文無法正常顯示。

# ==============================================================================
# 資料與輔助函式
# ==============================================================================

# --- 載入使用者資料 ---
# 檢查使用者資料檔案是否存在
if os.path.exists(USERS_CSV_PATH):
    # 使用 read_csv 讀取您的檔案
    users_df = pd.read_csv(USERS_CSV_PATH)
    # 確保欄位是字串，避免比對時因型別錯誤而失敗
    users_df['id_card'] = users_df['id_card'].astype(str)
    users_df['phone'] = users_df['phone'].astype(str)
    USERS_DATA_AVAILABLE = True
else:
    print(f"FATAL ERROR: 使用者資料 '{os.path.basename(USERS_CSV_PATH)}' 未找到！登入功能將會失效。", file=sys.stderr)
    users_df = pd.DataFrame() # 建立空的 DataFrame 以免後續程式碼出錯
    USERS_DATA_AVAILABLE = False

def hms_to_seconds(t):
    """將 '1h 23m 45s' 或 '23m 45s' 格式的時間字串轉換為總秒數"""
    h, m, s = 0, 0, 0
    if 'h' in t:
        h = int(t.split('h')[0])
        t = t.split('h')[1].strip()
    if 'm' in t:
        m = int(t.split('m')[0])
        t = t.split('m')[1].strip()
    if 's' in t:
        s = int(t.split('s')[0])
    return h * 3600 + m * 60 + s

def calculate_pace(distance_km, total_seconds):
    """計算配速 (分鐘/公里)"""
    if distance_km == 0 or total_seconds == 0:
        return "0'00\""
    pace_seconds_per_km = total_seconds / distance_km
    pace_minutes = int(pace_seconds_per_km // 60)
    pace_seconds = int(pace_seconds_per_km % 60)
    return f"{pace_minutes}'{pace_seconds:02d}\""

# ==============================================================================
# 資料爬取函式
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
    """從 Garmin Connect API 獲取資料"""
    match = re.search(r'/(\d+)$', url.strip())
    if not match:
        return {'error': '無效的 Garmin Connect 網址，找不到活動 ID。'}
    activity_id = match.group(1)

    # !!! 重要提醒 !!!
    # 下方的 'authorization' token 具有時效性，會過期！
    # 您必須手動從瀏覽器登入 Garmin Connect 後，透過開發者工具 (F12) -> 網路(Network)
    # 找到一個對 `activity-service` 的請求，並複製其請求標頭中的 `authorization` 值來取代此處的 token。
    auth_token = 'Bearer eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCIsImtpZCI6ImRpLW9hdXRoLXNpZ25lci1wcm9kLTIwMjQtcTEifQ.eyJzY29wZSI6WyJBVFBfUkVBRCIsIkNPTU1VTklUWV9DT1VSU0VfUkVBRCIsIkNPTk5FQ1RfUkVBRCIsIkdPTEZfQVBJX1JFQUQiLCJJTlNJR0hUU19SRUFEIl0sImlzcyI6Imh0dHBzOi8vZGlhdXRoLmdhcm1pbi5jb20iLCJjbGllbnRfdHlwZSI6IlVOREVGSU5FRCIsImV4cCI6MTc1Mzk4NTQ4MywiaWF0IjoxNzUzOTgxODgzLCJqdGkiOiI4Yjc2MTVmMi1lZDg4LTQ5NDUtOWNjYy0yMzUxNTRlYjE2NTUiLCJjbGllbnRfaWQiOiJDT05ORUNUX1dFQiJ9.NIXLYfENfkSWXCXMkf1MGNLLwKd_kIlwtSyxU0IQZVahjwiNuvp74qE9a6nL2SP7KFkvRLYBAdTZFc3ohyQLSWXUzf6yUyt7UGf6HPvpvEHNfaKtBM-1ANQqlH-151W4iKMAPjFMHM8Uoi9emSXzmvnntAeylaGvP-SsYYhaP2r2TIS0oz7GVbQv8e1k6qtByxpHTY6OvtHNz_Xy5ijnR5El33x-UMkj5sK9tFCa1Y6suB_lrwOugprsCwPivyoBfw3_JGBgADTmnNrIREOg0uvESgrep9DArQXBQ_RDV2rPIYj5SFrgEEmj6dYLce0BNA4xVloDKbnt3iqjojrEoQ'
    
    api_url = f'https://connect.garmin.com/activity-service/activity/{activity_id}/splits'
    headers = {
        'accept': 'application/json, text/javascript, */*; q=0.01',
        'accept-language': 'zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7',
        'authorization': auth_token,
        'di-backend': 'connectapi.garmin.com',
        'nk': 'NT',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36',
        'x-app-ver': '5.15.1.1',
        'x-lang': 'zh-TW',
        'x-requested-with': 'XMLHttpRequest',
    }

    try:
        response = requests.get(api_url, headers=headers)
        response.raise_for_status()
        data = response.json()
        
        laps = data.get('lapDTOs', [])
        if not laps:
            return {'error': '從 API 回傳的資料中找不到輪次 (lap) 數據。'}

        total_distance_m = sum(lap.get('distance', 0) for lap in laps)
        total_distance_km = round(total_distance_m / 1000, 2)

        start_time_ms = laps[0].get('startTimeGMT')
        end_time_ms = laps[-1].get('startTimeGMT') + laps[-1].get('elapsedDuration')
        total_seconds = (end_time_ms - start_time_ms) / 1000
        
        td = timedelta(seconds=total_seconds)
        hours, remainder = divmod(td.seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        formatted_time = f"{hours:02}:{minutes:02}:{seconds:02}"
        
        max_elevation = max(lap.get('maxElevation', 0) for lap in laps)
        avg_pace = calculate_pace(total_distance_km, total_seconds)

        return {
            'distance': f"{total_distance_km:.2f} km",
            'time': formatted_time,
            'elevation_gain': f"{max_elevation:.0f} m",
            'avg_pace': avg_pace,
            'source': 'Garmin Connect'
        }

    except requests.exceptions.HTTPError as e:
        if e.response.status_code == 401:
            return {'error': f'Garmin API 請求錯誤 (401 Unauthorized): "authorization" token 已過期或無效，請更新程式碼中的 token。'}
        return {'error': f'Garmin API 請求失敗，狀態碼: {e.response.status_code}'}
    except Exception as e:
        return {'error': f'處理 Garmin 資料時發生錯誤: {e}'}

# ==============================================================================
# PDF 產生函式
# ==============================================================================

def create_certificate_pdf(user_name, user_number, activity_data):
    """使用 ReportLab 產生 PDF 證書"""
    buffer = BytesIO()
    p = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter

    # 注意：已將字型改為 ReportLab 內建的 'Helvetica'。
    # 這將無法正確顯示中文。如果需要顯示中文，必須重新引入中文字型檔。
    p.setFont('Helvetica', 36)
    p.drawCentredString(width / 2, height - 2*cm, "Activity Certificate") # Changed to English

    p.setFont('Helvetica', 16)
    p.drawCentredString(width / 2, height - 3*cm, "Certificate of Completion")
    
    p.setFont('Helvetica', 18)
    p.drawCentredString(width / 2, height - 5*cm, "This is to certify that") # Changed to English

    p.setFont('Helvetica', 30)
    p.setFillColorRGB(0.8, 0.1, 0.1)
    p.drawCentredString(width / 2, height - 6.5*cm, user_name)
    
    p.setFillColorRGB(0, 0, 0)
    p.setFont('Helvetica', 16)
    p.drawCentredString(width / 2, height - 7.5*cm, f"(No: {user_number})") # Changed to English
    
    p.setFont('Helvetica', 18)
    p.drawCentredString(width / 2, height - 8.5*cm, "has successfully completed the challenge with the following results:") # Changed to English

    # --- 繪製成績表格 ---
    p.setFont('Helvetica', 14)
    table_y_start = height - 11*cm
    col1_x = 5*cm
    col2_x = 8*cm
    col3_x = 13*cm
    col4_x = 16*cm
    row_height = 1*cm

    # 標題
    p.drawString(col1_x, table_y_start, "Distance")
    p.drawString(col3_x, table_y_start, "Time")
    p.drawString(col1_x, table_y_start - row_height, "Elevation")
    p.drawString(col3_x, table_y_start - row_height, "Avg. Pace")
    
    # 數據
    p.setFont('Helvetica', 16)
    p.drawString(col2_x, table_y_start, activity_data.get('distance', 'N/A'))
    p.drawString(col4_x, table_y_start, activity_data.get('time', 'N/A'))
    p.drawString(col2_x, table_y_start - row_height, activity_data.get('elevation_gain', 'N/A'))
    p.drawString(col4_x, table_y_start - row_height, activity_data.get('avg_pace', 'N/A'))

    # --- 頁尾 ---
    p.setFont('Helvetica', 10)
    p.drawCentredString(width / 2, 3*cm, f"Source: {activity_data.get('source', 'N/A')}")
    p.drawCentredString(width / 2, 2.5*cm, f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    p.showPage()
    p.save()
    
    buffer.seek(0)
    return buffer

# ==============================================================================
# Flask 路由
# ==============================================================================

@app.route('/')
def index():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    return render_template('index.html', user_name=session.get('user_name'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if not USERS_DATA_AVAILABLE:
        return render_template('login.html', error='系統錯誤：無法讀取使用者資料庫。')

    if request.method == 'POST':
        id_card = request.form['id_card']
        phone = request.form['phone']
        
        user = users_df[(users_df['id_card'] == id_card) & (users_df['phone'] == phone)]
        
        if not user.empty:
            session['user_id'] = id_card
            session['user_name'] = user.iloc[0]['name']
            session['user_number'] = str(user.iloc[0]['number'])
            return redirect(url_for('index'))
        else:
            return render_template('login.html', error='身分證號碼或電話號碼錯誤')
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/generate', methods=['POST'])
def generate_certificate():
    if 'user_id' not in session:
        return redirect(url_for('login'))

    source_type = request.form.get('source_type')
    activity_url = request.form.get('activity_url')

    if not activity_url:
        flash('請輸入活動網址。', 'danger')
        return redirect(url_for('index'))

    activity_data = None
    if source_type == 'strava':
        activity_data = get_strava_data(activity_url)
    elif source_type == 'garmin':
        activity_data = get_garmin_data(activity_url)
    else:
        flash('無效的平台類型。', 'danger')
        return redirect(url_for('index'))

    if activity_data and 'error' in activity_data:
        flash(f"處理失敗: {activity_data['error']}", 'danger')
        return redirect(url_for('index'))
    
    if activity_data:
        user_name = session.get('user_name')
        user_number = session.get('user_number')
        
        pdf_buffer = create_certificate_pdf(user_name, user_number, activity_data)
        
        return Response(pdf_buffer,
                        mimetype='application/pdf',
                        headers={'Content-Disposition': 'attachment;filename=certificate.pdf'})

    flash('無法獲取活動資料，請檢查您的網址或網路連線。', 'danger')
    return redirect(url_for('index'))

# ==============================================================================
# 應用程式啟動
# ==============================================================================
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
