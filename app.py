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

# --- ReportLab ---
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import cm

# ==============================================================================
# Flask 應用程式設定
# ==============================================================================
app = Flask(__name__)
# 請務必更換為您自己的密鑰
app.secret_key = os.urandom(24) 

# --- 路徑設定 ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FONT_PATH = os.path.join(BASE_DIR, 'NotoSansTC-Regular.ttf')
USERS_EXCEL_PATH = os.path.join(BASE_DIR, 'users.xlsx')

# --- 字型註冊 ---
if os.path.exists(FONT_PATH):
    pdfmetrics.registerFont(TTFont('NotoSansTC', FONT_PATH))
    FONT_AVAILABLE = True
else:
    print("警告：中文字型 'NotoSansTC-Regular.ttf' 未找到。PDF 中的中文可能無法正常顯示。", file=sys.stderr)
    FONT_AVAILABLE = False

# ==============================================================================
# 資料與輔助函式
# ==============================================================================

# --- 載入使用者資料 ---
if os.path.exists(USERS_EXCEL_PATH):
    try:
        users_df = pd.read_excel(USERS_EXCEL_PATH, dtype=str)
        users_df.columns = users_df.columns.str.strip()
        users_df['id_card'] = users_df['id_card'].str.strip()
        users_df['phone'] = users_df['phone'].str.strip()
        USERS_DATA_AVAILABLE = True
    except Exception as e:
        print(f"FATAL ERROR: 讀取使用者資料時發生錯誤: {e}", file=sys.stderr)
        users_df = pd.DataFrame()
        USERS_DATA_AVAILABLE = False
else:
    print(f"FATAL ERROR: 使用者資料 '{os.path.basename(USERS_EXCEL_PATH)}' 未找到！登入功能將會失效。", file=sys.stderr)
    users_df = pd.DataFrame()
    USERS_DATA_AVAILABLE = False


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

# ==============================================================================
# PDF 產生函式
# ==============================================================================

def create_certificate_pdf(user_name, user_number, activity_data):
    """使用 ReportLab 產生 PDF 證書"""
    if not FONT_AVAILABLE:
        buffer = BytesIO()
        p = canvas.Canvas(buffer, pagesize=letter)
        p.setFont('Helvetica', 12)
        p.drawCentredString(letter[0]/2, letter[1]/2, "無法產生證書，因為缺少中文字型檔案。")
        p.showPage()
        p.save()
        buffer.seek(0)
        return buffer

    buffer = BytesIO()
    p = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter

    p.setFont('NotoSansTC', 36)
    p.drawCentredString(width / 2, height - 2*cm, "活動完賽證明")
    p.setFont('NotoSansTC', 16)
    p.drawCentredString(width / 2, height - 3*cm, "Certificate of Completion")
    p.setFont('NotoSansTC', 18)
    p.drawCentredString(width / 2, height - 5*cm, "茲證明參賽者")
    p.setFont('NotoSansTC', 30)
    p.setFillColorRGB(0.8, 0.1, 0.1)
    p.drawCentredString(width / 2, height - 6.5*cm, user_name)
    p.setFillColorRGB(0, 0, 0)
    p.setFont('NotoSansTC', 16)
    p.drawCentredString(width / 2, height - 7.5*cm, f"(編號: {user_number})")
    p.setFont('NotoSansTC', 18)
    p.drawCentredString(width / 2, height - 8.5*cm, "已成功完成本次挑戰，成績如下：")

    p.setFont('NotoSansTC', 14)
    table_y_start = height - 11*cm
    col1_x = 5*cm
    col2_x = 8*cm
    col3_x = 13*cm
    col4_x = 16*cm
    row_height = 1*cm

    p.drawString(col1_x, table_y_start, "總距離")
    p.drawString(col3_x, table_y_start, "總時長")
    p.drawString(col1_x, table_y_start - row_height, "最高海拔")
    p.drawString(col3_x, table_y_start - row_height, "平均配速")
    
    p.setFont('NotoSansTC', 16)
    p.drawString(col2_x, table_y_start, activity_data.get('distance', 'N/A'))
    p.drawString(col4_x, table_y_start, activity_data.get('time', 'N/A'))
    p.drawString(col2_x, table_y_start - row_height, activity_data.get('elevation_gain', 'N/A'))
    p.drawString(col4_x, table_y_start - row_height, activity_data.get('avg_pace', 'N/A'))

    p.setFont('NotoSansTC', 10)
    p.drawCentredString(width / 2, 3*cm, f"資料來源: {activity_data.get('source', 'N/A')}")
    p.drawCentredString(width / 2, 2.5*cm, f"證書產生時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

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
        flash('系統錯誤：無法讀取使用者資料庫，請聯繫管理員。', 'danger')
        return render_template('login.html', error='系統設定錯誤')

    if request.method == 'POST':
        id_card = request.form['id_card'].strip()
        phone = request.form['phone'].strip()
        
        user = users_df[(users_df['id_card'] == id_card) & (users_df['phone'] == phone)]
        
        if not user.empty:
            session['user_id'] = id_card
            session['user_name'] = user.iloc[0]['name']
            session['user_number'] = str(user.iloc[0]['user_number'])
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
        
        # 建立檔名，並進行 URL 編碼以支援中文
        filename = f"{user_name}_參加證明.pdf"
        encoded_filename = quote(filename)
        
        return Response(pdf_buffer,
                        mimetype='application/pdf',
                        headers={
                            'Content-Disposition': f"attachment; filename*=UTF-8''{encoded_filename}"
                        })

    flash('無法獲取活動資料，請檢查您的網址或網路連線。', 'danger')
    return redirect(url_for('index'))

# ==============================================================================
# 應用程式啟動
# ==============================================================================
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
