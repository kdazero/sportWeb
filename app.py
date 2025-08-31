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
    # (連線設定代碼維持不變)
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
def get_user_data():
    """從 Google Sheets 讀取使用者資料並轉換為 DataFrame。"""
    if not GSPREAD_AVAILABLE:
        return None
    try:
        with gsheet_lock:
            records = worksheet.get_all_records()
        if not records:
            print("警告：從 Google Sheet 讀取到的資料為空。", file=sys.stderr)
            return pd.DataFrame()
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
            # 將傳入的 user_id_card 也清理一次，確保查找時格式一致
            cleaned_id_card = str(user_id_card).strip().upper()
            
            # 讀取 B 欄所有資料並清理
            id_card_column = worksheet.col_values(2) # B 欄是第 2 欄
            cleaned_column = [str(val).strip().upper() for val in id_card_column]

            # 找到匹配的列
            match_row = -1
            try:
                # +1 是因為 list index 從 0 開始，但 sheet row 從 1 開始
                match_row = cleaned_column.index(cleaned_id_card) + 1
            except ValueError:
                # 找不到匹配
                pass

            if match_row == -1:
                print(f"紀錄失敗：在 Google Sheet 中找不到使用者 {cleaned_id_card}。", file=sys.stderr)
                return

            # 找到 'last_print' 所在的欄
            headers = worksheet.row_values(1)
            try:
                col_index = headers.index('last_print') + 1
            except ValueError:
                col_index = len(headers) + 1
                worksheet.update_cell(1, col_index, 'last_print')

            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            worksheet.update_cell(match_row, col_index, timestamp)
            print(f"紀錄成功：使用者 {cleaned_id_card} 的證書產生時間已更新。")

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
        print("\n--- 登入診斷開始 ---")
        
        id_card_input = request.form.get('id_card', '').strip().upper()
        phone_input = request.form.get('phone', '').strip()
        print(f"1. 使用者輸入 (已清理): id_card='{id_card_input}', phone='{phone_input}'")
        
        users_df = get_user_data()
        if users_df is None:
            flash('伺服器錯誤，無法讀取使用者資料。', 'danger')
            print("--- 診斷結束：無法讀取 DataFrame ---")
            return render_template('login.html')

        if users_df.empty:
            flash('資料庫中沒有使用者資料。', 'danger')
            print("--- 診斷結束：DataFrame 為空 ---")
            return render_template('login.html')

        print(f"2. 讀取到的欄位名稱: {users_df.columns.tolist()}")
        print("3. Google Sheet 前 2 筆原始資料:")
        print(users_df.head(2).to_string())

        # 檢查必要欄位是否存在
        required_columns = ['id_card', 'phone']
        for col in required_columns:
            if col not in users_df.columns:
                flash('資料庫欄位設定錯誤，請聯繫管理員。', 'danger')
                print(f"--- 診斷結束：缺少 '{col}' 欄位 ---")
                return render_template('login.html')

        df_id_card_col = users_df['id_card'].astype(str).str.strip().str.upper()
        df_phone_col = users_df['phone'].astype(str).str.strip()

        user = users_df[(df_id_card_col == id_card_input) & (df_phone_col == phone_input)]

        if not user.empty:
            print("4. 比對結果：成功！")
            print("--- 診斷結束 ---")
            user_info = user.iloc[0]
            session['user_id_card'] = user_info['id_card']
            session['user_name'] = user_info['name']
            session['user_number'] = user_info['user_number']
            
            # 檢查 'garmin_url' 是否存在
            if 'garmin_url' in user_info and pd.notna(user_info['garmin_url']):
                 session['garmin_url'] = user_info['garmin_url']
            else:
                 session['garmin_url'] = None # 或給一個預設值

            return redirect(url_for('fetch_activities'))
        else:
            print("4. 比對結果：失敗。")
            print("--- 診斷結束 ---")
            flash('身分證號或手機號碼錯誤。', 'danger')
            return render_template('login.html')

    return render_template('login.html')

@app.route('/fetch_activities')
def fetch_activities():
    if 'user_id_card' not in session:
        return redirect(url_for('login'))

    garmin_url = session.get('garmin_url')
    if not garmin_url: # 修正檢查方式
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

    update_user_log_timestamp(session.get('user_id_card'))

    activity_name = activity.get('activityName', 'N/A')
    distance_meters = activity.get('distance', 0)
    distance_km = round(distance_meters / 1000, 2) if distance_meters else 0
    moving_time_seconds = activity.get('movingTime', 0)
    moving_time_hms = seconds_to_hms(moving_time_seconds)
    
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

