import pandas as pd
import requests
from bs4 import BeautifulSoup
from flask import Flask, render_template, request, redirect, url_for, session, Response
import pdfkit
import datetime

# --- Flask App Initialization ---
app = Flask(__name__)
app.secret_key = 'your-very-secret-key' 

# --- Helper Functions ---
def load_user_data():
    """從 Excel 檔案讀取使用者資料"""
    try:
        df = pd.read_excel('users.xlsx', dtype=str)
        return df.to_dict('records')
    except FileNotFoundError:
        print("錯誤：找不到 users.xlsx 檔案。")
        return []

def scrape_garmin_data(url):
    """
    嘗試爬取 Garmin Connect 活動頁面的資料。
    """
    print(f"正在嘗試爬取網址: {url}")
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
        'Accept-Language': 'en-US,en;q=0.9,zh-TW;q=0.8',
        'Connection': 'keep-alive',
    }
    
    try:
        response = requests.get(url, headers=headers, timeout=15)
        
        # --- 新增的除錯程式碼 ---
        print("\n--- Requests 請求結果 ---")
        # 我們只印出前 500 個字元，避免洗版終端機
        print(response.text[:500] + "...")
        print("--- 請求結果結束 ---\n")
        
        response.raise_for_status() 
        soup = BeautifulSoup(response.text, 'html.parser')
        
        data_fields = soup.select('.DataBlock_dataField__t4-ai')
        
        print(f"找到了 {len(data_fields)} 個符合 class 的欄位。")

        if len(data_fields) >= 4:
            data = {
                'distance': data_fields[0].text.strip(),
                'time': data_fields[1].text.strip(),
                'avg_pace': data_fields[2].text.strip(),
                'elevation_gain': data_fields[3].text.strip(),
                'map_image_url': 'https://placehold.co/800x400/dddddd/333333?text=Map+Screenshot+Placeholder'
            }
            print("成功從 HTML 爬取到資料！")
            return data
        else:
            print("警告：未能找到足夠的資料欄位 (需要 4 個)。將啟用模擬資料。")
            return None

    except Exception as e:
        print(f"爬取時發生錯誤: {e}")
        return None

def get_mock_data():
    """如果爬蟲失敗，回傳一組模擬資料"""
    print("爬蟲失敗或資料不足，啟用模擬資料。")
    return {
        'distance': '10.52 公里',
        'time': '58:33',
        'avg_pace': "5'34\" /公里",
        'elevation_gain': '123 公尺',
        'map_image_url': 'https://placehold.co/800x400/dddddd/333333?text=Map+Screenshot+Placeholder'
    }

# --- Flask Routes ---
@app.route('/login', methods=['GET', 'POST'])
def login():
    """登入頁面"""
    if request.method == 'POST':
        id_card = request.form['id_card']
        phone = request.form['phone']
        
        users = load_user_data()
        user_found = next((user for user in users if user['id_card'] == id_card and user['phone'] == phone), None)
        
        if user_found:
            session['logged_in'] = True
            session['user_name'] = user_found['name']
            session['user_number'] = user_found['user_number']
            return redirect(url_for('index'))
        else:
            return render_template('login.html', error='身分證或電話號碼錯誤')
            
    return render_template('login.html')

@app.route('/')
def index():
    """提交網址的主頁面"""
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    return render_template('index.html', user_name=session.get('user_name'))

@app.route('/generate', methods=['POST'])
def generate_certificate():
    """生成證書 PDF 的核心功能"""
    if not session.get('logged_in'):
        return redirect(url_for('login'))
        
    garmin_url = request.form['garmin_url']
    
    activity_data = scrape_garmin_data(garmin_url)
    
    if activity_data is None:
        activity_data = get_mock_data()
    
    certificate_data = {
        'user_name': session.get('user_name'),
        'user_number': session.get('user_number'),
        'generation_date': datetime.date.today().strftime('%Y-%m-%d'),
        **activity_data
    }
    
    html_out = render_template('certificate.html', **certificate_data)
    
    try:
        # !! 重要 !! 請將此路徑修改為您電腦上 wkhtmltopdf.exe 的實際路徑
        path_wkhtmltopdf = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
        
        config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)
        
        pdf_file = pdfkit.from_string(html_out, False, configuration=config, options={"enable-local-file-access": ""})
        
        return Response(pdf_file,
                        mimetype='application/pdf',
                        headers={'Content-Disposition': 'attachment;filename=certificate.pdf'})

    except FileNotFoundError:
        print(f"錯誤：找不到 wkhtmltopdf.exe。請檢查 app.py 中的 `path_wkhtmltopdf` 路徑設定是否正確。")
        return "PDF 生成失敗：找不到 wkhtmltopdf.exe。請檢查伺服器日誌與設定。", 500

@app.route('/logout')
def logout():
    """登出"""
    session.clear()
    return redirect(url_for('login'))

# --- Main Execution ---
if __name__ == '__main__':
    app.run(debug=True, port=5000)