import pandas as pd
from flask import Flask, render_template, request

app = Flask(__name__)

# 主頁資料：讀取第 14 列為欄位名稱，最多 212 筆
def load_main_data():
    df = pd.read_excel('data.xlsx', header=13, nrows=212)
    df.columns = df.columns.str.replace('\n', '', regex=False)  # 移除欄位中的換行符號
    df = df.fillna('')  # 將 NaN / NaT 替換為空白
    return df

# 個人資料表：讀取 A1~J6（無標題列）
def load_person_data(sheet_name):
    df = pd.read_excel('data.xlsx', sheet_name=sheet_name, usecols="A:J", nrows=6, header=None)
    df = df.fillna('')
    return df.values.tolist()

@app.route('/')
def index():
    df = load_main_data()

    # 只顯示特定欄位
    display_columns = ['門市編號', '門市名稱', '鄉鎮市區', 'PMQ2檢核', 'EDC檢核', '發票機檢核', '數量']
    df = df[[col for col in display_columns if col in df.columns]]

    # 關鍵字搜尋（模糊比對）
    keyword = request.args.get('keyword', '')
    if keyword:
        df = df[df.apply(lambda row: row.astype(str).str.contains(keyword, case=False).any(), axis=1)]

    return render_template('index.html', tables=df.to_dict(orient='records'), keyword=keyword, personal=False)

@app.route('/<name>')
def person_page(name):
    try:
        tables = load_person_data(name)
        return render_template('index.html', tables=tables, keyword=name, personal=True)
    except Exception as e:
        return f"⚠️ 找不到工作表「{name}」或讀取錯誤：{str(e)}", 404

if __name__ == '__main__':
    app.run(debug=True)
