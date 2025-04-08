import pandas as pd
from flask import Flask, render_template, request

app = Flask(__name__)

# 讀取與整理 Excel 資料
def load_full_data():
    df = pd.read_excel('data.xlsx', header=13, nrows=212)
    df.columns = df.columns.str.replace('\n', '', regex=False)
    return df

# 讀取並篩選 A1:J6 區塊資料
def filter_columns_and_rows(df):
    # 篩選僅 A 到 J 欄
    df = df.iloc[:, 0:10]
    # 篩選前 6 行（A1~J6）
    df = df.head(6)
    return df

@app.route('/')
def index():
    df = load_full_data()

    # 篩選首頁要顯示的欄位
    df = df[['門市編號', '門市名稱', '鄉鎮市區', 'PMQ2檢核', 'EDC檢核', '發票機檢核', '數量']]

    # 處理關鍵字搜尋
    keyword = request.args.get('keyword', '')
    if keyword:
        df = df[df.apply(lambda row: row.astype(str).str.contains(keyword, case=False).any(), axis=1)]

    return render_template('index.html', tables=df.to_dict(orient='records'), keyword=keyword)

# 個人資料分頁
@app.route('/<name>')
def person_page(name):
    df = load_full_data()

    # 篩選出有該名字的列
    df_filtered = df[df.apply(lambda row: row.astype(str).str.contains(name, case=False).any(), axis=1)]
    
    # 過濾A1:J6的範圍並且只顯示A-J欄
    df_filtered = filter_columns_and_rows(df_filtered)

    return render_template('index.html', tables=df_filtered.to_dict(orient='records'), keyword=name)

if __name__ == '__main__':
    app.run(debug=True)
