import pandas as pd
from flask import Flask, render_template, request

app = Flask(__name__)

# 處理資料並清除 NaT/nan
def clean_df(df):
    df.columns = df.columns.str.replace('\n', '', regex=False)
    df = df.fillna('')
    return df

@app.route('/')
def index():
    df = pd.read_excel('data.xlsx', sheet_name=0, header=13, nrows=212, usecols="A:J")
    df = clean_df(df)

    # 篩選顯示欄位
    df = df[['門市編號', '門市名稱', '鄉鎮市區', 'PMQ2檢核', 'EDC檢核', '發票機檢核', '數量']]

    # 搜尋
    keyword = request.args.get('keyword', '')
    if keyword:
        df = df[df.apply(lambda row: row.astype(str).str.contains(keyword, case=False).any(), axis=1)]

    return render_template('index.html', tables=df.to_dict(orient='records'), keyword=keyword, personal_page=False)

@app.route('/<name>')
def personal(name):
    sheet_map = {
        '吳宗鴻': '吳宗鴻',
        '湯家瑋': '湯家瑋',
        '狄澤洋': '狄澤洋'
    }

    sheet_name = sheet_map.get(name, None)
    if not sheet_name:
        return f"找不到{name}的分頁", 404

    df = pd.read_excel('data.xlsx', sheet_name=sheet_name, usecols="A:J", nrows=6)
    df = clean_df(df)

    return render_template('index.html', tables=df.to_dict(orient='records'), keyword=name, personal_page=True)

if __name__ == '__main__':
    app.run(debug=True)
