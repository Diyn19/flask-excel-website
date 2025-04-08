<<<<<<< HEAD
import pandas as pd
from flask import Flask, render_template, request

app = Flask(__name__)

@app.route('/')
def index():
    # 讀取 Excel 檔案
    df = pd.read_excel('data.xlsx', header=13, nrows=212)

    # 清除欄位名稱中的換行符號
    df.columns = df.columns.str.replace('\n', '', regex=False)

    # 篩選所需欄位
    df = df[['門市編號', '門市名稱', '鄉鎮市區', 'PMQ2檢核', 'EDC檢核', '發票機檢核', '數量']]

    # 處理關鍵字搜尋
    keyword = request.args.get('keyword', '')
    if keyword:
        df = df[df.apply(lambda row: row.astype(str).str.contains(keyword, case=False).any(), axis=1)]

    return render_template('index.html', tables=df.to_dict(orient='records'), keyword=keyword)

if __name__ == '__main__':
    app.run(debug=True)
=======
import pandas as pd
from flask import Flask, render_template, request

app = Flask(__name__)

@app.route('/')
def index():
    # 讀取 Excel 檔案
    df = pd.read_excel('data.xlsx', header=13, nrows=212)

    # 清除欄位名稱中的換行符號
    df.columns = df.columns.str.replace('\n', '', regex=False)

    # 篩選所需欄位
    df = df[['門市編號', '門市名稱', '鄉鎮市區', 'PMQ2檢核', 'EDC檢核', '發票機檢核', '數量']]

    # 處理關鍵字搜尋
    keyword = request.args.get('keyword', '')
    if keyword:
        df = df[df.apply(lambda row: row.astype(str).str.contains(keyword, case=False).any(), axis=1)]

    return render_template('index.html', tables=df.to_dict(orient='records'), keyword=keyword)

if __name__ == '__main__':
    app.run(debug=True)
>>>>>>> de8697e5cd127917afb2586f9bab1730c1f4d5a1
