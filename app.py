import pandas as pd
from flask import Flask, render_template, request

app = Flask(__name__)

def load_person_sheet(sheet_name):
    df = pd.read_excel('data.xlsx', sheet_name=sheet_name, header=None)  # 不設定 header，保留所有資料
    df = df.iloc[:6, :10]  # A1~J6
    df.fillna('', inplace=True)
    return df.values.tolist()  # 轉為 list of lists，方便前端用 for loop 渲染

@app.route('/')
def index():
    df = pd.read_excel('data.xlsx', header=13, nrows=212)
    df.columns = df.columns.str.replace('\n', '', regex=False)
    df = df[['門市編號', '門市名稱', '鄉鎮市區', 'PMQ2檢核', 'EDC檢核', '發票機檢核', '數量']]
    keyword = request.args.get('keyword', '')
    if keyword:
        df = df[df.apply(lambda row: row.astype(str).str.contains(keyword, case=False).any(), axis=1)]
    return render_template('index.html', tables=df.to_dict(orient='records'), keyword=keyword)

@app.route('/<name>')
def person(name):
    df_data = load_person_sheet(name)
    return render_template('person.html', person_name=name, table_data=df_data)

if __name__ == '__main__':
    app.run(debug=True)
