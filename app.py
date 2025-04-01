from flask import Flask, render_template
import pandas as pd

app = Flask(__name__)

# 讀取 Excel 檔案
def read_excel():
    df = pd.read_excel("data.xlsx")  # 確保 data.xlsx 在同個資料夾
    return df.to_dict(orient="records")  # 轉成字典格式，方便在 HTML 顯示

@app.route("/")
def home():
    data = read_excel()  # 讀取 Excel 數據
    return render_template("index.html", data=data)

if __name__ == "__main__":
    app.run(debug=True)
