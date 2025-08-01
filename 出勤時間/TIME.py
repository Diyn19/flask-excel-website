from flask import Flask, render_template
import pandas as pd
import os
import io
import base64
import matplotlib
matplotlib.use('Agg')  # 非 GUI 模式
import matplotlib.pyplot as plt
from matplotlib import rcParams
import matplotlib.font_manager as fm

# 設定中文字體
font_path = 'C:/Windows/Fonts/msjh.ttc'  # 微軟正黑體
font_prop = fm.FontProperties(fname=font_path)
rcParams['font.family'] = font_prop.get_name()

app = Flask(__name__)

@app.route('/attendance')
def attendance():
    excel_path = 'data.xlsx'

    # 讀取版本號
    try:
        version_df = pd.read_excel(excel_path, sheet_name='首頁', header=None, usecols="G", nrows=1)
        version = version_df.iloc[0, 0]
    except:
        version = "無法讀取版本號"

    # 讀取摘要與明細資料（保留你原本的）
    df_summary = pd.read_excel(excel_path, sheet_name='出勤時間', usecols="A:E", nrows=2)
    detail_1 = pd.read_excel(excel_path, sheet_name='出勤時間', usecols="A:Q", skiprows=3, nrows=3)
    detail_2 = pd.read_excel(excel_path, sheet_name='出勤時間', usecols="A:Q", skiprows=7, nrows=3)
    detail_3 = pd.read_excel(excel_path, sheet_name='出勤時間', usecols="A:Q", skiprows=11, nrows=3)

    # 讀取曲線圖資料
    # 時間軸：B12:P12 → index=11，col=1~15（0-based）
    # 姓名：A13:A15 → index=12~14，col=0
    # 出勤次數：B13:P15 → index=12~14，col=1~15

    df_chart = pd.read_excel(excel_path, sheet_name='出勤時間', header=None)

    # 取得時間軸字串
    x = df_chart.iloc[11, 1:16].tolist()  # B12:P12 (Excel 1-based, DataFrame 0-based)
    # 確認 x 是什麼格式，如果是數字（Excel 時間序列），轉成時間格式字串
    if all(isinstance(v, (int, float)) for v in x):
        # Excel 日期是從 1900-01-01 起算，時間是一天的小數部分
        # 假設這裡是時間欄，直接用 pd.Timedelta 轉小時分鐘
        x = [(pd.Timestamp("1900-01-01") + pd.Timedelta(hours=hour)).strftime("%H:%M") for hour in range(8, 23)]
    else:
        # 否則直接轉成字串（保險用）
        x = [str(v) for v in x]

    # 取得人員姓名
    names = df_chart.iloc[12:15, 0].tolist()

    # 取得次數資料，轉成 list of list，3 行×15 欄
    y_data = df_chart.iloc[12:15, 1:16].values.tolist()

    # 畫圖
    fig, ax = plt.subplots(figsize=(10, 5))
    for i, y in enumerate(y_data):
        ax.plot(x, y, marker='o', label=names[i])

    ax.set_xlabel("時間")
    ax.set_ylabel("出勤次數")
    ax.set_title("出勤時間曲線圖")
    ax.legend()
    plt.xticks(rotation=45)
    plt.tight_layout()

    # 圖片轉 base64
    img = io.BytesIO()
    plt.savefig(img, format='png')
    img.seek(0)
    plot_url = base64.b64encode(img.read()).decode('utf-8')
    plt.close()

    return render_template(
        'attendance.html',
        summary_table=df_summary.to_html(index=False, classes='dataframe'),
        detail_table_1=detail_1.to_html(index=False, classes='dataframe'),
        detail_table_2=detail_2.to_html(index=False, classes='dataframe'),
        detail_table_3=detail_3.to_html(index=False, classes='dataframe'),
        version=version,
        plot_url=plot_url
    )

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
