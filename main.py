from flask import Flask, render_template, Response
import io
import xlwt
import os

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/table/download')
def download_report():
    result = [
        ["id", "data_1", "data_2"],
        ["id", "data_1", "data_2"],
        ["id", "data_1", "data_2"],
        ["id", "data_1", "data_2"],
    ]
    output = io.BytesIO()
    wb = xlwt.Workbook()
    ws = wb.add_sheet('contenido')

    ws.write(0,0,"Title_1")
    ws.write(0,1,"Title_2")
    ws.write(0,2,"Title_3")

    idx = 0
    for row in result:
        ws.write(idx+1, 0, f"row {idx+1}")
        ws.write(idx+1, 1, f"row {idx+1}")
        ws.write(idx+1, 2, f"row {idx+1}")
        idx += 1

    wb.save(output)
    output.seek(0)

    return Response(output, mimetype="application/ms-excel", headers={"Content-Disposition":"attachment;filename=data.xls"})

if __name__ == '__main__':
    app.run(debug=True, port=os.getenv("PORT", default=5000))
