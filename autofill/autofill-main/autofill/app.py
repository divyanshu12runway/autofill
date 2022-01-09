from io import BytesIO
from flask import Flask, redirect, url_for, render_template, request, send_file
import os
import pythoncom
import win32com.client


app = Flask(__name__)


@app.route("/", methods=["POST", "GET"])
def login():
    if request.method == "POST":
        user = request.form["nm"]
        pythoncom.CoInitialize()
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = True
        excel.EnableEvents = True
        # Disable protected mode
        excel.DisplayAlerts = False
        excel_file = os.path.abspath('PHE FINAL.xlsx')
        # Disable macros prompts
        excel.AskToUpdateLinks = True
        wb = excel.Workbooks.Open(excel_file)
        ws = wb.Worksheets[1]

        #filling values start
        ws.Cells(7,3).value = user
        ws.Cells(3,3).value = request.form["v33"]
        ws.Cells(4,3).value = request.form["v43"]
        ws.Cells(5,3).value = request.form["v53"]
        ws.Cells(6,3).value = request.form["v63"]
        ws.Cells(8,3).value = request.form["v83"]
        ws.Cells(9,3).value = request.form["v93"]
        ws.Cells(10,3).value = request.form["v103"]
        ws.Cells(11,3).value = request.form["v113"]
        ws.Cells(12,3).value = request.form["v123"]
        ws.Cells(13,3).value = request.form["v133"]
        ws.Cells(14,3).value = request.form["v143"]
        ws.Cells(15,3).value = request.form["v153"]
        ws.Cells(16,3).value = request.form["v163"]
        ws.Cells(17,3).value = request.form["v173"]
        ws.Cells(18,3).value = request.form["v183"]
        ws.Cells(19,3).value = request.form["v193"]
        ws.Cells(23,3).value = request.form["v233"]
        ws.Cells(23,4).value = request.form["v234"]
        ws.Cells(23,5).value = request.form["v235"]
        ws.Cells(21,3).value = request.form["v213"]
        ws.Cells(21,5).value = request.form["v215"]
        ws.Cells(21,7).value = request.form["v217"]
        ws.Cells(26,2).value = request.form["v262"]
        ws.Cells(27,2).value = request.form["v272"]
        ws.Cells(28,2).value = request.form["v282"]
        ws.Cells(29,2).value = request.form["v292"]
        ws.Cells(30,2).value = request.form["v302"]
        ws.Cells(26,3).value = request.form["v263"]
        ws.Cells(27,3).value = request.form["v273"]
        ws.Cells(28,3).value = request.form["v283"]
        ws.Cells(29,3).value = request.form["v293"]
        ws.Cells(30,3).value = request.form["v303"]
        ws.Cells(26,4).value = request.form["v264"]
        ws.Cells(27,4).value = request.form["v274"]
        ws.Cells(28,4).value = request.form["v284"]
        ws.Cells(29,4).value = request.form["v294"]
        ws.Cells(30,4).value = request.form["v304"]
        ws.Cells(26,5).value = request.form["v265"]
        ws.Cells(27,5).value = request.form["v275"]
        ws.Cells(28,5).value = request.form["v285"]
        ws.Cells(29,5).value = request.form["v295"]
        ws.Cells(30,5).value = request.form["v305"]
        ws.Cells(26,6).value = request.form["v266"]
        ws.Cells(27,6).value = request.form["v276"]
        ws.Cells(28,6).value = request.form["v286"]
        ws.Cells(29,6).value = request.form["v296"]
        ws.Cells(30,6).value = request.form["v306"]
        ws.Cells(26,7).value = request.form["v267"]
        ws.Cells(27,7).value = request.form["v277"]
        ws.Cells(28,7).value = request.form["v287"]
        ws.Cells(29,7).value = request.form["v297"]
        ws.Cells(30,7).value = request.form["v307"]
        #filling values end
        wb.SaveAs(os.path.abspath('PHE FINAL.xlsx'))
        ws = wb.Worksheets[0]
        ws.PageSetup.Zoom = False
        ws.PageSetup.FitToPagesTall = False
        ws.PageSetup.FitToPagesWide = 1
        ws.PageSetup.PrintArea = 'A1:H286'
        # ws_index_list = [0]
        # wb.WorkSheets(ws_index_list).Select()
        pdf_file = os.path.abspath('PHE FINAL.pdf')
        wb.ActiveSheet.ExportAsFixedFormat(0, pdf_file)
        excel.Application.Quit()
        return redirect(url_for('uploaded_file', filename=pdf_file))
    return render_template("form.html")

@app.route('/<filename>')
def uploaded_file(filename):
    with open(filename, 'rb') as f:
        file_io = BytesIO(f.read())
    return send_file(file_io, download_name=os.path.basename(filename), as_attachment=True)
# @app.route("/<usr>")
# def user(usr):
#     return f"<h1>{usr}</h1>"

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)