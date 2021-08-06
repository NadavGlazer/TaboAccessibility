from flask.helpers import send_from_directory
from flask import Flask, render_template, request, send_file, redirect, url_for
from pdfkit.api import from_file
from werkzeug.utils import secure_filename
import time
from datetime import date
import _thread
import platform
import pdfkit
import os
from PIL import Image
import os


app = Flask(__name__)
html_2_pic_template = """<html><head>
  <meta charset="utf-8">
  <title>My Page Title</title>
  <style>
   body {
        height: 842px;
        width: 595px;
        /* to centre page on screen*/
        margin-left: auto;
        margin-right: auto;
    }
    .PDiv {
      margin: auto;
      border: 3px solid black;
      padding: 5px;
      width: 300px;
      height: 150px;
      text-align: right;
    }

    .firstPic {
      display: table-cell;
      width:80%;      
      
      -webkit-transform: rotate(90deg);
      -moz-transform: rotate(90deg);
      -o-transform: rotate(90deg);
      -ms-transform: rotate(90deg);
      transform: rotate(90deg);
    }

    .thirdPic {
      margin:auto; 
      display: flex;
      justify-content: center;
      width: 100%;
      height: 20%;
      margin-top:10px;  
    }

    .picDiv {
      width: 100%;
      display: flex;
      justify-content: center;
      margin-top:150px;
      height: 30%;

    }
  </style>
</head>

<body  width="2480" heigth="3508">
  <div>
    <div>
      <img src="C:/Users/Nadav1/TaboProject/TaboAccessibility/header.jfif" width="1000">
      </div>
    <div class="PDiv">
      <h2 style="white-space: pre-line;">text_from_form</h>
    </div>
    <div class="picDiv">
      <div class="firstPic">
          <img style="width:550px;height:400px;object-fit: fill;margin-bottom:10px;" src="pic_one">
          <img style="width:550px;height:400px;object-fit: fill;" src="pic_two">
      </div>     
    </div>
    <div class="thirdPic">
        <img style="width:800px;height:100px;object-fit: fill;" src="pic_three">   
      </div>  

    <div style="float:right; margin-bottom:0%;">
        <h2>page_number</h2>
    </div>
  </div>

</body>
</html>"""
UPLOAD_FOLDER = "pdfpictures/static/uploads/"
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["MAX_CONTENT_PATH"] = 10000
app.config["TIME_OUT"] = 0


@app.route("/")
def index():
    print(request.cookies)
    print(request.remote_addr)
    return render_template("index.html")


@app.route("/Star", methods=["GET", "POST"])
def trying():
    if request.method == "POST":
        current_time = (
            str(date.today().strftime("%d/%m/%Y"))
            + "_"
            + str(time.strftime("%H:%M:%S", time.localtime()))
        )
        current_time = current_time.replace("/", "_")
        current_time = current_time.replace(":", "_")
        current_time = current_time.replace(" ", "")
        segment_amount = 1
        information = []
        for i in range(1, 8):
            f = request.files.getlist("file" + str(i))
            t = request.form.get("text" + str(i))
            temp_info = []
            if str(t) == None or "application/octet-stream" in str(f):
                continue
            temp_info.append(segment_amount)
            temp_info.append(len(f))
            temp_info.append(str(t).replace("  ", " "))
            counter = 1
            for image in f:
                temp_file_name = secure_filename(image.filename)
                filename = (
                    str(i)
                    + "_"
                    + str(counter)
                    + "__"
                    + current_time
                    + "."
                    + temp_file_name.split(".", 1)[1]
                )
                filename = filename.replace("/", "_")
                filename = filename.replace(":", "_")
                filename = filename.replace(" ", "")
                counter += 1
                image.save(os.path.join(app.config["UPLOAD_FOLDER"], filename))

                temp_info.append(app.config["UPLOAD_FOLDER"] + filename)

            information.append(temp_info)
            segment_amount += 1
        print(information)

        final_htmls = []
        page_counter=1
        for part in information:
            if part[1] == 3:
                html_template = set_html_template(
                    str(part[3]),
                    str(part[4]),
                    str(part[5]),
                    str(part[2]),
                    html_2_pic_template,
                    page_counter,
                    part[0],                    
                )
            html_filename = str(part[0]) + "__" + str(current_time) + ".html"
            html_filename = html_filename.replace("/", "_")
            html_filename = html_filename.replace(":", "_")
            html_filename = html_filename.replace(" ", "")
            f = open(html_filename, "a", encoding="utf-8")
            f.write(html_template)
            f.close()
            final_htmls.append(html_filename)
            page_counter+=1
        print(final_htmls)
        options = {
            "enable-local-file-access": None,
            "page-size": "Letter",
            "margin-top": "0.75in",
            "margin-right": "0.75in",
            "margin-bottom": "0.75in",
            "margin-left": "0.75in",
            "encoding": "UTF-8",
            "no-outline": None,
        }

        pdfkit.from_file(final_htmls, current_time + ".pdf", options=options)
        print(current_time + ".pdf")
    return render_template("finish.html", pdf_name=current_time + ".pdf")


@app.route("/UploadFile", methods=["GET", "POST"])
def upload_file():
    file_name = request.form.get("filename")
    print(file_name)
    return send_file("../" + file_name, as_attachment=True)


@app.route("/display/<filename>")
def display_image(filename):
    # print('display_image filename: ' + filename)
    return redirect(url_for("static", filename="uploads/" + filename))


def set_html_template(
    pic_one, pic_two,pic_three, text, html_2_pic_template, page_num, page_amount
):
    html_template = html_2_pic_template
    html_template = html_template.replace("pic_one", pic_one)
    html_template = html_template.replace("pic_two", pic_two)
    html_template = html_template.replace("pic_three", pic_three)


    html_template = html_template.replace("text_from_form", text)
    html_template = html_template.replace(
        "page_number", "עמוד " + str(page_num) + " מתוך " + str(page_amount)
    )
    print(html_template)
    return html_template


if __name__ == "__main__":
    app.run(debug=False)
