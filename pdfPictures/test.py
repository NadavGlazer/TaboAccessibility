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
import imgkit
from fpdf import FPDF
from html2image import Html2Image


app = Flask(__name__)
html_2_pic_template = """
<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Document</title>
    <style>
        .MainDiv {
            width: 794px;
            height: 1123px;
        }

        .HeaderImage {
            width: 750px;
            display: block;
            margin-left: auto;
            margin-right: auto;
        }

        .TitleDiv {
            margin: auto;
            border: 3px solid black;
            padding: 5px;
            width: 300px;
            height: 100px;
            margin-top: 10px;
        }
        .TitleText{
            text-align: right;
            margin-top: 0px;
            white-space: pre-line;
            line-height: 0.8;
            font-size: 25px;
            font-weight: 200;
        }
        .HorizontalPicsDiv {
            display: flex;
            justify-content: center;
            margin-top: 30px;

        }

        .FirstPicDiv {
            display: table-cell;
        }

        .FirstImage {
            width: 300px;
            height: 480px;
            margin-right: 1px;
        }

        .SecondPicDiv {
            display: table-cell;
        }

        .SecondImage {
            width: 300px;
            height: 480px;
            margin-left: 1px;
        }

        .VerticalPicDiv {
            display: flex;
            justify-content: center;
            margin-top: -2px;
        }

        .ThirdImage {
            width: 602px;
            height: 350px;
        }

        .PageNumberDiv {
            float: right;
            margin-right: 30px;
            margin-top: 10px;
        }
    </style>
</head>

<body>
    <div class="MainDiv">
        <img src="C:/Users/Nadav1/TaboProject/TaboAccessibility/header.jfif" class="HeaderImage">
        <div class="TitleDiv">
            <p class="TitleText">text_from_form</p>
        </div>
        <div class="HorizontalPicsDiv">
            <div class="FirstPicDiv">
                <img class="FirstImage" src="pic_one">
            </div>
            <div class="SecondPicDiv">
                <img class="SecondImage" src="pic_two">
            </div>
        </div>
        <div class="VerticalPicDiv">
            <img class="ThirdImage" src="pic_three">
        </div>
        <div class="PageNumberDiv">
            <h2>page_number</h2>
        </div>
    </div>

</body>

</html>"""
# UPLOAD_FOLDER = "pdfpictures/static/uploads/"
UPLOAD_FOLDER = (
    "C:/Users/Nadav1/TaboProject/TaboAccessibility/pdfPictures/static/uploads/"
)
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

        final_images = []
        page_counter = 1
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

            hti = Html2Image()
            hti.screenshot(
                html_file=html_filename, save_as=(html_filename[:-5] + ".PNG"), size=(794, 1123)
            )
            final_images.append(html_filename[:-5] + ".PNG")
            page_counter += 1
        print(final_images)

        pdf = FPDF()
        for image in final_images:
            pdf.add_page()
            pdf.image("header.png", w = 190, h=15)
            pdf.image(image, w = 190, h = 246,type="PNG")
        pdf.output(current_time + ".pdf", "F")
        print(current_time + ".pdf")
    return render_template("finish.html", pdf_name=current_time + ".pdf")


@app.route("/UploadFile", methods=["GET", "POST"])
def upload_file():
    file_name = request.form.get("filename")
    print(file_name)
    return send_file("../" + file_name, as_attachment=True)


def set_html_template(
    pic_one, pic_two, pic_three, text, html_2_pic_template, page_num, page_amount
):
    f = open("pdfPictures/templates/3pic.html", "r")
    html_template= f.read()
    f.close()

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
