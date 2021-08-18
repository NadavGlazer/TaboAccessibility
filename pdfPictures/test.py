import re
from flask.helpers import send_from_directory
from flask import Flask, render_template, request, send_file, redirect, url_for
from pdfkit.api import from_file
from werkzeug.utils import secure_filename
import time
from datetime import date
import os
from fpdf import FPDF
from html2image import Html2Image
from random import randint


app = Flask(__name__)

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


@app.route("/LoopStarter", methods=["GET", "POST"])
def LoopStarter():
    if request.method == "POST":
        id = randint(1000, 9999)
        current_time = get_current_time()
        text_file = open(str(id) + "__" + str(current_time) + ".txt", "w")
        text_file.close()

        if request.form.get("Mixed"):
            return render_template(
                "3ImagesMixTemplate.html",
                ID=id,
                time=str(current_time),
                PageNumber=1,
            )
        else:
            return render_template(
                "3ImagesHorizontalTemplate.html",
                ID=id,
                time=str(current_time),
                PageNumber=1,
            )


@app.route("/LoopContinue", methods=["GET", "POST"])
def LoopContinue():
    if request.method == "POST":
        # Extracting all the info from the form, both the hidden and the shown
        title_text = request.form.get("TitleTextF")
        first_pic = request.files.get("FirstPic")
        second_pic = request.files.get("SecondPic")
        third_pic = request.files.get("ThirdPic")
        current_time = request.form.get("CurrentTime")
        id = request.form.get("ID")
        TemplateType = request.form.get("TemplateType")
        page_number = request.form.get("PageNumber")

        if (
            "application/octet-stream" in str(first_pic)
            or "application/octet-stream" in str(second_pic)
            or "application/octet-stream" in str(third_pic)
        ):
            print("Got no files")
        else:

            page_number = int(page_number)
            temp_info = ""

            # "Information" has the following data: page number, pic amount, html template, title text, the pictures
            temp_info = str(page_number)
            temp_info = temp_info + "*" + str(TemplateType[0])
            temp_info = temp_info + "*" + str(TemplateType)
            temp_info = (
                temp_info + "*" + str(title_text).replace("  ", " ").replace("*", "@")
            )

            # Saving the first picure with the id,time and number and page number and saving in in the information array
            temp_file_name = secure_filename(first_pic.filename)
            file_type = temp_file_name.split(".", 1)[1]
            file_name = create_file_name(file_type, id, current_time, page_number, 1)
            first_pic.save(os.path.join(app.config["UPLOAD_FOLDER"], file_name))
            temp_info = temp_info + "*" + str(app.config["UPLOAD_FOLDER"] + file_name)

            # Saving the second picure with the id,time and number and page number and saving in in the information array
            temp_file_name = secure_filename(second_pic.filename)
            file_type = temp_file_name.split(".", 1)[1]
            file_name = create_file_name(file_type, id, current_time, page_number, 2)
            second_pic.save(os.path.join(app.config["UPLOAD_FOLDER"], file_name))
            temp_info = temp_info + "*" + str(app.config["UPLOAD_FOLDER"] + file_name)

            # Saving the third picure with the id,time and number and page number and saving in in the information array
            temp_file_name = secure_filename(third_pic.filename)
            file_type = temp_file_name.split(".", 1)[1]
            file_name = create_file_name(file_type, id, current_time, page_number, 3)
            third_pic.save(os.path.join(app.config["UPLOAD_FOLDER"], file_name))
            temp_info = temp_info + "*" + str(app.config["UPLOAD_FOLDER"] + file_name)

            text_file = open(id + "__" + current_time + ".txt", "a")
            text_file.write(temp_info + "\n")
            text_file.close()
            print(temp_info)

            page_number += 1

        if request.form.get("NewHorizontal"):
            return render_template(
                "3ImagesHorizontalTemplate.html",
                ID=id,
                time=current_time,
                PageNumber=page_number,
            )
        elif request.form.get("NewMix"):
            return render_template(
                "3ImagesMixTemplate.html",
                ID=id,
                time=current_time,
                PageNumber=page_number,
            )
        else:
            text_file = open(id + "__" + current_time + ".txt", "r+")

            information = []
            for line in text_file:
                information.append(line.split("*"))

            print(information)

            if information == None:
                return render_template("index.html")
            text_file.close()

            final_images = []
            page_amount = information[-1][0]
            for part in information:
                html_template = set_html_template(
                    str(part[4]),
                    str(part[5]),
                    str(part[6]),
                    str(part[3]),
                    part[0],
                    page_amount,
                    str(part[2]),
                )
                image_file_name = (
                    str(part[0]) + "__" + id + "__" + str(current_time) + ".png"
                )

                hti = Html2Image()
                hti.screenshot(
                    html_str=html_template,
                    save_as=(image_file_name[:-5] + ".png"),
                    size=(794, 1123),
                )
                final_images.append(image_file_name[:-5] + ".png")
            print(final_images)

            pdf = FPDF()
            for image in final_images:
                pdf.add_page()
                pdf.image("header.png", w=190, h=15)
                pdf.image(image, w=190, h=246, type="png")
            pdf.output(id + "__" + current_time + ".pdf", "F")
            print(id + "__" + current_time + ".pdf")
        return render_template(
            "finish.html", pdf_name=id + "__" + current_time + ".pdf"
        )


@app.route("/Star", methods=["GET", "POST"])
def trying():
    if request.method == "POST":
        current_time = get_current_time()
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
        page_amount = information[-1][0]
        for part in information:
            if part[1] == 3:
                html_template = set_html_template(
                    str(part[3]),
                    str(part[4]),
                    str(part[5]),
                    str(part[2]),
                    part[0],
                    page_amount,
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
                html_file=html_filename,
                save_as=(html_filename[:-5] + ".png"),
                size=(794, 1123),
            )
            final_images.append(html_filename[:-5] + ".png")
        print(final_images)

        pdf = FPDF()
        for image in final_images:
            pdf.add_page()
            pdf.image("header.png", w=190, h=15)
            pdf.image(image, w=190, h=246, type="png")
        pdf.output(current_time + ".pdf", "F")
        print(current_time + ".pdf")
    return render_template("finish.html", pdf_name=current_time + ".pdf")


@app.route("/UploadFile", methods=["GET", "POST"])
def upload_file():
    file_name = request.form.get("filename")
    return send_file("../" + file_name, as_attachment=True)


def set_html_template(
    pic_one, pic_two, pic_three, text, page_num, page_amount, file_type
):
    html_template_name = "pdfPictures/templates/" + file_type + ".html"
    f = open(html_template_name, "r")
    html_template = f.read()
    f.close()

    html_template = html_template.replace("pic_one", pic_one)
    html_template = html_template.replace("pic_two", pic_two)
    html_template = html_template.replace("pic_three", pic_three)

    html_template = html_template.replace("text_from_form", text)
    html_template = html_template.replace(
        "page_number", "עמוד " + str(page_num) + " מתוך " + str(page_amount)
    )
    return html_template


def get_current_time():
    current_time = (
        str(date.today().strftime("%d/%m/%Y"))
        + "_"
        + str(time.strftime("%H:%M:%S", time.localtime()))
    )
    current_time = current_time.replace("/", "_")
    current_time = current_time.replace(":", "_")
    current_time = current_time.replace(" ", "")
    return current_time


def create_file_name(file_type, id, current_time, page, counter):
    file_name = (
        str(page)
        + "_"
        + str(counter)
        + "__"
        + id
        + "__"
        + current_time
        + "."
        + file_type
    )
    file_name = file_name.replace("/", "_")
    file_name = file_name.replace(":", "_")
    file_name = file_name.replace(" ", "")
    return file_name


if __name__ == "__main__":
    app.run(debug=False)
