# TODO: always use isort
# TODO: remove unused imports
# TODO: format doc using the shortcut
# TODO: use pylint on files when you finish a major work
# TODO: always use auto save
# TODO: use winkey+V
import os
from random import randint

from flask import Flask, redirect, render_template, request, send_file, url_for
from flask.helpers import send_from_directory
from fpdf import FPDF
from html2image import Html2Image
from pdfkit.api import from_file
from werkzeug.utils import secure_filename
import utils
import json


app = Flask(__name__)
# TODO: add this information to readme and move to config.json file
json_file = open("pdfPictures\config.json", encoding="utf8")
json_data = json.load(json_file)

UPLOAD_FOLDER = json_data["main_computer_upload_folder"]
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["MAX_CONTENT_PATH"] = json_data["MAX_CONTENT_PATH"]
app.config["TIME_OUT"] = json_data["TIME_OUT"]


# TODO: remove prints and use logger - https://flask.palletsprojects.com/en/2.0.x/logging/
@app.route("/")
def index():
    print(request.cookies)
    print(request.remote_addr)
    return render_template("index.html")


# TODO: remove get
@app.route("/LoopStarter", methods=["POST"])
def LoopStarter():
    if request.method == "POST":
        # TODO: dont use saved word id
        file_id = randint(
            json_data["id_random_range"][0], json_data["id_random_range"][1]
        )
        current_time = utils.get_current_time()
        # TODO: move the file name creation to utils class
        # TODO: read about context managers (with open ....)
        open(utils.generate_text_file_name(file_id, current_time), "a").close()

        # TODO: read about short if
        if request.form.get("Mixed"):
            template_name = json_data["3_images_mixed_html_template_name"]
        elif request.form.get("Vertical"):
            template_name = json_data["3_images_vertical_html_template_name"]

        return render_template(
            template_name,
            FileID=file_id,
            Time=current_time,
            PageNumber=1,
        )


@app.route("/LoopContinue", methods=["POST"])
def LoopContinue():
    if request.method == "POST":
        # Extracting all the info from the form, both the hidden and the shown
        title_text = request.form.get("TitleTextF")
        first_pic = request.files.get("FirstPic")
        second_pic = request.files.get("SecondPic")
        third_pic = request.files.get("ThirdPic")
        current_time = request.form.get("CurrentTime")
        file_id = request.form.get("FileID")
        TemplateType = request.form.get("TemplateType")
        page_number = request.form.get("PageNumber")

        if (
            "application/octet-stream" in str(first_pic)
            or "application/octet-stream" in str(second_pic)
            or "application/octet-stream" in str(third_pic)
        ):
            print("Got no files")
        else:

            temp_info = ""

            # "Information" has the following data: page number, pic amount, html template, title text, the pictures
            temp_info = page_number
            temp_info = temp_info + "*" + str(TemplateType[0])
            temp_info = temp_info + "*" + str(TemplateType)
            temp_info = utils.get_fixed_title_text(temp_info, title_text)

            # Saving the first picure with the id,time and number and page number and saving in in the information array
            temp_file_name = secure_filename(first_pic.filename)
            file_type = utils.get_file_type(temp_file_name)
            file_name = utils.generate_picture_file_name(
                file_type, file_id, current_time, page_number, 1
            )
            first_pic.save(os.path.join(app.config["UPLOAD_FOLDER"], file_name))
            temp_info = temp_info + "*" + str(app.config["UPLOAD_FOLDER"] + file_name)

            # Saving the second picure with the id,time and number and page number and saving in in the information array
            temp_file_name = secure_filename(second_pic.filename)
            file_type = utils.get_file_type(temp_file_name)
            file_name = utils.generate_picture_file_name(
                file_type, file_id, current_time, page_number, 2
            )
            second_pic.save(os.path.join(app.config["UPLOAD_FOLDER"], file_name))
            temp_info = temp_info + "*" + str(app.config["UPLOAD_FOLDER"] + file_name)

            # Saving the third picure with the id,time and number and page number and saving in in the information array
            temp_file_name = secure_filename(third_pic.filename)
            file_type = utils.get_file_type(temp_file_name)
            file_name = utils.generate_picture_file_name(
                file_type, file_id, current_time, page_number, 3
            )
            third_pic.save(os.path.join(app.config["UPLOAD_FOLDER"], file_name))
            temp_info = temp_info + "*" + str(app.config["UPLOAD_FOLDER"] + file_name)

            print(temp_info)

            text_file = open(
                utils.generate_text_file_name(file_id, current_time),
                "a",
                encoding="utf-8",
            )
            text_file.write(temp_info + "\n")
            text_file.close()

            page_number = int(page_number)
            page_number += 1

        if request.form.get("NewVertical"):
            return render_template(
                json_data["3_images_vertical_html_template_name"],
                FileID=file_id,
                Time=current_time,
                PageNumber=page_number,
            )
        elif request.form.get("NewMix"):
            return render_template(
                json_data["3_images_mixed_html_template_name"],
                FileID=file_id,
                Time=current_time,
                PageNumber=page_number,
            )
        else:
            text_file = open(
                utils.generate_text_file_name(file_id, current_time),
                "r+",
                encoding="utf-8",
            )

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
                html_template = utils.set_html_template(
                    str(part[4]),
                    str(part[5]),
                    str(part[6]),
                    str(part[3]),
                    part[0],
                    page_amount,
                    str(part[2]),
                )
                image_file_name = utils.generate_page_image_file_name(
                    str(part[0]), file_id, current_time, json_data["page_image_type"]
                )
                hti = Html2Image()
                hti.screenshot(
                    html_str=html_template,
                    save_as=image_file_name,
                    size=(
                        json_data["screenshot_size"][0],
                        json_data["screenshot_size"][1],
                    ),
                )
                final_images.append(image_file_name)
            print(final_images)

            pdf = FPDF()
            for image in final_images:
                pdf.add_page()
                pdf.image(
                    json_data["header_picture_path"],
                    w=json_data["header_picure_size"][0],
                    h=json_data["header_picure_size"][1],
                )
                pdf.image(
                    image,
                    w=json_data["pdf_body_size"][0],
                    h=json_data["pdf_body_size"][1],
                    type=json_data["page_image_type"],
                )
            pdf_file_name = utils.generate_pdf_file_name(file_id, current_time)
            pdf.output(pdf_file_name, "F")
            print(pdf_file_name)
        return render_template("finish.html", pdf_name=pdf_file_name)


@app.route("/UploadFile", methods=["GET", "POST"])
def upload_file():
    file_name = request.form.get("filename")
    return send_file("../" + file_name, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=False)
