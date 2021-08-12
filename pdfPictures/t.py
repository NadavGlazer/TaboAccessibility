import pdfkit
from fpdf import FPDF
import imgkit
from html2image import Html2Image

options = {"enable-local-file-access": None}

#imgkit.from_file("C:/Users/Nadav1/TaboProject/TaboAccessibility/pdfPictures/templates/3pic.html","test.jpg",options=options)
pdf = FPDF()
        # imagelist is the list with all image filenames
pdf.add_page()
pdf.image("1__12_08_2021_15_15_59.png", w = 190, h = 260,type="png")
pdf.output("yourfile.pdf", "F")
# hti = Html2Image()
# hti.screenshot(
#     html_str="<img src='C:/Users/Nadav1/TaboProject/TaboAccessibility/pdfPictures/static/uploads/1_1__11_08_2021_14_19_41.jpg'>",
#     save_as='blue_page.png'
# )
