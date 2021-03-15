import pdfplumber
import pandas as pd
import openpyxl
import openpyxl.styles
from openpyxl.styles import Font, Border


def pdf_to_txt(file):
    lines = []
    with pdfplumber.open(file) as pdf:
        pages = pdf.pages

        for page in pages:
            text = page.extract_text()
            for line in text.split('\n'):
                lines.append(line)

    df = pd.DataFrame(lines)
    excel_file_name = file[:-4] + " result.xlsx"
    df.to_excel(excel_file_name)
    information_extractor(excel_file_name)


def information_extractor(excelFile):
    book = openpyxl.load_workbook(r'C:\Users\Nadav\PycharmProjects\TaboAccessibility/' + excelFile)
    sheet = book.active

    people_count = 2
    company_count = 2
    passport_count = 2

    a_row = sheet['A']
    b_row = sheet['B']

    for cell in b_row:
        if isinstance(cell.internal_value, str):
            result = find_file_type(cell.internal_value)
            if result == 1:
                type_of_file = result
                sheet.cell(row=1, column=10).value = "סוג קובץ:"
                sheet.cell(row=1, column=10).font = Font(size=11, bold=True)
                sheet.cell(row=2, column=10).value = "בתים משותפים"
                sheet.cell(row=2, column=10).font = Font(size=11, bold=False)
                break
            elif result == 2:
                type_of_file = result
                break

    print(type_of_file)

    for cell in a_row:
        cell.value = None
        cell.font = Font(size=11, bold=False)
        cell.border = Border()

    for cell in b_row:
        info = cell.internal_value
        cell.value = None
        cell.border = Border()
        if isinstance(info, str):
            if "ז.ת" in info:
                info = info + " "
                info = " ".join(info.split())
                info = " " + info
                print(info)
                if type_of_file == 2:
                    sheet.cell(row=people_count, column=1).value = info[:info.find("ז.ת")]
                    sheet.cell(row=people_count, column=1).font = Font(size=11, bold=False)
                    sheet.cell(row=people_count, column=2).value = info[info.find("ז.ת")+3:info.find("ז.ת") + find_name_shared_rights(info)][::-1]
                    sheet.cell(row=people_count, column=2).font = Font(size=11, bold=False)
                    print(info[:info.find("ז.ת")])
                    print(info[info.find("ז.ת")+3: info.find("ז.ת") + find_name_shared_rights(info)][::-1])
                if type_of_file == 1:
                    if " " in info[info.find("ז.ת") - 9:info.find("ז.ת") - 1]:
                        sheet.cell(row=people_count, column=1).value = info[info.find("ז.ת") - 8:info.find("ז.ת") - 1]
                        print(info[info.find("ז.ת") - 8:info.find("ז.ת") - 1])
                    elif " " in info[info.find("ז.ת") - 10:info.find("ז.ת") - 1]:
                        sheet.cell(row=people_count, column=1).value = info[info.find("ז.ת") - 9:info.find("ז.ת") - 1]
                        print(info[info.find("ז.ת") - 9:info.find("ז.ת") - 1])
                    elif " " in info[info.find("ז.ת") - 11:info.find("ז.ת") - 1]:
                        sheet.cell(row=people_count, column=1).value = info[info.find("ז.ת") - 10:info.find("ז.ת") - 1]
                        print(info[info.find("ז.ת") - 10:info.find("ז.ת") - 1])
                    elif " " in info[info.find("ז.ת") - 12:info.find("ז.ת") - 1]:
                        sheet.cell(row=people_count, column=1).value = info[info.find("ז.ת") - 11:info.find("ז.ת") - 1]
                        print(info[info.find("ז.ת") - 11:info.find("ז.ת") - 1])
                    elif " " in info[info.find("ז.ת") - 13:info.find("ז.ת") - 1]:
                        sheet.cell(row=people_count, column=1).value = info[info.find("ז.ת") - 12:info.find("ז.ת") - 1]
                        print(info[info.find("ז.ת") - 12:info.find("ז.ת") - 1])
                    else:
                        sheet.cell(row=people_count, column=1).value = info[info.find("ז.ת") - 13:info.find("ז.ת") - 1]
                        print(info[info.find("ז.ת") - 13:info.find("ז.ת") - 1])
                    sheet.cell(row=people_count, column=1).font = Font(size=11, bold=False)
                    sheet.cell(row=people_count, column=2).value = info[info.find("ז.ת")+3:info.find("ז.ת") + find_name_shared_homes(info)][::-1]
                    sheet.cell(row=people_count, column=2).font = Font(size=11, bold=False)
                    print(info[info.find("ז.ת")+3:info.find("ז.ת") + find_name_shared_homes(info)][::-1])
                people_count += 1
            elif "הרבח" in info and 'התנכשמ' not in info:
                info += " "
                info = " " + info
                info = " ".join(info.split())
                print(info)
                if type_of_file == 1:
                    sheet.cell(row=company_count, column=4).value = info[info.find("הרבח") - 10:info.find("הרבח") - 1]
                    print(info[info.find("הרבח") - 10:info.find("הרבח") - 1])
                    sheet.cell(row=company_count, column=4).font = Font(size=11, bold=False)
                    sheet.cell(row=company_count, column=5).value = info[info.find("הרבח")+4:info.find("הרבח") + find_company_name_shared_homes(info)][::-1]
                    sheet.cell(row=company_count, column=5).font = Font(size=11, bold=False)
                    print(info[info.find("הרבח")+4:info.find("הרבח") + find_company_name_shared_homes(info)][::-1])

                if type_of_file == 2:
                    sheet.cell(row=company_count, column=4).value = info[info.find("הרבח") - 10:info.find("הרבח") - 1]
                    print(info[info.find("הרבח") - 10:info.find("הרבח") - 1])
                    sheet.cell(row=company_count, column=4).font = Font(size=11, bold=False)
                    sheet.cell(row=company_count, column=5).value = info[info.find("הרבח") + 4:info.find("הרבח") + find_company_name_shared_rights(info)][::-1]
                    sheet.cell(row=company_count, column=5).font = Font(size=11, bold=False)
                    print(info[info.find("הרבח") + 4:info.find("הרבח") + find_company_name_shared_rights(info)][::-1])
                company_count += 1
            elif "ןוכרד" in info:
                info += " "
                info = " " + info
                info = " ".join(info.split())
                print(info)
                if type_of_file == 1:
                    sheet.cell(row=passport_count, column=7).value = info[info.find("ןוכרד") - 10:info.find("ןוכרד") - 1]
                    sheet.cell(row=passport_count, column=7).font = Font(size=11, bold=False)
                    sheet.cell(row=passport_count, column=8).value = info[info.find("ןוכרד") + 5:info.find("ןוכרד") + find_passport_name_shared_homes(info)][::-1]
                    sheet.cell(row=passport_count, column=8).font = Font(size=11, bold=False)
                    print(info[info.find("ןוכרד") - 10:info.find("ןוכרד") - 1])
                    print(info[info.find("ןוכרד") + 5:info.find("ןוכרד") + find_passport_name_shared_homes(info)][::-1])
                passport_count += 1

    sheet.cell(row=1, column=1).value = "ת.ז"
    sheet.cell(row=1, column=1).font = Font(size=11, bold=True)
    sheet.cell(row=1, column=2).value = "שם"
    sheet.cell(row=1, column=2).font = Font(size=11, bold=True)
    sheet.cell(row=1, column=4).value = "מספר"
    sheet.cell(row=1, column=4).font = Font(size=11, bold=True)
    sheet.cell(row=1, column=5).value = "שם חברה"
    sheet.cell(row=1, column=5).font = Font(size=11, bold=True)
    sheet.cell(row=1, column=7).value = "מספר דרכון"
    sheet.cell(row=1, column=7).font = Font(size=11, bold=True)
    sheet.cell(row=1, column=8).value = "שם"
    sheet.cell(row=1, column=8).font = Font(size=11, bold=True)

    book.save(excelFile)


def find_name_shared_rights(info):
    length = 3
    info = info[info.find("ז.ת") + 3:]

    length += info.find(" ") + 1
    info = info[info.find(" ") + 1:]

    info = info[::-1]
    print(info)
    info = " ".join(info.split())
    info = info[name_reason_filtering_shared_rights(info):]

    info = " ".join(info.split())
    print(info)
    length += len(info)

    return length


def name_reason_filtering_shared_rights(info):
    index_of_reason = 0
    possible_reasons = ["הערת אזהרה סעיף 621", "צוואה על פי הסכם", 'צו ניהול ע"י אפוטרופוס', "עדכון פרטי זיהוי - חוכר", "עדכון פרטי זיהוי", "תיקון טעות סופר", "מכר ללא תמורה", "מכר לפי צו בית משפט",
                        "מכר", "משכנתה", "ירושה", "העברת שכירות ללא תמורה", "העברת שכירות", "שכירות", "בשלמות", "עודף", "צוואה", "שנוי שם", "תיקון טעות סופר", "תיקון צו בית משותף", "לפי צו בית משפט",
                        "על פי הסכם"]
    for reason in possible_reasons:
        if reason in info:
            index_of_reason = info.find(reason) + len(reason)
    if "הערת אזהרה סעיף 621" in info:
        index_of_reason = info.find("הערת אזהרה סעיף 621") + len("הערת אזהרה סעיף 621")

    elif "צוואה על פי הסכם" in info:
        index_of_reason = info.find("צוואה על פי הסכם") + len("צוואה על פי הסכם")

    elif 'צו ניהול ע"י אפוטרופוס' in info:
        index_of_reason = info.find('צו ניהול ע"י אפוטרופוס') + len('צו ניהול ע"י אפוטרופוס')

    elif "עדכון פרטי זיהוי" in info:
        if "עדכון פרטי זיהוי - חוכר" in info:
            index_of_reason = info.find("עדכון פרטי זיהוי - חוכר") + len("עדכון פרטי זיהוי - חוכר")

        else:
            index_of_reason = info.find("עדכון פרטי זיהוי") + len("עדכון פרטי זיהוי")

    elif "תיקון טעות סופר" in info:
        index_of_reason = info.find("תיקון טעות סופר") + len("תיקון טעות סופר")

    elif "מכר" in info:
        if "מכר ללא תמורה" in info:
            index_of_reason = info.find("מכר ללא תמורה") + len("מכר ללא תמורה")

        else:
            index_of_reason = info.find("מכר") + len("מכר")

    elif "משכנתה" in info:
        index_of_reason = info.find("משכנתה") + len("משכנתה")

    elif "ירושה" in info:
        index_of_reason = info.find("ירושה") + len("ירושה")

    elif "שכירות" in info:
        if "העברת שכירות ללא תמורה" in info:
            index_of_reason = info.find("העברת שכירות ללא תמורה") + len("העברת שכירות ללא תמורה")

        elif "העברת שכירות" in info:
            index_of_reason = info.find("העברת שכירות") + len("העברת שכירות")

        else:
            index_of_reason = info.find("שכירות") + len("שכירות")

    elif "בשלמות" in info:
        index_of_reason = info.find("בשלמות") + len("בשלמות")

    elif "עודף" in info:
        index_of_reason = info.find("עודף") + len("עודף")

    elif "צוואה" in info:
        index_of_reason = info.find("צוואה") + len("צוואה")

    elif "שנוי שם" in info:
        index_of_reason = info.find("שנוי שם") + len("שנוי שם")

    elif "תיקון טעות סופר" in info:
        index_of_reason = info.find("תיקון טעות סופר") + len("תיקון טעות סופר")

    elif "תיקון צו בית משותף" in info:
        index_of_reason = info.find("תיקון צו בית משותף") + len("תיקון צו בית משותף")

    elif "לפי צו בית משפט" in info:
        index_of_reason = info.find("לפי צו בית משפט") + len("תיקון צו בית משותף")

    elif "על פי הסכם" in info:
        index_of_reason = info.find("על פי הסכם") + len("על פי הסכם")
    return index_of_reason


def name_reason_filtering_shared_homes(info):
    if "הערת אזהרה סעיף 621" in info:
        info = info.replace("הערת אזהרה סעיף 621", "")

    elif "צוואה על פי הסכם" in info:
        info = info.replace("צוואה על פי הסכם", "")

    elif 'צו ניהול ע"י אפוטרופוס' in info:
        info = info.replace('צו ניהול ע"י אפוטרופוס', "")

    elif "עדכון פרטי זיהוי" in info:
        if "עדכון פרטי זיהוי - חוכר" in info:
            info = info.replace("עדכון פרטי זיהוי - חוכר", "")

        else:
            info = info.replace("עדכון פרטי זיהוי", "")

    elif "תיקון טעות סופר" in info:
        info = info.replace("תיקון טעות סופר", "")

    elif "מכר" in info:
        if "מכר ללא תמורה" in info:
            info = info.replace("מכר ללא תמורה", "")
        elif "מכר לפי צו בית משפט" in info:
            info = info.replace("מכר לפי צו בית משפט", "")
        else:
            info = info.replace("מכר", "")

    elif "משכנתה" in info:
        info = info.replace("משכנתה", "")

    elif "ירושה" in info:
        if "ירושה על פי הסכם" in info:
            info = info.replace("ירושה על פי הסכם", "")
        else:
            info = info.replace("ירושה", "")

    elif "שכירות" in info:
        if "העברת שכירות ללא תמורה" in info:
            info = info.replace("העברת שכירות ללא תמורה", "")

        elif "העברת שכירות" in info:
            info = info.replace("העברת שכירות", "")

        else:
            info = info.replace("שכירות", "")

    elif "בשלמות" in info:
        info = info.replace("בשלמות", "")

    elif "עודף" in info:
        info = info.replace("עודף", "")

    elif "צוואה" in info:
        info = info.replace("צוואה", "")

    elif "שנוי שם" in info:
        info = info.replace("שנוי שם", "")

    elif "תיקון טעות סופר" in info:
        info = info.replace("תיקון טעות סופר", "")

    elif "תיקון צו בית משותף" in info:
        info = info.replace("תיקון צו בית משותף", "")

    return info


def find_name_shared_homes(info):
    length = 3
    info = info[info.find("ז.ת")+3:]

    length += info.find(" ") + 1
    info = info[info.find(" ")+1:]

    info = info[::-1]
    print(info)

    info = name_reason_filtering_shared_homes(info)

    info = " ".join(info.split())
    print(info)
    length += len(info)

    return length


def find_passport_name_shared_homes(info):
    length = 5
    info = info[info.find("ןוכרד") + 5:]

    length += info.find(" ") + 1
    info = info[info.find(" ") + 1:]

    info = info[::-1]
    print(info)

    info = name_reason_filtering_shared_homes(info)

    info = " ".join(info.split())
    print(info)
    length += len(info)

    return length


def find_company_name_shared_homes(info):
    length = 4
    info = info[info.find("הרבח") + 4:]

    length += info.find(" ") + 1
    info = info[info.find(" ") + 1:]

    info = info[::-1]
    print(info)

    if 'הערת אזהרה תמ"א 83' in info:
        info = info.replace('הערת אזהרה תמ"א 83', "")
    elif "הערת אזהרה סעיף 621" in info:
        info = info.replace("הערת אזהרה סעיף 621", "")
    elif "מכר" in info:
        info = info.replace("מכר", "")

    info = " ".join(info.split())

    print(info)
    length += len(info)

    return length


def find_company_name_shared_rights(info):
    length = 4
    info = info[info.find("הרבח") + 4:]

    length += info.find(" ") + 1
    info = info[info.find(" ") + 1:]

    info = info[::-1]
    print(info)

    if 'הערת אזהרה תמ"א 83' in info:
        info = info[:info.find('הערת אזהרה תמ"א 83')]
    elif "הערת אזהרה סעיף 621" in info:
        info = info[:info.find("הערת אזהרה סעיף 621")]

    info = " ".join(info.split())

    print(info)
    length += len(info)

    return length


def find_file_type(info):
    if "םיפתושמ םיתב" in info:
        print(info)
        return 1
    elif "תויוכזה סקנפמ" in info:
        print(info)
        return 2
    else:
        return 0


pdf_to_txt('269.pdf')
#print(find_name_shared_rights(info))
#print(info[info.find("ז.ת") + 3:info.find("ז.ת") + find_name_shared_rights(info)][::-1])






