import pdfplumber
import pandas as pd
import openpyxl
import openpyxl.styles
from openpyxl.styles import Font, Border
import time
from datetime import date
import json
# TODO: add .gitignore to this file - read more about this type of file and reach me out if you need me DONE
# TODO: please move those const to JSON file called config DONE
# TODO: Please read on PEP8 how constants should look on python DONE


""" Converting every line in the pdf into a line into an excel it created with the name of the pdf + "result" """
def pdf_to_txt(file):
    # TODO: please create a function docstring to all the functions - use """ and then press enter where this line is located DONE
    lines = []
    with pdfplumber.open(file) as pdf:
        # TODO: you use pages only once, remove this and replace the for loop with this line: DONE
        # for page in pdf.pages:
        for page in pdf.pages:
            text = page.extract_text()
            #  TODO: make the '\n' a constant \ configable NO NEED
            for line in text.split('\n'):
                lines.append(line)

    # TODO: if you use a var once, consider just removing it, unless it made for readability - but thats not the case DONE
    # TODO: please change "df" var to be more informative. "df" is not understandable. DONE
    temporary_excel = pd.DataFrame(lines)
    # TODO: please explaine why you used the "-4" and consider to move it to a constant DONE
    #removing the .pfd and replacing it with its new name and file type
    excel_file_name = file[:-4] + " result.xlsx"
    temporary_excel.to_excel(excel_file_name)
    # TODO: you already has a var that contains file[:-4], why not to use it? why to calculate it again? DONE
    information_extractor(excel_file_name, file[:-4])


#Gets the excel file and checks for the information
# TODO: excelFile is not formatted as var names should be in python. read more about it on PEP8. DONE
def information_extractor(excel_file, file_name):
    # TODO: move this path to a config file NO NEED
    json_file = open('config.json', encoding="utf8")
    json_data = json.load(json_file)

    book = openpyxl.load_workbook(r'C:\Users\Nadav\PycharmProjects\TaboAccessibility/' + excel_file)
    sheet = book.active

    # TODO: move to config DONE
    people_row_count = json_data['excel_row_information_start']
    company_row_count = json_data['excel_row_information_start']
    passport_row_count = json_data['excel_row_information_start']

    #A and B row in the excel, have different data at first
    a_row_in_excel = sheet['A']
    b_row_in_excel = sheet['B']

    #Finds the type of file and writing it on the excel
    type_of_file = 0
    for cell in b_row_in_excel:
        if isinstance(cell.internal_value, str):
            type_of_file = find_file_type(cell.internal_value, sheet)
            if type_of_file == 1 or type_of_file == 2:
                # TODO: create a function for this, this is a redundant code DONE
                # TODO: make sure to explain those breaks DONE
                print(type_of_file)
                #break if type of file were found
                break
    # TODO: create type_of_file var before the for loop and on the top of this function. it's best practice DONE

    #Resetting the design on the cells
    # TODO: a_row is not an informative name, I guess this is the a row of the excel - please add a comment to explain it DONE
    # TODO: please add comments to explain more of your code - I saw you add some, but please make sure to add more
    for cell in a_row_in_excel:
        # TODO: there is some redundant code here, consider moving it to a function DONE
        clear_excel_cell(cell)

    #Checking for the information
    for cell in b_row_in_excel:
        info = cell.internal_value
        clear_excel_cell(cell)

        if isinstance(info, str):

            #Checking if there`s ID in the line
            if "ז.ת" in info:
                info = info + " "
                info = " ".join(info.split())
                info = " " + info
                print(info)

                if type_of_file == 1:
                    # Find the ID and putting it in the excel (more complicated check, ID comes in multiple lengths)
                    if " " in info[info.find("ז.ת") - 9:info.find("ז.ת") - 1]:
                        id_value = info[info.find("ז.ת") - 8:info.find("ז.ת") - 1]
                    elif " " in info[info.find("ז.ת") - 10:info.find("ז.ת") - 1]:
                        id_value = info[info.find("ז.ת") - 9:info.find("ז.ת") - 1]
                    elif " " in info[info.find("ז.ת") - 11:info.find("ז.ת") - 1]:
                        id_value = info[info.find("ז.ת") - 10:info.find("ז.ת") - 1]
                    elif " " in info[info.find("ז.ת") - 12:info.find("ז.ת") - 1]:
                        id_value = info[info.find("ז.ת") - 11:info.find("ז.ת") - 1]
                    elif " " in info[info.find("ז.ת") - 13:info.find("ז.ת") - 1]:
                        id_value = info[info.find("ז.ת") - 12:info.find("ז.ת") - 1]
                    else:
                        id_value = info[info.find("ז.ת") - 13:info.find("ז.ת") - 1]
                    sheet.cell(row=people_row_count, column=1).value = id_value
                    sheet.cell(row=people_row_count, column=1).font = Font(size=11, bold=False)

                    # Find the name and putting it in the excel by certain distance from the ID and the reason
                    name_value = info[info.find("ז.ת") + 3:info.find("ז.ת") + find_name_shared_homes(info)][::-1]
                    sheet.cell(row=people_row_count, column=2).value = name_value
                    sheet.cell(row=people_row_count, column=2).font = Font(size=11, bold=False)

                    # Printing for debugging
                    print(id_value)
                    print(name_value)

                if type_of_file == 2:
                    #Find the ID and putting it in the excel (ID is always in the start of the file, very simple check)
                    id_value = info[:info.find("ז.ת")]
                    sheet.cell(row=people_row_count, column=1).value = id_value
                    sheet.cell(row=people_row_count, column=1).font = Font(size=11, bold=False)

                    #Find the name and putting it in the excel by certain distance from the ID and the reason
                    name_value = info[info.find("ז.ת") + 3:info.find("ז.ת") + find_name_shared_rights(info)][::-1]
                    sheet.cell(row=people_row_count, column=2).value = name_value
                    sheet.cell(row=people_row_count, column=2).font = Font(size=11, bold=False)

                    #Printing for debugging
                    print(id_value)
                    print(name_value)

                #Adding 1 to the index of where the program will write
                people_row_count += 1
            #Checking if there`s company and not mortgage in the line
            elif "הרבח" in info and 'התנכשמ' not in info:
                info += " "
                info = " " + info
                info = " ".join(info.split())
                print(info)

                if type_of_file == 1:
                    # Find the company ID and putting it in the excel
                    company_id_value = info[info.find("הרבח") - 10:info.find("הרבח") - 1]
                    sheet.cell(row=company_row_count, column=4).value = company_id_value
                    sheet.cell(row=company_row_count, column=4).font = Font(size=11, bold=False)

                    #Find the name and putting it in the excel by certain distance from the ID and the reason
                    company_name_value = info[info.find("הרבח") + 4:info.find("הרבח") + find_company_name_shared_homes(info)][::-1]
                    sheet.cell(row=company_row_count, column=5).value = company_name_value
                    sheet.cell(row=company_row_count, column=5).font = Font(size=11, bold=False)

                    #Printing for debugging
                    print(company_name_value)
                    print(company_id_value)

                if type_of_file == 2:
                    # Find the company ID and putting it in the excel
                    company_id_value = info[info.find("הרבח") - 10:info.find("הרבח") - 1]
                    sheet.cell(row=company_row_count, column=4).value = company_id_value
                    sheet.cell(row=company_row_count, column=4).font = Font(size=11, bold=False)

                    #Find the name and putting it in the excel by certain distance from the ID and the reason
                    company_name_value = info[info.find("הרבח") + 4:info.find("הרבח") + find_company_name_shared_rights(info)][::-1]
                    sheet.cell(row=company_row_count, column=5).value = company_name_value
                    sheet.cell(row=company_row_count, column=5).font = Font(size=11, bold=False)

                    #Printing for debugging
                    print(company_name_value)
                    print(company_id_value)

                #Adding 1 to the index of where the program will write
                company_row_count += 1
            #Checking if there`s passport and not mortgage in the line
            elif "ןוכרד" in info:
                info += " "
                info = " " + info
                info = " ".join(info.split())
                print(info)

                if type_of_file == 1:
                    #Find the passport and putting it in the excel (more complicated check, ID comes in multiple lengths)
                    if " " in info[info.find("ןוכרד") - 9:info.find("ןוכרד") - 1]:
                        passport_value = info[info.find("ןוכרד") - 8:info.find("ןוכרד") - 1]
                    elif " " in info[info.find("ןוכרד") - 10:info.find("ןוכרד") - 1]:
                        passport_value = info[info.find("ןוכרד") - 9:info.find("ןוכרד") - 1]
                    elif " " in info[info.find("ןוכרד") - 11:info.find("ןוכרד") - 1]:
                        passport_value = info[info.find("ןוכרד") - 10:info.find("ןוכרד") - 1]
                    elif " " in info[info.find("ןוכרד") - 12:info.find("ןוכרד") - 1]:
                        passport_value = info[info.find("ןוכרד") - 11:info.find("ןוכרד") - 1]
                    elif " " in info[info.find("ןוכרד") - 13:info.find("ןוכרד") - 1]:
                        passport_value = info[info.find("ןוכרד") - 12:info.find("ןוכרד") - 1]
                    else:
                        passport_value = info[info.find("ןוכרד") - 13:info.find("ןוכרד") - 1]
                    sheet.cell(row=passport_row_count, column=7).value = passport_value
                    sheet.cell(row=passport_row_count, column=7).font = Font(size=11, bold=False)

                    # Find the name and putting it in the excel by certain distance from the passport and the reason
                    passport_name_value = info[info.find("ןוכרד") + 5:info.find("ןוכרד") + find_passport_name_shared_homes(info)][::-1]
                    sheet.cell(row=passport_row_count, column=8).value = passport_name_value
                    sheet.cell(row=passport_row_count, column=8).font = Font(size=11, bold=False)

                    #Printing for debugging
                    print(passport_value)
                    print(passport_name_value)
                #Adding 1 to the index of where the program will write
                passport_row_count += 1

    # def f(sheet, row: int, col: int, value: str):
    #     pass

    #Adding titles
    # TODO: please make a function for this DONE
    write_excel_titles(sheet)

    #Saving the excel
    book.save(excel_file)

    #Adding the information to the information file
    information_file = open("InformationFile.txt", "a")
    information_file.write("\n" + file_name + " " + str(date.today().strftime("%d/%m/%Y")) + " " + str(time.strftime("%H:%M:%S", time.localtime())))
    information_file.close()


#gets a sheet and writes titles in it
def write_excel_titles(sheet):
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


#recive excel cell and clears it
def clear_excel_cell(cell):
    cell.value = None
    cell.font = Font(size=11, bold=False)
    cell.border = Border()


#writes file type
def write_file_type_in_excel(file_type, sheet):
    sheet.cell(row=1, column=10).value = "סוג קובץ:"
    sheet.cell(row=1, column=10).font = Font(size=11, bold=True)
    sheet.cell(row=2, column=10).value = file_type
    sheet.cell(row=2, column=10).font = Font(size=11, bold=False)


#Returning the distance of the name from the ID
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


#Returning the string without the reason in order to find the name
def name_reason_filtering_shared_rights(info):
    index_of_reason = 0
    json_file = open('config.json', encoding="utf8")
    data = json.load(json_file)

    for reason in data['possible_name_reasons']:
        if reason in info:
            index_of_reason = info.find(reason) + len(reason)
            break
    json_file.close()
    return index_of_reason


#Returning the string without the reason in order to find the name
def name_reason_filtering_shared_homes(info):
    json_file = open('config.json', encoding="utf8")
    data = json.load(json_file)

    for reason in data['possible_name_reasons']:
        if reason in info:
            info = info.replace(reason, "")
            break
    json_file.close()
    return info


#Returning the distance of the name from the ID
def find_name_shared_homes(info):
    length = 3
    info = info[info.find("ז.ת") + 3:]

    length += info.find(" ") + 1
    info = info[info.find(" ") + 1:]

    info = info[::-1]
    print(info)

    info = name_reason_filtering_shared_homes(info)

    info = " ".join(info.split())
    print(info)
    length += len(info)

    return length


#Returning the distance of the name from the passport ID
def find_passport_name_shared_homes(info):
    length = 5
    info = info[info.find("ןוכרד") + 5:]

    info = " ".join(info.split())

    length += info.find(" ") + 1
    info = info[info.find(" ") + 1:]

    info = " ".join(info.split())

    info = info[::-1]
    print(info)

    info = name_reason_filtering_shared_homes(info)

    info = " ".join(info.split())
    print(info)
    length += len(info)

    return length


#Returning the distance of the name from the company ID
def find_company_name_shared_homes(info):
    length = 4
    info = info[info.find("הרבח") + 4:]

    length += info.find(" ") + 1
    info = info[info.find(" ") + 1:]

    info = info[::-1]
    print(info)

    json_file = open('config.json', encoding="utf8")
    data = json.load(json_file)

    for reason in data['possible_company_name_reasons']:
        if reason in info:
            info = info.replace(reason, "")
    json_file.close()
    info = " ".join(info.split())

    print(info)
    length += len(info)

    return length


#Returning the distance of the name from the company ID
def find_company_name_shared_rights(info):
    length = 4
    info = info[info.find("הרבח") + 4:]

    length += info.find(" ") + 1
    info = info[info.find(" ") + 1:]

    info = info[::-1]
    print(info)

    json_file = open('config.json', encoding="utf8")
    data = json.load(json_file)

    for reason in data['possible_company_name_reasons']:
        if reason in info:
            info = info[:info.find(reason)]
    json_file.close()
    info = " ".join(info.split())

    print(info)
    length += len(info)

    return length


#Returning the type of the file presented by numbers
def find_file_type(info, sheet):
    if "םיפתושמ םיתב" in info:
        print(info)
        write_file_type_in_excel("בתים משותפים", sheet)
        return 1
    elif "תויוכזה סקנפמ" in info:
        print(info)
        write_file_type_in_excel("פנקס זכויות", sheet)
        return 2
    else:
        return 0


pdf_to_txt('352.pdf')
#print(find_passport_name_shared_homes(info))
#print(info[info.find("ןוכרד") - 10:info.find("ןוכרד") - 1])
#print(info[info.find("ןוכרד") + 5:info.find("ןוכרד") + find_passport_name_shared_homes(info)][::-1])


