import xlsxwriter
import PySimpleGUI as sg

sg.theme('DarkTeal9')

wk_book = xlsxwriter.Workbook("Guest_Sign_In.xlsx")
wk_sheet = wk_book.add_worksheet('Page1')

# Create the rows for entry


wk_sheet.write(0, 0, "Name")
wk_sheet.write(0, 1, "title")
wk_sheet.write(0, 2, "reason")
wk_sheet.write(0, 3, "student_id")

# Layout for the window GUI

layout = [
    [sg.Text('Please fill out the following fields:')],
    [sg.Text('Name', size=(15, 1)), sg.InputText(key="Name")],
    [sg.Text('Title', size=(15, 1)), sg.InputText(key="title")],
    [sg.Text('Reason', size=(15, 1)), sg.InputText(key='reason')],
    [sg.Text('StudentID', size=(15, 1)), sg.InputText(key='student_id')],

    [sg.Submit(), sg.Exit()]
]

window = sg.Window('Simple data entry form', layout)

# Creating a list to store dictionary info.

info_list = []

# functionality of the buttons

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':  # Exit closes the window and Excel Workbook
        break
    if event == 'Submit':  # Submit takes the data from info_list and writes to Excel sheet.
        s = info_list.append(values)
        for index, entry in enumerate(info_list):
            wk_sheet.write(index+1, 0, entry["Name"])
            wk_sheet.write(index+1, 1, entry["title"])
            wk_sheet.write(index+1, 2, entry["reason"])
            wk_sheet.write(index+1, 3, entry["student_id"])

        print(info_list)

wk_book.close()
window.close()
