import xlsxwriter

# Creating the workbook and worksheet

wk_book = xlsxwriter.Workbook("Guest_Sign_In.xlsx")
wk_sheet = wk_book.add_worksheet('Page1')

# Create the rows for entry

wk_sheet.write(0, 0, "Name")
wk_sheet.write(0, 1, "appt")
wk_sheet.write(0, 2, "reason")
wk_sheet.write(0, 3, "student_id")

# Function to ask user input for their info and writes it in.

def sign_in():

    # Gathering user input.
    name = str(input("Name?:"))
    appt = str(input("Appt? Y/N:"))
    reason = str(input("Reason:"))
    student_id = str(int(input("Student_ID:")))

    # Create the dictionary
    data_list = [{
        "Name": name,
        "appt": appt,
        "reason": reason,
        "student_id": student_id,
    }]

    # Writes information from dictionary into Excel sheet.
    for index, entry in enumerate(data_list):
        wk_sheet.write(index + 1, 0, entry["Name"])
        wk_sheet.write(index + 1, 1, entry["appt"])
        wk_sheet.write(index + 1, 2, entry["reason"])
        wk_sheet.write(index + 1, 3, entry["student_id"])


sign_in()

wk_book.close()


