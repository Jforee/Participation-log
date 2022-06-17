import PySimpleGUI as sg

sg.theme('DarkTeal9')


layout = [
    [sg.Text('Please fill out the following fields:')],
    [sg.Text('Name', size=(15, 1)), sg.InputText(key='Name')],
    [sg.Text('Appt', size=(15, 1)), sg.InputText(key='Appt')],
    [sg.Text('Reason', size=(15, 1)), sg.InputText(key='Reason')],
    [sg.Text('StudentID', size=(15, 1)), sg.InputText(key='StudentID')],

    [sg.Submit(), sg.Exit()]
]

window = sg.Window('Simple data entry form', layout)

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Submit':
        print(event, values)
window.close()