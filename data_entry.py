from pathlib import Path
import PySimpleGUI as sg
import xlwings as xw  # pip install xlwings

# Add some color to the window
sg.theme('DarkTeal9')

EXCEL_FILE = Path(__file__).parent / "Data_Entry.xlsx"

layout = [
    [sg.Text('Please fill out the following fields:')],
    [sg.Text('Select Sheet', size=(15,1)), sg.Combo(['Sheet1', 'Sheet2'], key='-SHEET_NAME-')],
    [sg.Text('Name', size=(15,1)), sg.InputText(key='Name')],
    [sg.Text('City', size=(15,1)), sg.InputText(key='City')],
    [sg.Text('Favorite Colour', size=(15,1)), sg.Combo(['Green', 'Blue', 'Red'], key='Favorite Colour')],
    [sg.Text('I speak', size=(15,1)),
                            sg.Checkbox('German', key='German'),
                            sg.Checkbox('Spanish', key='Spanish'),
                            sg.Checkbox('English', key='English')],
    [sg.Text('No. of Children', size=(15,1)), sg.Spin([i for i in range(0,16)],
                                                       initial_value=0, key='Children')],
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

window = sg.Window('Simple data entry form', layout)

def clear_input():
    for key in values:
        window[key]('')
    return None


while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Clear':
        clear_input()
    if event == 'Submit':
        with xw.App(visible=False) as app:
            # Delete SHEET_NAME from values dict
            sheet_name = values['-SHEET_NAME-']
            del values['-SHEET_NAME-']
            entries = list(values.values())
            wb = app.books.open(EXCEL_FILE)
            sht = wb.sheets(sheet_name)
            last_row = sht.used_range.last_cell.row + 1
            sht.range('A' + str(last_row)).value = entries
            wb.save()
            wb.close()
        sg.popup('Data saved!')
        clear_input()
window.close()
