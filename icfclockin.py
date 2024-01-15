import PySimpleGUI as sg
import datetime
import openpyxl
import os

sg.theme('DarkTeal9')

layout = [[sg.Text('Työajan Seuranta', font=('Helvetica', 30))],
          [sg.Text('')],
          [sg.Button('Aloita työpäivä', size=(30, 3)), sg.Button('Päätä työpäivä', size=(30, 3))]]

window = sg.Window('ICF työajanseuranta', layout)

clocked_in = False
start_time = None

# Check if the Excel file exists
if os.path.isfile('Työajanseuranta.xlsx'):
    workbook = openpyxl.load_workbook('Työajanseuranta.xlsx')
else:
    workbook = openpyxl.Workbook()

# Select the active worksheet
worksheet = workbook.active

# Set the column headers
worksheet['A1'] = 'Työpäivä alkoi'
worksheet['B1'] = 'Työpäivä päättyi'
worksheet['C1'] = 'Työpäivän kesto'

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED:
        break
    elif event == 'Aloita työpäivä':
        if not clocked_in:
            start_time = datetime.datetime.now()
            clocked_in = True
            sg.popup('Työpäivä alkoi ' + str(start_time))
        else:
            sg.popup('Työpäiväsi on jo käynnissä!')
    elif event == 'Päätä työpäivä':
        if clocked_in:
            end_time = datetime.datetime.now()
            clocked_in = False
            time_diff = end_time - start_time

            # Write the data to the worksheet
            row = [start_time, end_time, time_diff]
            worksheet.append(row)

            # Save the workbook
            workbook.save('Työajanseuranta.xlsx')

            sg.popup('Työpäivä päättyi ' + str(end_time) + '\nTyöpäivän kesto ' + str(time_diff))
        else:
            sg.popup('Et ole aloittanut työpäivääsi!')