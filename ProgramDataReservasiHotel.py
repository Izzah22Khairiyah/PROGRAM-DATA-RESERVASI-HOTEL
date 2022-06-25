import PySimpleGUI as sg
import pandas as pd
import traceback

sg.theme('SandyBeach')

dataexcel = "C:/Users/IZZAH KHAIRIYAH/Documents/data materi kuliah/semester 2/algoritma dan pemrograman/Tubes/DataReservasiHotel.xlsx"
df = pd.read_excel(dataexcel)

layout = [
    [sg.Text('Masukan data pemesanan dibawah:')],
    [sg.Text('Nama', size=(15,1)), sg.InputText(key='Nama')],
    [sg.Text('Asal', size=(15,1)), sg.InputText(key='Asal')],
    [sg.Text('No. Telepon', size=(15,1)), sg.InputText(key='No. Telepon')],
    [sg.Text('Tipe Kamar', size=(15,1)), sg.Combo(['Single', 'Double', 'Suite'], key='Tipe Kamar')],
    [sg.CalendarButton('Tanggal Check-In', size=(15,1), close_when_date_chosen=True, no_titlebar=False, format=('%Y-%m-%d') )],
    [sg.Text('Berapa Malam', size=(15,1)), sg.Spin([i for i in range(1,90)],
                                                       initial_value=1, key='Lama Stay')],
    [sg.Submit(), sg.Button('Hapus'), sg.Exit()]
]

window = sg.Window('Program Pencatatan Reservasi Hotel', layout)

def hapusinput():
    for key in values:
        window[key]('')
    return None

def logError(error) :
    print (f"Program error")
    traceback.print_tb(error.__traceback__)

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Hapus':
        hapusinput()
    if event == 'Submit':
        new_record = pd.DataFrame(values, index=[0])
        df = pd.concat([df, new_record], ignore_index=True)
        df.to_excel(dataexcel, index=False)
        sg.popup('Data Berhasil di Input!')
        hapusinput()
window.close()