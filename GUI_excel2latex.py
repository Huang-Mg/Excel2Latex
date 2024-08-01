import func_excel2latex
import pyperclip
from PySimpleGUI import theme, Text, In, FileBrowse, Combo, Button, Frame, Window, ML, popup

theme('Darkgrey9')

layout_ML = [[ML(key='-OUTPUTTEXT-', size=(70, 20), autoscroll=True, disabled=True)]]
layout = [
    [Text('Excel File', size=(9, 1)),
     In(key='-FOLDERNAME-', size=(50, 1), disabled=True, enable_events=True, background_color='grey', text_color='black'),
     FileBrowse(target='-FOLDERNAME-', size=(6, 1), file_types=(('All Files', '*.xlsx'),))],
    [Text('Sheet Name', size=(9, 1)),
     Combo([], key='-SHEETNAME-', size=(48, 1), background_color='white', text_color='black', readonly=True, enable_events=True),
     Button('Convert', key='-CONVERT-', size=(6, 1))],
    [Frame(title='Output', layout=layout_ML)],
    [Button('Copy Text', key='-COPY-', size=(10, 1)), Button('Exit', key='-EXIT-', size=(10, 1))]
]

window = Window('Excel2Latex', layout, finalize=True)
window.TKroot.iconbitmap('e2l.ico')
while True:
    event, values = window.read()
    if event == None:
        break

    if event == '-FOLDERNAME-':
        file_name = values['-FOLDERNAME-']
        list_name = func_excel2latex.get_excel_sheet_name_list(file_name)
        window['-SHEETNAME-'].update(values=list_name)

    if event == '-SHEETNAME-':
        sh_name = values['-SHEETNAME-']

    if event == '-CONVERT-':
        window['-OUTPUTTEXT-'].update(value='')
        try:
            func_excel2latex.excel_convert_to_text(file_name, sh_name)
            with open('tex.txt', 'r') as f:
                content = f.read()
            window['-OUTPUTTEXT-'].update(value=content)
        except:
            popup('Please Choose File and Sheet First!', title='ERROR')

    if event == '-COPY-':
        text = values['-OUTPUTTEXT-']
        pyperclip.copy(text)

    if event == '-EXIT-':
        break
window.close()
