# -*- coding: utf-8 -*-
"""
Created on Fri Mar 10 09:08:00 2023

@author: Hakeem Alavi
"""
import PySimpleGUI as sg
import datetime
from pathlib import Path
from docxtpl import DocxTemplate

document_path = Path.cwd() / "Requisition-Generator.docx"
doc = DocxTemplate(document_path)

SYMBOL_RIGHT =    '▶'
SYMBOL_DOWN =  '▼'

def collapse(layout, key):

    # sg.pin allows us to diplay or hide the column
    return sg.pin(sg.Column(layout, key=key))


today = datetime.datetime.today()

section_1 = [
    [sg.Text("S.L No. 1:", expand_x="30"), sg.Input(key="SL_1")],
    [sg.Text("Description 1 :", expand_x="30"), sg.Input(key="DESC_1")],
    [sg.Text("Unit 1:", expand_x="30"), sg.Input(key="UNIT_1")],
    [sg.Text("Quantity 1:", expand_x="30"), sg.Input(key="QTY_1")],
    [sg.Text("Rate 1:", expand_x="30"), sg.Input(key="RATE_1")],
    [sg.Text("Details 1:", expand_x="30"), sg.Input(key="DETAILS_1")],
    [sg.Text("")],
    
    [sg.Text("S.L No. 2:", expand_x="30"), sg.Input(key="SL_2")],
    [sg.Text("Description 2 :", expand_x="30"), sg.Input(key="DESC_2")],
    [sg.Text("Unit 2:", expand_x="30"), sg.Input(key="UNIT_2")],
    [sg.Text("Quantity 2:", expand_x="30"), sg.Input(key="QTY_2")],
    [sg.Text("Rate 2:", expand_x="30"), sg.Input(key="RATE_2")],
    [sg.Text("Details 2:", expand_x="30"), sg.Input(key="DETAILS_2")],
    [sg.Text("")],
    
    
]

section_2 = [
    [sg.Text("S.L No. 3:", expand_x="30"), sg.Input(key="SL_3")],
    [sg.Text("Description 3 :", expand_x="30"), sg.Input(key="DESC_3")],
    [sg.Text("Unit 3:", expand_x="30"), sg.Input(key="UNIT_3")],
    [sg.Text("Quantity 3:", expand_x="30"), sg.Input(key="QTY_3")],
    [sg.Text("Rate 3:", expand_x="30"), sg.Input(key="RATE_3")],
    [sg.Text("Details 3:", expand_x="30"), sg.Input(key="DETAILS_3")],
    [sg.Text("")],
    
    [sg.Text("S.L No. 4:", expand_x="30"), sg.Input(key="SL_4")],
    [sg.Text("Description 4 :", expand_x="30"), sg.Input(key="DESC_4")],
    [sg.Text("Unit 4:", expand_x="30"), sg.Input(key="UNIT_4")],
    [sg.Text("Quantity 4:", expand_x="30"), sg.Input(key="QTY_4")],
    [sg.Text("Rate 4:", expand_x="30"), sg.Input(key="RATE_4")],
    [sg.Text("Details 4:", expand_x="30"), sg.Input(key="DETAILS_4")],
    [sg.Text("")],
   
]

section_3 = [
    [sg.Text("Department:", expand_x="30"), sg.Input(key="DEPARTMENT")],
    [sg.Text("Requisition no:", expand_x="30"), sg.Input(key="REQ_NO")],
    [sg.Text("L.P.O No.:", expand_x="30"), sg.Input(key="LPO_NO")],
    [sg.Text("")],
    
    [sg.Text("Purpose:", expand_x="30"), sg.Input(key="PURPOSE")],
    [sg.Text("Date:", expand_x="30"), sg.Input(key="DATE")],
    [sg.Text("Sign:", expand_x="30"), sg.Input(key="SIGN")],
    [sg.Text("")],
    
    [sg.Text("User:", expand_x="30"), sg.Input(key="USER")],
    [sg.Text("Approved by:", expand_x="30"), sg.Input(key="APPROVED")],
    [sg.Text("Accountant:", expand_x="30"), sg.Input(key="ACCOUNTANT")],
    [sg.Text("")],
    ]




layout = [
    # [sg.Text("Department name:"), sg.Input(key="DEPARTMENT", do_not_clear=False)],
    # [sg.Text("Requisition no.:"), sg.Input(key="REQ_NO", do_not_clear=False)],
    # [sg.Text("L.P.O No.:"), sg.Input(key="LPO_NO", do_not_clear=False)],
    
    # [sg.Text("S.L No. 1:"), sg.Input(key="SL_1", do_not_clear=False)],
    # [sg.Text("Description 1:"), sg.Input(key="DESC_1", do_not_clear=False)],
    # [sg.Text("Unit 1:"), sg.Input(key="UNIT_1", do_not_clear=False)],
    # [sg.Text("Quantity 1:"), sg.Input(key="QTY_1", do_not_clear=False)],
    # [sg.Text("Rate 1:"), sg.Input(key="RATE_1", do_not_clear=False)],
    # [sg.Text("Details 1:"), sg.Input(key="DETAILS_1", do_not_clear=False)],
    
    # [sg.Text("S.L No. 2:"), sg.Input(key="SL_2", do_not_clear=False)],
    # [sg.Text("Description 2:"), sg.Input(key="DESC_2", do_not_clear=False)],
    # [sg.Text("Unit 2:"), sg.Input(key="UNIT_2", do_not_clear=False)],
    # [sg.Text("Quantity 2:"), sg.Input(key="QTY_2", do_not_clear=False)],
    # [sg.Text("Rate 2:"), sg.Input(key="RATE_2", do_not_clear=False)],
    # [sg.Text("Details 2:"), sg.Input(key="DETAILS_2", do_not_clear=False)],
    
    # [sg.Text("S.L No. 3:"), sg.Input(key="SL_3", do_not_clear=False)],
    # [sg.Text("Description 3:"), sg.Input(key="DESC_3", do_not_clear=False)],
    # [sg.Text("Unit 3:"), sg.Input(key="UNIT_3", do_not_clear=False)],
    # [sg.Text("Quantity 3:"), sg.Input(key="QTY_3", do_not_clear=False)],
    # [sg.Text("Rate 3:"), sg.Input(key="RATE_3", do_not_clear=False)],
    # [sg.Text("Details 3:"), sg.Input(key="DETAILS_3", do_not_clear=False)],
    
    # [sg.Text("S.L No. 4:"), sg.Input(key="SL_4", do_not_clear=False)],
    # [sg.Text("Description 4:"), sg.Input(key="DESC_4", do_not_clear=False)],
    # [sg.Text("Unit 4:"), sg.Input(key="UNIT_4", do_not_clear=False)],
    # [sg.Text("Quantity 4:"), sg.Input(key="QTY_4", do_not_clear=False)],
    # [sg.Text("Rate 4:"), sg.Input(key="RATE_4", do_not_clear=False)],
    # [sg.Text("Details 4:"), sg.Input(key="DETAILS_4", do_not_clear=False)],
    
    # [sg.Text("Purpose of purchase:"), sg.Input(key="PURPOSE", do_not_clear=False)],
    # [sg.Text("Date:"), sg.Input(key="DATE", do_not_clear=False)],
    # [sg.Text("Sign:"), sg.Input(key="SIGN", do_not_clear=False)],
    
    # [sg.Text("User:"), sg.Input(key="USER", do_not_clear=False)],
    # [sg.Text("Approved by:"), sg.Input(key="APPROVED", do_not_clear=False)],
    # [sg.Text("Accountant:"), sg.Input(key="ACCOUNTANT", do_not_clear=False)],
    
    
    [sg.Text(SYMBOL_DOWN, enable_events=True, key='-OPEN_SEC1-'),
    sg.Text('Section 1 - ITEMS')],
    [collapse(section_1, '-SEC_1-')],

    [sg.Text(SYMBOL_DOWN, enable_events=True, key='-OPEN_SEC2-'),
    sg.Text('Section 2 - MORE ITEMS')],
    [collapse(section_2, '-SEC_2-')],
    
    [sg.Text(SYMBOL_DOWN, enable_events=True, key='-OPEN_SEC3-'),
    sg.Text('Section 3 - ESSENTIAL DETAILS')],
    [collapse(section_3, '-SEC_3-')],
    
    [sg.Text("")],
    [sg.Button("Create Requisition"),sg.Button("Clear All"), sg.Exit()],
       ]            
                    

window = sg.Window("Requisition Generator", layout, size=(650, 650), element_justification="center")

def clear_inputs():
    for key in values:
        window['DEPARTMENT'].update("")
        window['REQ_NO'].update("")
        window['LPO_NO'].update("")
        window['PURPOSE'].update("")
        window['DATE'].update("")
        window['SIGN'].update("")
        window['USER'].update("")
        window['APPROVED'].update("")
        window['ACCOUNTANT'].update("")
        window['SL_1'].update("")
        window['DESC_1'].update("")
        window['UNIT_1'].update("")
        window['QTY_1'].update("")
        window['RATE_1'].update("")
        window['DETAILS_1'].update("")
        window['SL_2'].update("")
        window['DESC_2'].update("")
        window['UNIT_2'].update("")
        window['QTY_2'].update("")
        window['RATE_2'].update("")
        window['DETAILS_2'].update("")
        window['SL_3'].update("")
        window['DESC_3'].update("")
        window['UNIT_3'].update("")
        window['QTY_3'].update("")
        window['RATE_3'].update("")
        window['DETAILS_3'].update("")
        window['SL_4'].update("")
        window['DESC_4'].update("")
        window['UNIT_4'].update("")
        window['QTY_4'].update("")
        window['RATE_4'].update("")
        window['DETAILS_4'].update("")
    return None

opened1 = True
opened2 = True
opened3 = True

while True:
    event, values = window.read() 
    if event == sg.WIN_CLOSED or event == "Exit":
        break
    if event == 'Create Requisition':
        # Add calculated fields to our dict
        values["TODAY"] = today.strftime("%Y-%m-%d")
        

        # Render the template, save new word document & inform user
        doc.render(values)
        output_path = Path.cwd()  / f"{values['REQ_NO']}-Requisition.docx"
        doc.save(output_path)
        sg.popup("File saved", f"File has been saved here: {output_path}")
    if event == '-OPEN_SEC1-':
        opened1 = not opened1
        window['-OPEN_SEC1-'].update(SYMBOL_DOWN if opened1 else SYMBOL_RIGHT)
        window['-SEC_1-'].update(visible=opened1)
    
    if event == '-OPEN_SEC2-':
        opened2 = not opened2
        window['-OPEN_SEC2-'].update(SYMBOL_DOWN if opened2 else SYMBOL_RIGHT)
        window['-SEC_2-'].update(visible=opened2)
        
    if event == '-OPEN_SEC3-':
        opened3 = not opened3
        window['-OPEN_SEC3-'].update(SYMBOL_DOWN if opened3 else SYMBOL_RIGHT)
        window['-SEC_3-'].update(visible=opened3)
        
    if event == 'Clear All':   
        clear_inputs()

window.close()













# second_layout = [
#     [sg.Text("S.L No. 1:"), sg.Input(key="SL_1", do_not_clear=False)],
#     [sg.Text("Description 1:"), sg.Input(key="DESC_1", do_not_clear=False)],
#     [sg.Text("Unit 1:"), sg.Input(key="UNIT_1", do_not_clear=False)],
#     [sg.Text("Quantity 1:"), sg.Input(key="QTY_1", do_not_clear=False)],
#     [sg.Text("Rate 1:"), sg.Input(key="RATE_1", do_not_clear=False)],
#     [sg.Text("Details 1:"), sg.Input(key="DETAILS_1", do_not_clear=False)],
    
#     [sg.Text("S.L No. 2:"), sg.Input(key="SL_2", do_not_clear=False)],
#     [sg.Text("Description 2:"), sg.Input(key="DESC_2", do_not_clear=False)],
#     [sg.Text("Unit 2:"), sg.Input(key="UNIT_2", do_not_clear=False)],
#     [sg.Text("Quantity 2:"), sg.Input(key="QTY_2", do_not_clear=False)],
#     [sg.Text("Rate 2:"), sg.Input(key="RATE_2", do_not_clear=False)],
#     [sg.Text("Details 2:"), sg.Input(key="DETAILS_2", do_not_clear=False)],
#     ]

# third_layout = [
#     [sg.Text("S.L No. 3:"), sg.Input(key="SL_3", do_not_clear=False)],
#     [sg.Text("Description 3:"), sg.Input(key="DESC_3", do_not_clear=False)],
#     [sg.Text("Unit 3:"), sg.Input(key="UNIT_3", do_not_clear=False)],
#     [sg.Text("Quantity 3:"), sg.Input(key="QTY_3", do_not_clear=False)],
#     [sg.Text("Rate 3:"), sg.Input(key="RATE_3", do_not_clear=False)],
#     [sg.Text("Details 3:"), sg.Input(key="DETAILS_3", do_not_clear=False)],
    
#     [sg.Text("S.L No. 4:"), sg.Input(key="SL_4", do_not_clear=False)],
#     [sg.Text("Description 4:"), sg.Input(key="DESC_4", do_not_clear=False)],
#     [sg.Text("Unit 4:"), sg.Input(key="UNIT_4", do_not_clear=False)],
#     [sg.Text("Quantity 4:"), sg.Input(key="QTY_4", do_not_clear=False)],
#     [sg.Text("Rate 4:"), sg.Input(key="RATE_4", do_not_clear=False)],
#     [sg.Text("Details 4:"), sg.Input(key="DETAILS_4", do_not_clear=False)],
#     ]

# fourth_layout = [
#     [sg.Text("Purpose of purchase:"), sg.Input(key="PURPOSE", do_not_clear=False)],
#     [sg.Text("Date:"), sg.Input(key="DATE", do_not_clear=False)],
#     [sg.Text("Sign:"), sg.Input(key="SIGN", do_not_clear=False)],
    
#     [sg.Text("User:"), sg.Input(key="USER", do_not_clear=False)],
#     [sg.Text("Approved by:"), sg.Input(key="APPROVED", do_not_clear=False)],
#     [sg.Text("Accountant:"), sg.Input(key="ACCOUNTANT", do_not_clear=False)],
#     ]

# tab_group = [[sg.TabGroup(
#                     [[sg.Tab('Page 1', first_layout,  element_justification= 'right'),

#                     sg.Tab('Page 2', second_layout,   element_justification= 'right'),
                    
#                     sg.Tab('Page 3', third_layout,   element_justification= 'right'),
                    
#                     sg.Tab('Page 4', fourth_layout,   element_justification= 'right')]]
#                     ), 
#                     sg.Button('Create Requisition'), sg.Button('Exit')
#                     ]]                      

