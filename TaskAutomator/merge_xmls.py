from tkinter import Variable
from xml.dom import minidom
import xml.etree.ElementTree as xml
import os.path
from canvas import Canvas
from global_instance import *

class merge_xmls():
    def __init__(self, canvas, *args, **kwargs):
        self.file_path = my_path
        self.canvas = canvas
        self.cards = self.canvas.get_cards()
    
    def read_xmls(self):
        root = xml.Element('root')
        for card in self.cards:
            path = self.file_path +'files\\'+ card + '.xml'
            if(os.path.isfile(path)):
                mydoc = minidom.parse(path)
                item = mydoc.getElementsByTagName('action')
                if(item[0].firstChild.data == "openurl"):
                    url_data = mydoc.getElementsByTagName('link')
                    browser_data = mydoc.getElementsByTagName('browser')
                    session_data = mydoc.getElementsByTagName('session')

                    print(url_data[0].firstChild.data, browser_data[0].firstChild.data, session_data[0].firstChild.data)
                    
                    action = xml.Element('Action')
                    session = xml.Element('session')
                    link = xml.Element('url')
                    browser = xml.Element('browser')
                    action.set('type', 'openurl')
                    session.text = session_data[0].firstChild.data
                    link.text = url_data[0].firstChild.data
                    browser.text = browser_data[0].firstChild.data
                    action.append(session)
                    action.append(link)
                    action.append(browser)
                    root.append(action)
                    
                if(item[0].firstChild.data == "click"):
                    xpath_data = mydoc.getElementsByTagName('xpath')
                    print(xpath_data[0].firstChild.data)
                    action = xml.Element('Action')
                    xpath = xml.Element('xpath')
                    xpath.text = xpath_data[0].firstChild.data
                    action.set('type', 'click')
                    action.append(xpath)
                    root.append(action)
                if(item[0].firstChild.data == "readtext"):
                    xpath_data = mydoc.getElementsByTagName('xpath')
                    variable_data = mydoc.getElementsByTagName('variable')
                    print(xpath_data[0].firstChild.data)
                    action = xml.Element('Action')
                    xpath = xml.Element('xpath')
                    variable_name = xml.Element('variable')
                    xpath.text = xpath_data[0].firstChild.data
                    variable_name.text = variable_data[0].firstChild.data
                    action.set('type', 'readtext')
                    action.append(xpath)
                    action.append(variable_name)
                    root.append(action)
                if(item[0].firstChild.data == "inputtext"):
                    xpath_data = mydoc.getElementsByTagName('xpath')
                    value_data = mydoc.getElementsByTagName('value')
                    print(xpath_data[0].firstChild.data)
                    action = xml.Element('Action')
                    xpath = xml.Element('xpath')
                    value = xml.Element('value')
                    xpath.text = xpath_data[0].firstChild.data
                    value.text = value_data[0].firstChild.data
                    action.set('type', 'inputtext')
                    action.append(xpath)
                    action.append(value)
                    root.append(action)
                if(item[0].firstChild.data == "messagebox"):
                    variable_data = mydoc.getElementsByTagName('variable')
                    print(variable_data[0].firstChild.data)
                    action = xml.Element('Action')
                    variable_name = xml.Element('variable')
                    variable_name.text = variable_data[0].firstChild.data
                    action.set('type', 'messagebox')
                    action.append(variable_name)
                    root.append(action)
                if(item[0].firstChild.data == "excel_open"):
                    session_data = mydoc.getElementsByTagName('session')
                    location_data = mydoc.getElementsByTagName('location')
                    filename_data = mydoc.getElementsByTagName('filename')
                    action = xml.Element('Action')
                    session = xml.Element('session')
                    location = xml.Element('location')
                    filename = xml.Element('filename')
                    session.text = session_data[0].firstChild.data
                    location.text = location_data[0].firstChild.data
                    filename.text = filename_data[0].firstChild.data
                    action.set('type', 'excel_open')
                    action.append(session)
                    action.append(location)
                    action.append(filename)
                    root.append(action)
                if(item[0].firstChild.data == "excel_create"):
                    session_data = mydoc.getElementsByTagName('session')
                    location_data = mydoc.getElementsByTagName('location')
                    filename_data = mydoc.getElementsByTagName('filename')
                    action = xml.Element('Action')
                    session = xml.Element('session')
                    location = xml.Element('location')
                    filename = xml.Element('filename')
                    session.text = session_data[0].firstChild.data
                    location.text = location_data[0].firstChild.data
                    filename.text = filename_data[0].firstChild.data
                    action.set('type', 'excel_create')
                    action.append(session)
                    action.append(location)
                    action.append(filename)
                    root.append(action)
                if(item[0].firstChild.data == "excel_save_workbook"):
                    session_data = mydoc.getElementsByTagName('session')
                    location_data = mydoc.getElementsByTagName('location')
                    filename_data = mydoc.getElementsByTagName('filename')
                    action = xml.Element('Action')
                    session = xml.Element('session')
                    location = xml.Element('location')
                    filename = xml.Element('filename')
                    session.text = session_data[0].firstChild.data
                    location.text = location_data[0].firstChild.data
                    filename.text = filename_data[0].firstChild.data
                    action.set('type', 'excel_save_workbook')
                    action.append(session)
                    action.append(location)
                    action.append(filename)
                    root.append(action)
                if(item[0].firstChild.data == "excel_delete_file"):
                    location_data = mydoc.getElementsByTagName('location')
                    filename_data = mydoc.getElementsByTagName('filename')
                    action = xml.Element('Action')
                    location = xml.Element('location')
                    filename = xml.Element('filename')
                    location.text = location_data[0].firstChild.data
                    filename.text = filename_data[0].firstChild.data
                    action.set('type', 'excel_delete_file')
                    action.append(location)
                    action.append(filename)
                    root.append(action)
                if(item[0].firstChild.data == "excel_delete_worksheet"):
                    session_data = mydoc.getElementsByTagName('session')
                    sheetname_data = mydoc.getElementsByTagName('sheetname')
                    action = xml.Element('Action')
                    session = xml.Element('session')
                    sheetname = xml.Element('sheetname')
                    session.text = session_data[0].firstChild.data
                    sheetname.text = sheetname_data[0].firstChild.data
                    action.set('type', 'excel_delete_worksheet')
                    action.append(session)
                    action.append(sheetname)
                    root.append(action)
                if(item[0].firstChild.data == "excel_create_worksheet"):
                    session_data = mydoc.getElementsByTagName('session')
                    sheetname_data = mydoc.getElementsByTagName('sheetname')
                    position_data = mydoc.getElementsByTagName('position')
                    value_data = mydoc.getElementsByTagName('value')
                    action = xml.Element('Action')
                    session = xml.Element('session')
                    sheetname = xml.Element('sheetname')
                    position = xml.Element('position')
                    value = xml.Element('value')
                    session.text = session_data[0].firstChild.data
                    sheetname.text = sheetname_data[0].firstChild.data
                    position.text = position_data[0].firstChild.data
                    value.text = value_data[0].firstChild.data
                    action.set('type', 'excel_create_worksheet')
                    action.append(session)
                    action.append(sheetname)
                    action.append(position)
                    action.append(value)
                    root.append(action)
                if(item[0].firstChild.data == "excel_set_value"):
                    session_data = mydoc.getElementsByTagName('session')
                    sheetname_data = mydoc.getElementsByTagName('sheetname')
                    cell_data = mydoc.getElementsByTagName('cell')
                    value_data = mydoc.getElementsByTagName('value')
                    action = xml.Element('Action')
                    session = xml.Element('session')
                    sheetname = xml.Element('sheetname')
                    cell = xml.Element('cell')
                    value = xml.Element('value')
                    session.text = session_data[0].firstChild.data
                    sheetname.text = sheetname_data[0].firstChild.data
                    cell.text = cell_data[0].firstChild.data
                    value.text = value_data[0].firstChild.data
                    action.set('type', 'excel_set_value')
                    action.append(session)
                    action.append(sheetname)
                    action.append(cell)
                    action.append(value)
                    root.append(action)
                if(item[0].firstChild.data == "excel_get_value"):
                    session_data = mydoc.getElementsByTagName('session')
                    sheetname_data = mydoc.getElementsByTagName('sheetname')
                    cell_data = mydoc.getElementsByTagName('cell')
                    save_to_data = mydoc.getElementsByTagName('variable')
                    action = xml.Element('Action')
                    session = xml.Element('session')
                    sheetname = xml.Element('sheetname')
                    cell = xml.Element('cell')
                    save_to = xml.Element('variable')
                    session.text = session_data[0].firstChild.data
                    sheetname.text = sheetname_data[0].firstChild.data
                    cell.text = cell_data[0].firstChild.data
                    save_to.text = save_to_data[0].firstChild.data
                    action.set('type', 'excel_get_value')
                    action.append(session)
                    action.append(sheetname)
                    action.append(cell)
                    action.append(save_to)
                    root.append(action)
                if(item[0].firstChild.data == "excel_set_formula"):
                    session_data = mydoc.getElementsByTagName('session')
                    sheetname_data = mydoc.getElementsByTagName('sheetname')
                    cell_data = mydoc.getElementsByTagName('cell')
                    formula_data = mydoc.getElementsByTagName('formula')
                    action = xml.Element('Action')
                    session = xml.Element('session')
                    sheetname = xml.Element('sheetname')
                    cell = xml.Element('cell')
                    formula = xml.Element('formula')
                    session.text = session_data[0].firstChild.data
                    sheetname.text = sheetname_data[0].firstChild.data
                    cell.text = cell_data[0].firstChild.data
                    formula.text = formula_data[0].firstChild.data
                    action.set('type', 'excel_set_formula')
                    action.append(session)
                    action.append(sheetname)
                    action.append(cell)
                    action.append(formula)
                    root.append(action)
                if(item[0].firstChild.data == "excel_merge_unmerge"):
                    session_data = mydoc.getElementsByTagName('session')
                    sheetname_data = mydoc.getElementsByTagName('sheetname')
                    merge_unmerge_data = mydoc.getElementsByTagName('merge_unmerge')
                    from_cell_data = mydoc.getElementsByTagName('from_cell')
                    to_cell_data = mydoc.getElementsByTagName('to_cell')
                    action = xml.Element('Action')
                    session = xml.Element('session')
                    sheetname = xml.Element('sheetname')
                    merge_unmerge = xml.Element('merge_unmerge')
                    from_cell = xml.Element('from_cell')
                    to_cell = xml.Element('to_cell')
                    session.text = session_data[0].firstChild.data
                    sheetname.text = sheetname_data[0].firstChild.data
                    merge_unmerge.text = merge_unmerge_data[0].firstChild.data
                    from_cell.text = from_cell_data[0].firstChild.data
                    to_cell.text = to_cell_data[0].firstChild.data
                    action.set('type', 'excel_merge_unmerge')
                    action.append(session)
                    action.append(sheetname)
                    action.append(merge_unmerge)
                    action.append(from_cell)
                    action.append(to_cell)
                    root.append(action)
                if(item[0].firstChild.data == "excel_copy_paste"):
                    copy_session_data = mydoc.getElementsByTagName('copy_session')
                    copy_sheetname_data = mydoc.getElementsByTagName('copy_sheetname')
                    copy_start_data = mydoc.getElementsByTagName('copy_start')
                    copy_end_data = mydoc.getElementsByTagName('copy_end')
                    paste_session_data = mydoc.getElementsByTagName('paste_session')
                    paste_sheetname_data = mydoc.getElementsByTagName('paste_sheetname')
                    paste_start_data = mydoc.getElementsByTagName('paste_start')
                    paste_end_data = mydoc.getElementsByTagName('paste_end')
                    action = xml.Element('Action')
                    copy_session = xml.Element('copy_session')
                    copy_sheetname = xml.Element('copy_sheetname')
                    copy_start = xml.Element('copy_start')
                    copy_end = xml.Element('copy_end')
                    paste_session = xml.Element('paste_session')
                    paste_sheetname = xml.Element('paste_sheetname')
                    paste_start = xml.Element('paste_start')
                    paste_end = xml.Element('paste_end')
                    copy_session.text = copy_session_data[0].firstChild.data
                    copy_sheetname.text = copy_sheetname_data[0].firstChild.data
                    copy_start.text = copy_start_data[0].firstChild.data
                    copy_end.text = copy_end_data[0].firstChild.data
                    paste_session.text = paste_session_data[0].firstChild.data
                    paste_sheetname.text = paste_sheetname_data[0].firstChild.data
                    paste_start.text = paste_start_data[0].firstChild.data
                    paste_end.text = paste_end_data[0].firstChild.data
                    action.set('type', 'excel_copy_paste')
                    action.append(copy_session)
                    action.append(copy_sheetname)
                    action.append(copy_start)
                    action.append(copy_end)
                    action.append(paste_session)
                    action.append(paste_sheetname)
                    action.append(paste_start)
                    action.append(paste_end)
                    root.append(action)
    
        tree = xml.ElementTree(root)   
        save_path_file = self.file_path + "testtest" + ".xml"
        xmlstr = minidom.parseString(xml.tostring(root)).toprettyxml(indent="   ")
        with open(save_path_file, "w") as f:
            f.write(xmlstr)
            
