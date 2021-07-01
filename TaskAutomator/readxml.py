from xml.dom import minidom
from my_selenium import MySelenium
from excel_test import MyExcel
from global_instance import *

# parse an xml file by name
mydoc = minidom.parse(my_path + 'testtest.xml')

items = mydoc.getElementsByTagName('Action')
sel = None

# all item attributes
for elem in items:
    actionName = elem.attributes['type'].value
    if(actionName == 'openurl'):
        url = elem.getElementsByTagName('url')
        browser = elem.getElementsByTagName('browser')
        sel = MySelenium()
        MySelenium.openurl(sel, url[0].firstChild.data, browser[0].firstChild.data)
    elif(actionName == 'click'):
        xpath = elem.getElementsByTagName('xpath')
        if(len(xpath) > 0):
            MySelenium.click_by_xpath(sel, xpath[0].firstChild.data)
        id = elem.getElementsByTagName('id')
        if(len(id) > 0):
            MySelenium.click_by_id(sel, id[0].firstChild.data)
        name = elem.getElementsByTagName('name')
        if(len(name) > 0):
            MySelenium.click_by_name(sel, name[0].firstChild.data)
        class_name = elem.getElementsByTagName('class_name')
        if(len(class_name) > 0):
            MySelenium.click_by_class_name(sel, class_name[0].firstChild.data)
        css_selector = elem.getElementsByTagName('css_selector')
        if(len(css_selector) > 0):
            MySelenium.click_by_css_selector(sel, css_selector[0].firstChild.data)
        link_text = elem.getElementsByTagName('link_text')
        if(len(link_text) > 0):
            MySelenium.click_by_link_text(sel, link_text[0].firstChild.data)
        partial_link_text = elem.getElementsByTagName('partial_link_text')
        if(len(partial_link_text) > 0):
            MySelenium.click_by_partial_link_text(sel, partial_link_text[0].firstChild.data)
    elif(actionName == 'inputtext'):
        xpath = elem.getElementsByTagName('xpath')
        value = elem.getElementsByTagName('value')
        if(len(xpath) > 0):
            MySelenium.inputtext_by_xpath(sel, xpath[0].firstChild.data, value[0].firstChild.data)
        id = elem.getElementsByTagName('id')
        if(len(id) > 0):
            MySelenium.inputtext_by_id(sel, id[0].firstChild.data, value[0].firstChild.data)
        name = elem.getElementsByTagName('name')
        if(len(name) > 0):
            MySelenium.inputtext_by_name(sel, name[0].firstChild.data, value[0].firstChild.data)
        class_name = elem.getElementsByTagName('class_name')
        if(len(class_name) > 0):
            MySelenium.inputtext_by_class_name(sel, class_name[0].firstChild.data, value[0].firstChild.data)
        css_selector = elem.getElementsByTagName('css_selector')
        if(len(css_selector) > 0):
            MySelenium.inputtext_by_css_selector(sel, css_selector[0].firstChild.data, value[0].firstChild.data)
        link_text = elem.getElementsByTagName('link_text')
        if(len(link_text) > 0):
            MySelenium.inputtext_by_link_text(sel, link_text[0].firstChild.data, value[0].firstChild.data)
        partial_link_text = elem.getElementsByTagName('partial_link_text')
        if(len(partial_link_text) > 0):
            MySelenium.inputtext_by_partial_link_text(sel, partial_link_text[0].firstChild.data, value[0].firstChild.data)
    elif(actionName == 'readtext'):
        xpath = elem.getElementsByTagName('xpath')
        save_to = elem.getElementsByTagName('variable')
        MySelenium.readtext_by_xpath(sel, xpath[0].firstChild.data, save_to[0].firstChild.data)
    elif(actionName == 'messagebox'):
        variable_name = elem.getElementsByTagName('variable')
        MySelenium.my_messagebox(sel, variable_name[0].firstChild.data)
    elif(actionName == 'clear'):
        xpath = elem.getElementsByTagName('xpath')
        MySelenium.clear_by_xpath(sel, xpath[0].firstChild.data)
    elif(actionName == 'sendkeys'):
        key_value = elem.getElementsByTagName('key_value')
        xpath = elem.getElementsByTagName('xpath')
        if(len(xpath) > 0):
            MySelenium.send_keys_by_xpath(sel, xpath[0].firstChild.data, key_value[0].firstChild.data)
        id = elem.getElementsByTagName('id')
        if(len(id) > 0):
            MySelenium.send_keys_by_id(sel, id[0].firstChild.data, key_value[0].firstChild.data)
        name = elem.getElementsByTagName('name')
        if(len(name) > 0):
            MySelenium.send_keys_by_name(sel, name[0].firstChild.data, key_value[0].firstChild.data)
    elif(actionName == 'wait'):
        seconds = elem.getElementsByTagName('seconds')
        MySelenium.wait(sel, int(seconds[0].firstChild.data))
    elif(actionName == 'excel_create'):
        session = elem.getElementsByTagName('session')
        location = elem.getElementsByTagName('location')
        filename = elem.getElementsByTagName('filename')
        MyExcel.excel_create(session[0].firstChild.data, location[0].firstChild.data, filename[0].firstChild.data)
    elif(actionName == 'excel_open'):
        session = elem.getElementsByTagName('session')
        location = elem.getElementsByTagName('location')
        filename = elem.getElementsByTagName('filename')
        MyExcel.excel_open(session[0].firstChild.data, location[0].firstChild.data, filename[0].firstChild.data)
    elif(actionName == 'excel_delete_file'):
        location = elem.getElementsByTagName('location')
        filename = elem.getElementsByTagName('filename')
        MyExcel.excel_delete_file(location[0].firstChild.data, filename[0].firstChild.data)
    elif(actionName == 'excel_create_worksheet'):
        session = elem.getElementsByTagName('session')
        sheetname = elem.getElementsByTagName('sheetname')
        position = elem.getElementsByTagName('position')
        value = elem.getElementsByTagName('value')
        MyExcel.excel_create_worksheet(session[0].firstChild.data, sheetname[0].firstChild.data, position[0].firstChild.data, value[0].firstChild.data)
    elif(actionName == 'excel_delete_worksheet'):
        session = elem.getElementsByTagName('session')
        sheetname = elem.getElementsByTagName('sheetname')
        MyExcel.excel_delete_worksheet(session[0].firstChild.data, sheetname[0].firstChild.data)
    elif(actionName == 'excel_save_workbook'):
        session = elem.getElementsByTagName('session')
        location = elem.getElementsByTagName('location')
        filename = elem.getElementsByTagName('filename')
        MyExcel.excel_save_workbook(session[0].firstChild.data, location[0].firstChild.data, filename[0].firstChild.data)
    elif(actionName == 'excel_set_value'):
        print('hello')
        session = elem.getElementsByTagName('session')
        sheetname = elem.getElementsByTagName('sheetname')
        cell = elem.getElementsByTagName('cell')
        value = elem.getElementsByTagName('value')
        MyExcel.excel_set_value(session[0].firstChild.data, sheetname[0].firstChild.data, cell[0].firstChild.data, value[0].firstChild.data)
    elif(actionName == 'excel_get_value'):
        session = elem.getElementsByTagName('session')
        sheetname = elem.getElementsByTagName('sheetname')
        cell = elem.getElementsByTagName('cell')
        variable = elem.getElementsByTagName('variable')
        MyExcel.excel_get_value(session[0].firstChild.data, sheetname[0].firstChild.data, cell[0].firstChild.data, variable[0].firstChild.data)
    elif(actionName == 'excel_set_formula'):
        session = elem.getElementsByTagName('session')
        sheetname = elem.getElementsByTagName('sheetname')
        cell = elem.getElementsByTagName('cell')
        formula = elem.getElementsByTagName('formula')
        MyExcel.excel_get_value(session[0].firstChild.data, sheetname[0].firstChild.data, cell[0].firstChild.data, formula[0].firstChild.data)
    elif(actionName == 'excel_merge_unmerge'):
        session = elem.getElementsByTagName('session')
        sheetname = elem.getElementsByTagName('sheetname')
        merge_unmerge = elem.getElementsByTagName('merge_unmerge')
        from_cell = elem.getElementsByTagName('from_cell')
        to_cell = elem.getElementsByTagName('to_cell')
        MyExcel.excel_merge_unmerge(session[0].firstChild.data, sheetname[0].firstChild.data, merge_unmerge[0].firstChild.data, from_cell[0].firstChild.data , to_cell[0].firstChild.data)
    elif(actionName == 'excel_copy_paste'):
        copy_session = elem.getElementsByTagName('copy_session')
        copy_sheetname = elem.getElementsByTagName('copy_sheetname')
        copy_start = elem.getElementsByTagName('copy_start')
        copy_end = elem.getElementsByTagName('copy_end')
        paste_session = elem.getElementsByTagName('paste_session')
        paste_sheetname = elem.getElementsByTagName('paste_sheetname')
        paste_start = elem.getElementsByTagName('paste_start')
        paste_end = elem.getElementsByTagName('paste_end')
        MyExcel.excel_copy_paste(copy_session[0].firstChild.data, copy_sheetname[0].firstChild.data, copy_start[0].firstChild.data, copy_end[0].firstChild.data, paste_session[0].firstChild.data, paste_sheetname[0].firstChild.data, paste_start[0].firstChild.data, paste_end[0].firstChild.data)
    

        
        

