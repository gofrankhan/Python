from xml.dom import minidom
import os.path
from global_instance import *

# parse an xml file by name
class ReadFileFromXML():
    def __init__(self, *args, **kwargs):
        self.file_path = my_path + "files\\"

    def read_xml_open_url(self, new_card):
        path = self.file_path + new_card + '.xml'
        if(os.path.isfile(path)):
            mydoc = minidom.parse(path)
            item = mydoc.getElementsByTagName('root')
            url = mydoc.getElementsByTagName('link')
            browser = mydoc.getElementsByTagName('browser')
            session = mydoc.getElementsByTagName('session')
            return url[0].firstChild.data, browser[0].firstChild.data, session[0].firstChild.data    
        return "", "Google Chrome", ""

    def read_xml_click(self, new_card):
        path = self.file_path + new_card + '.xml'
        if(os.path.isfile(path)):
            mydoc = minidom.parse(path)
            item = mydoc.getElementsByTagName('root')
            xpath = mydoc.getElementsByTagName('xpath')
            return xpath[0].firstChild.data
        return ""

    def read_xml_clear(self, new_card):
        path = self.file_path + new_card + '.xml'
        if(os.path.isfile(path)):
            mydoc = minidom.parse(path)
            item = mydoc.getElementsByTagName('root')
            xpath = mydoc.getElementsByTagName('xpath')
            return xpath[0].firstChild.data
        return ""  

    def read_xml_read_text(self, new_card):
        path = self.file_path + new_card + '.xml'
        if(os.path.isfile(path)):
            mydoc = minidom.parse(path)
            item = mydoc.getElementsByTagName('root')
            xpath = mydoc.getElementsByTagName('xpath')
            return xpath[0].firstChild.data
        return ""

    def read_xml_input_text(self, new_card):
        path = self.file_path + new_card + '.xml'
        if(os.path.isfile(path)):
            mydoc = minidom.parse(path)
            item = mydoc.getElementsByTagName('root')
            xpath = mydoc.getElementsByTagName('xpath')
            value = mydoc.getElementsByTagName('value')
            return xpath[0].firstChild.data, value[0].firstChild.data
        return "", ""

    def read_xml_excel_open(self, new_card):
        path = self.file_path + new_card + '.xml'
        if(os.path.isfile(path)):
            mydoc = minidom.parse(path)
            item = mydoc.getElementsByTagName('root')
            session = mydoc.getElementsByTagName('session')
            location = mydoc.getElementsByTagName('location')
            filename = mydoc.getElementsByTagName('filename')
            return session[0].firstChild.data, location[0].firstChild.data, filename[0].firstChild.data    
        return "", "", "" 

    def read_xml_excel_create(self, new_card):
        path = self.file_path + new_card + '.xml'
        if(os.path.isfile(path)):
            mydoc = minidom.parse(path)
            item = mydoc.getElementsByTagName('root')
            session = mydoc.getElementsByTagName('session')
            location = mydoc.getElementsByTagName('location')
            filename = mydoc.getElementsByTagName('filename')
            return session[0].firstChild.data, location[0].firstChild.data, filename[0].firstChild.data    
        return "", "", "" 
    
    def read_xml_excel_delete_file(self, new_card):
        path = self.file_path + new_card + '.xml'
        if(os.path.isfile(path)):
            mydoc = minidom.parse(path)
            item = mydoc.getElementsByTagName('root')
            location = mydoc.getElementsByTagName('location')
            filename = mydoc.getElementsByTagName('filename')
            return location[0].firstChild.data, filename[0].firstChild.data    
        return "", ""

    def read_xml_excel_create_worksheet(self, new_card):
        path = self.file_path + new_card + '.xml'
        if(os.path.isfile(path)):
            mydoc = minidom.parse(path)
            item = mydoc.getElementsByTagName('root')
            session = mydoc.getElementsByTagName('session')
            sheetname = mydoc.getElementsByTagName('sheetname')
            position = mydoc.getElementsByTagName('position')
            value = mydoc.getElementsByTagName('value')
            return session[0].firstChild.data, sheetname[0].firstChild.data, position[0].firstChild.data, value[0].firstChild.data     
        return "", "", "", ""

    def read_xml_excel_delete_worksheet(self, new_card):
        path = self.file_path + new_card + '.xml'
        if(os.path.isfile(path)):
            mydoc = minidom.parse(path)
            item = mydoc.getElementsByTagName('root')
            session = mydoc.getElementsByTagName('session')
            sheetname = mydoc.getElementsByTagName('sheetname')
            return session[0].firstChild.data, sheetname[0].firstChild.data 
        return "", ""

    def read_xml_excel_save_workbook(self, new_card):
        path = self.file_path + new_card + '.xml'
        if(os.path.isfile(path)):
            mydoc = minidom.parse(path)
            item = mydoc.getElementsByTagName('root')
            session = mydoc.getElementsByTagName('session')
            location = mydoc.getElementsByTagName('location')
            filename = mydoc.getElementsByTagName('filename')
            return session[0].firstChild.data, location[0].firstChild.data, filename[0].firstChild.data
        return "", "", ""

    def read_xml_excel_set_value(self, new_card):
        path = self.file_path + new_card + '.xml'
        if(os.path.isfile(path)):
            mydoc = minidom.parse(path)
            item = mydoc.getElementsByTagName('root')
            session = mydoc.getElementsByTagName('session')
            sheetname = mydoc.getElementsByTagName('sheetname')
            cell = mydoc.getElementsByTagName('cell')
            value = mydoc.getElementsByTagName('value')
            return session[0].firstChild.data, sheetname[0].firstChild.data, cell[0].firstChild.data, value[0].firstChild.data
        return "", "", "", ""
    
    def read_xml_excel_set_formula(self, new_card):
        path = self.file_path + new_card + '.xml'
        if(os.path.isfile(path)):
            mydoc = minidom.parse(path)
            item = mydoc.getElementsByTagName('root')
            session = mydoc.getElementsByTagName('session')
            sheetname = mydoc.getElementsByTagName('sheetname')
            cell = mydoc.getElementsByTagName('cell')
            formula = mydoc.getElementsByTagName('formula')
            return session[0].firstChild.data, sheetname[0].firstChild.data, cell[0].firstChild.data, formula[0].firstChild.data
        return "", "", "", ""

    def read_xml_excel_get_value(self, new_card):
        path = self.file_path + new_card + '.xml'
        if(os.path.isfile(path)):
            mydoc = minidom.parse(path)
            item = mydoc.getElementsByTagName('root')
            session = mydoc.getElementsByTagName('session')
            sheetname = mydoc.getElementsByTagName('sheetname')
            cell = mydoc.getElementsByTagName('cell')
            save_to = mydoc.getElementsByTagName('variable')
            return session[0].firstChild.data, sheetname[0].firstChild.data, cell[0].firstChild.data, save_to[0].firstChild.data
        return "", "", "", ""

    def read_xml_excel_merge_unmerge(self, new_card):
        path = self.file_path + new_card + '.xml'
        if(os.path.isfile(path)):
            mydoc = minidom.parse(path)
            item = mydoc.getElementsByTagName('root')
            session = mydoc.getElementsByTagName('session')
            sheetname = mydoc.getElementsByTagName('sheetname')
            merge_unmerge = mydoc.getElementsByTagName('merge_unmerge')
            from_cell = mydoc.getElementsByTagName('from_cell')
            to_cell = mydoc.getElementsByTagName('from_cell')
            return session[0].firstChild.data, sheetname[0].firstChild.data, merge_unmerge[0].firstChild.data, from_cell[0].firstChild.data, to_cell[0].firstChild.data
        return "", "", "", "", ""

    def read_xml_excel_copy_paste(self, new_card):
        path = self.file_path + new_card + '.xml'
        if(os.path.isfile(path)):
            mydoc = minidom.parse(path)
            item = mydoc.getElementsByTagName('root')
            copy_session = mydoc.getElementsByTagName('copy_session')
            copy_sheetname = mydoc.getElementsByTagName('copy_sheetname')
            copy_start = mydoc.getElementsByTagName('copy_start')
            copy_end = mydoc.getElementsByTagName('copy_end')
            paste_session = mydoc.getElementsByTagName('paste_session')
            paste_sheetname = mydoc.getElementsByTagName('paste_sheetname')
            paste_start = mydoc.getElementsByTagName('paste_start')
            paste_end = mydoc.getElementsByTagName('paste_end')
            return copy_session[0].firstChild.data, copy_sheetname[0].firstChild.data, copy_start[0].firstChild.data, copy_end[0].firstChild.data, paste_session[0].firstChild.data, paste_sheetname[0].firstChild.data, paste_start[0].firstChild.data, paste_end[0].firstChild.data
        return "", "", "", "", "", "", "", ""
                                                                   
        

