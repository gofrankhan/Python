from tkinter import *
from tkinter import ttk
from global_instance import *
import xml.etree.ElementTree as xml
from readfilesxml import ReadFileFromXML

class Properties(ttk.Frame):
    def __init__(self, parent, *args, **kwargs):
        ttk.Frame.__init__(self, parent , *args, **kwargs)
        self.parent_frame = parent
        self.path = my_path + "\\files"
        self.lbl_empty = Label(self.parent_frame, text = "Emply", height = 10, padx = 10, pady = 10)
        self.lbl_empty.grid(row = 0, column = 0, columnspan = 2 , padx = 10, pady = 10)
        self.read_file_from_xml = ReadFileFromXML()

    def reset_all(self, nameText):
        for widget in self.parent_frame.winfo_children():
            widget.destroy()
        self.lbl_empty = Label(self.parent_frame, text = nameText, height = 10)
        self.lbl_empty.grid(row = 0, column = 0, columnspan = 2, sticky=NW)
    
    def save_excel_open(self, new_card):

        root = xml.Element('root')
        session = xml.Element('session')
        location = xml.Element('location')
        filename = xml.Element('filename')
        session.text = self.entry_session.get()
        location.text = self.entry_location.get()
        filename.text = self.entry_filename.get()
        action = xml.Element('action')
        action.text = "excel_open"
        root.append(action)
        root.append(session)
        root.append(location)
        root.append(filename)
        tree = xml.ElementTree(root)
        
        from xml.dom import minidom
        save_path_file = self.path + "\\" + new_card + ".xml"
        xmlstr = minidom.parseString(xml.tostring(root)).toprettyxml(indent="   ")
        with open(save_path_file, "w") as f:
            f.write(xmlstr)

    def excel_open(self, new_card):
        self.reset_all('Open Workbook')
        self.lbl_session = Label(self.parent_frame, text="Session Name:")
        self.entry_session = Entry(self.parent_frame, width = 30)

        self.lbl_location = Label(self.parent_frame, text="Location:")
        self.entry_location = Entry(self.parent_frame, width = 30)

        self.lbl_filename = Label(self.parent_frame, text="File Name:")
        self.entry_filename= Entry(self.parent_frame, width = 30)

        self.lbl_session.grid(row = 1, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_session.grid(row = 1, column = 1, padx = 10, pady = 10, sticky=NW)
        self.lbl_location.grid(row = 2, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_location.grid(row = 2, column = 1, padx = 10, pady = 10, sticky=NW)
        self.lbl_filename.grid(row = 3, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_filename.grid(row = 3, column = 1, padx = 10, pady = 10, sticky=NW)

        btn_save = Button(self.parent_frame, text="Save", command = lambda: self.save_excel_open(new_card))
        btn_save.grid(row = 4, column = 0, padx = 10, pady = 10, sticky=W)

        session, location, filename = self.read_file_from_xml.read_xml_excel_open(new_card)
        self.entry_session.insert(0, session)
        self.entry_location.insert(0, location)
        self.entry_filename.insert(0, filename)
    
    def save_excel_create(self, new_card):

        root = xml.Element('root')
        session = xml.Element('session')
        location = xml.Element('location')
        filename = xml.Element('filename')
        session.text = self.entry_session.get()
        location.text = self.entry_location.get()
        filename.text = self.entry_filename.get()
        action = xml.Element('action')
        action.text = "excel_create"
        root.append(action)
        root.append(session)
        root.append(location)
        root.append(filename)
        tree = xml.ElementTree(root)
        
        from xml.dom import minidom
        save_path_file = self.path + "\\" + new_card + ".xml"
        xmlstr = minidom.parseString(xml.tostring(root)).toprettyxml(indent="   ")
        with open(save_path_file, "w") as f:
            f.write(xmlstr)


    def excel_create(self, new_card):
        self.reset_all('Create Workbook')
        self.lbl_session = Label(self.parent_frame, text="Session Name:")
        self.entry_session = Entry(self.parent_frame, width = 30)

        self.lbl_location = Label(self.parent_frame, text="Location:")
        self.entry_location = Entry(self.parent_frame, width = 30)

        self.lbl_filename = Label(self.parent_frame, text="File Name:")
        self.entry_filename= Entry(self.parent_frame, width = 30)

        self.lbl_session.grid(row = 1, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_session.grid(row = 1, column = 1, padx = 10, pady = 10, sticky=NW)
        self.lbl_location.grid(row = 2, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_location.grid(row = 2, column = 1, padx = 10, pady = 10, sticky=NW)
        self.lbl_filename.grid(row = 3, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_filename.grid(row = 3, column = 1, padx = 10, pady = 10, sticky=NW)

        btn_save = Button(self.parent_frame, text="Save", command = lambda: self.save_excel_create(new_card))
        btn_save.grid(row = 4, column = 0, padx = 10, pady = 10, sticky=W)
        session, location, filename = self.read_file_from_xml.read_xml_excel_create(new_card)
        self.entry_session.insert(0, session)
        self.entry_location.insert(0, location)
        self.entry_filename.insert(0, filename)
    
    def save_excel_delete_file(self, new_card):

        root = xml.Element('root')
        location = xml.Element('location')
        filename = xml.Element('filename')
        location.text = self.entry_location.get()
        filename.text = self.entry_filename.get()
        action = xml.Element('action')
        action.text = "excel_delete_file"
        root.append(action)
        root.append(location)
        root.append(filename)
        tree = xml.ElementTree(root)
        
        from xml.dom import minidom
        save_path_file = self.path + "\\" + new_card + ".xml"
        xmlstr = minidom.parseString(xml.tostring(root)).toprettyxml(indent="   ")
        with open(save_path_file, "w") as f:
            f.write(xmlstr)

    def excel_delete_file(self, new_card):
        self.reset_all('Delete Workbook')

        self.lbl_location = Label(self.parent_frame, text="Location:")
        self.entry_location = Entry(self.parent_frame, width = 30)

        self.lbl_filename = Label(self.parent_frame, text="File Name:")
        self.entry_filename= Entry(self.parent_frame, width = 30)

        self.lbl_location.grid(row = 2, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_location.grid(row = 2, column = 1, padx = 10, pady = 10, sticky=NW)
        self.lbl_filename.grid(row = 3, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_filename.grid(row = 3, column = 1, padx = 10, pady = 10, sticky=NW)

        btn_save = Button(self.parent_frame, text="Save", command = lambda: self.save_excel_delete_file(new_card))
        btn_save.grid(row = 4, column = 0, padx = 10, pady = 10, sticky=W)

        location, filename = self.read_file_from_xml.read_xml_excel_delete_file(new_card)
        self.entry_location.insert(0, location)
        self.entry_filename.insert(0, filename)
    
    def save_excel_create_worksheet(self, new_card):

        root = xml.Element('root')
        session = xml.Element('session')
        sheetname = xml.Element('sheetname')
        position = xml.Element('position')
        value = xml.Element('value')
        session.text = self.entry_session.get()
        sheetname.text = self.entry_sheet_name.get()
        if(self.var_sheet_position.get() == 'default'):
            position.text = self.var_sheet_position.get()
            value.text = '0'
        else:
            position.text = self.var_sheet_position.get()
            value.text = self.entry_sheet_position.get()
        action = xml.Element('action')
        action.text = "excel_create_worksheet"
        root.append(action)
        root.append(session)
        root.append(sheetname)
        root.append(position)
        root.append(value)
        tree = xml.ElementTree(root)
        
        from xml.dom import minidom
        save_path_file = self.path + "\\" + new_card + ".xml"
        xmlstr = minidom.parseString(xml.tostring(root)).toprettyxml(indent="   ")
        with open(save_path_file, "w") as f:
            f.write(xmlstr)
    
    def excel_create_worksheet(self, new_card):
        self.reset_all('Create Worksheet')

        self.lbl_session = Label(self.parent_frame, text="Session Name:")
        self.entry_session  = Entry(self.parent_frame, width = 30)
        self.lbl_session .grid(row = 1, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_session.grid(row = 1, column = 1, padx = 10, pady = 10, sticky=NW)

        self.lbl_sheet_name = Label(self.parent_frame, text="Sheet Name:")
        self.entry_sheet_name  = Entry(self.parent_frame, width = 30)
        self.lbl_sheet_name .grid(row = 2, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_sheet_name.grid(row = 2, column = 1, padx = 10, pady = 10, sticky=NW)

        self.var_sheet_position = StringVar()
        self.my_sheet_position_end = Checkbutton(self.parent_frame, text="Default[END]", variable=self.var_sheet_position, onvalue="default", offvalue="off", anchor = 'w')
        self.my_sheet_position_end.deselect()
        self.my_sheet_position_end.grid(row = 3, column = 0, padx = 10, pady = 10, sticky=NW)

        self.my_sheet_position = Checkbutton(self.parent_frame, text="Enter Position:", variable=self.var_sheet_position, onvalue="positional", offvalue="off", anchor = 'w')
        self.my_sheet_position.deselect()
        self.entry_sheet_position = Entry(self.parent_frame, width = 30)
        self.my_sheet_position.grid(row = 4, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_sheet_position.grid(row = 4, column = 1, padx = 10, pady = 10, sticky=NW)

        btn_save = Button(self.parent_frame, text="Save", command = lambda: self.save_excel_create_worksheet(new_card))
        btn_save.grid(row = 5, column = 0, padx = 10, pady = 10, sticky=W)

        session, sheetname, position, value = self.read_file_from_xml.read_xml_excel_create_worksheet(new_card)
        self.entry_session.insert(0, session)
        self.entry_sheet_name.insert(0, sheetname)
        if(position == 'default'):
            self.my_sheet_position_end.select()
        elif(position == 'positional'):
            self.my_sheet_position.select()
            self.entry_sheet_position.insert(0, value)
    
    def save_excel_delete_worksheet(self, new_card):

        root = xml.Element('root')
        session = xml.Element('session')
        sheetname = xml.Element('sheetname')
        session.text = self.entry_session.get()
        sheetname.text = self.entry_sheet_name.get()
        action = xml.Element('action')
        action.text = "excel_delete_worksheet"
        root.append(action)
        root.append(session)
        root.append(sheetname)
        tree = xml.ElementTree(root)
        
        from xml.dom import minidom
        save_path_file = self.path + "\\" + new_card + ".xml"
        xmlstr = minidom.parseString(xml.tostring(root)).toprettyxml(indent="   ")
        with open(save_path_file, "w") as f:
            f.write(xmlstr)

    def excel_delete_worksheet(self, new_card):
        self.reset_all('Delete Worksheet')

        self.lbl_session = Label(self.parent_frame, text="Session:")
        self.entry_session = Entry(self.parent_frame, width = 30)

        self.lbl_sheet_name = Label(self.parent_frame, text="Sheet Name:")
        self.entry_sheet_name= Entry(self.parent_frame, width = 30)

        self.lbl_session.grid(row = 2, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_session.grid(row = 2, column = 1, padx = 10, pady = 10, sticky=NW)
        self.lbl_sheet_name.grid(row = 3, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_sheet_name.grid(row = 3, column = 1, padx = 10, pady = 10, sticky=NW)

        btn_save = Button(self.parent_frame, text="Save", command = lambda: self.save_excel_delete_worksheet(new_card))
        btn_save.grid(row = 4, column = 0, padx = 10, pady = 10, sticky=W)

        session, sheetname = self.read_file_from_xml.read_xml_excel_delete_worksheet(new_card)
        self.entry_session.insert(0, session)
        self.entry_sheet_name.insert(0, sheetname)

    def save_excel_save_workbook(self, new_card):

        root = xml.Element('root')
        session = xml.Element('session')
        location = xml.Element('location')
        filename = xml.Element('filename')
        session.text = self.entry_session.get()
        location.text = self.entry_location.get()
        filename.text = self.entry_filename.get()
        action = xml.Element('action')
        action.text = "excel_save_workbook"
        root.append(action)
        root.append(session)
        root.append(location)
        root.append(filename)
        tree = xml.ElementTree(root)
        
        from xml.dom import minidom
        save_path_file = self.path + "\\" + new_card + ".xml"
        xmlstr = minidom.parseString(xml.tostring(root)).toprettyxml(indent="   ")
        with open(save_path_file, "w") as f:
            f.write(xmlstr)

    def excel_save_workbook(self, new_card):
        self.reset_all('Save Workbook')
        self.lbl_session = Label(self.parent_frame, text="Session Name:")
        self.entry_session = Entry(self.parent_frame, width = 30)

        self.lbl_location = Label(self.parent_frame, text="Location:")
        self.entry_location = Entry(self.parent_frame, width = 30)

        self.lbl_filename = Label(self.parent_frame, text="File Name:")
        self.entry_filename= Entry(self.parent_frame, width = 30)

        self.lbl_session.grid(row = 1, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_session.grid(row = 1, column = 1, padx = 10, pady = 10, sticky=NW)
        self.lbl_location.grid(row = 2, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_location.grid(row = 2, column = 1, padx = 10, pady = 10, sticky=NW)
        self.lbl_filename.grid(row = 3, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_filename.grid(row = 3, column = 1, padx = 10, pady = 10, sticky=NW)

        btn_save = Button(self.parent_frame, text="Save", command = lambda: self.save_excel_save_workbook(new_card))
        btn_save.grid(row = 4, column = 0, padx = 10, pady = 10, sticky=W)

        session, location, filename = self.read_file_from_xml.read_xml_excel_save_workbook(new_card)
        self.entry_session.insert(0, session)
        self.entry_location.insert(0, location)
        self.entry_filename.insert(0, filename)


    def save_excel_set_value(self, new_card):

        root = xml.Element('root')
        session = xml.Element('session')
        sheetname = xml.Element('sheetname')
        cell = xml.Element('cell')
        value = xml.Element('value')
        session.text = self.entry_session.get()
        sheetname.text = self.entry_sheet_name.get()
        cell.text = self.entry_cell.get()
        value.text = self.entry_value.get()
        action = xml.Element('action')
        action.text = "excel_set_value"
        root.append(action)
        root.append(session)
        root.append(sheetname)
        root.append(cell)
        root.append(value)
        tree = xml.ElementTree(root)
        
        from xml.dom import minidom
        save_path_file = self.path + "\\" + new_card + ".xml"
        xmlstr = minidom.parseString(xml.tostring(root)).toprettyxml(indent="   ")
        with open(save_path_file, "w") as f:
            f.write(xmlstr)

    def excel_set_value(self, new_card):
        self.reset_all('Set Value')
        self.lbl_session = Label(self.parent_frame, text="Session Name:")
        self.entry_session = Entry(self.parent_frame, width = 30)

        self.lbl_sheet_name = Label(self.parent_frame, text="Sheet Name:")
        self.entry_sheet_name = Entry(self.parent_frame, width = 30)

        self.lbl_value = Label(self.parent_frame, text="Value:")
        self.entry_value = Entry(self.parent_frame, width = 30)

        self.lbl_cell = Label(self.parent_frame, text="Cell:")
        self.entry_cell = Entry(self.parent_frame, width = 30)


        self.lbl_session.grid(row = 1, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_session.grid(row = 1, column = 1, padx = 10, pady = 10, sticky=NW)
        self.lbl_sheet_name.grid(row = 2, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_sheet_name.grid(row = 2, column = 1, padx = 10, pady = 10, sticky=NW)
        self.lbl_value.grid(row = 3, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_value.grid(row = 3, column = 1, padx = 10, pady = 10, sticky=NW)
        self.lbl_cell.grid(row = 4, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_cell.grid(row = 4, column = 1, padx = 10, pady = 10, sticky=NW)

        btn_save = Button(self.parent_frame, text="Save", command = lambda: self.save_excel_set_value(new_card))
        btn_save.grid(row = 5, column = 0, padx = 10, pady = 10, sticky=W)

        session, sheetname, cell, value = self.read_file_from_xml.read_xml_excel_set_value(new_card)
        self.entry_session.insert(0, session)
        self.entry_sheet_name.insert(0, sheetname)
        self.entry_cell.insert(0, cell)
        self.entry_value.insert(0, value)
    
    def save_excel_set_formula(self, new_card):

        root = xml.Element('root')
        session = xml.Element('session')
        sheetname = xml.Element('sheetname')
        cell = xml.Element('cell')
        formula = xml.Element('formula')
        session.text = self.entry_session.get()
        sheetname.text = self.entry_sheet_name.get()
        cell.text = self.entry_cell.get()
        formula.text = self.entry_formula.get()
        action = xml.Element('action')
        action.text = "excel_set_formula"
        root.append(action)
        root.append(session)
        root.append(sheetname)
        root.append(cell)
        root.append(formula)
        tree = xml.ElementTree(root)
        
        from xml.dom import minidom
        save_path_file = self.path + "\\" + new_card + ".xml"
        xmlstr = minidom.parseString(xml.tostring(root)).toprettyxml(indent="   ")
        with open(save_path_file, "w") as f:
            f.write(xmlstr)
    


        
