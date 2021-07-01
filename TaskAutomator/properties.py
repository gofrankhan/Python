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
    

    def excel_set_formula(self, new_card):
        self.reset_all('Set Formula')
        self.lbl_session = Label(self.parent_frame, text="Session Name:")
        self.entry_session = Entry(self.parent_frame, width = 30)

        self.lbl_sheet_name = Label(self.parent_frame, text="Sheet Name:")
        self.entry_sheet_name = Entry(self.parent_frame, width = 30)

        self.lbl_formula = Label(self.parent_frame, text="Formula:")
        self.entry_formula = Entry(self.parent_frame, width = 30)

        self.lbl_cell = Label(self.parent_frame, text="Cell:")
        self.entry_cell = Entry(self.parent_frame, width = 30)


        self.lbl_session.grid(row = 1, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_session.grid(row = 1, column = 1, padx = 10, pady = 10, sticky=NW)
        self.lbl_sheet_name.grid(row = 2, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_sheet_name.grid(row = 2, column = 1, padx = 10, pady = 10, sticky=NW)
        self.lbl_formula.grid(row = 3, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_formula.grid(row = 3, column = 1, padx = 10, pady = 10, sticky=NW)
        self.lbl_cell.grid(row = 4, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_cell.grid(row = 4, column = 1, padx = 10, pady = 10, sticky=NW)

        btn_save = Button(self.parent_frame, text="Save", command = lambda: self.save_excel_set_formula(new_card))
        btn_save.grid(row = 5, column = 0, padx = 10, pady = 10, sticky=W)

        session, sheetname, cell, formula = self.read_file_from_xml.read_xml_excel_set_formula(new_card)
        self.entry_session.insert(0, session)
        self.entry_sheet_name.insert(0, sheetname)
        self.entry_cell.insert(0, cell)
        self.entry_formula.insert(0, formula)
    
    def save_excel_merge_unmerge(self, new_card):

        root = xml.Element('root')
        session = xml.Element('session')
        sheetname = xml.Element('sheetname')
        merge_unmerge = xml.Element('merge_unmerge')
        from_cell = xml.Element('from_cell')
        to_cell = xml.Element('to_cell')
        session.text = self.entry_session.get()
        sheetname.text = self.entry_sheet_name.get()
        if(self.var_merge_unmerge.get() == 'merge'):
            merge_unmerge.text = 'merge'
        elif(self.var_merge_unmerge.get() == 'unmerge'):
            merge_unmerge.text = 'unmerge'
        from_cell.text = self.entry_from_cell.get()
        to_cell.text = self.entry_to_cell.get()
        action = xml.Element('action')
        action.text = "excel_merge_unmerge"
        root.append(action)
        root.append(session)
        root.append(sheetname)
        root.append(merge_unmerge)
        root.append(from_cell)
        root.append(to_cell)
        tree = xml.ElementTree(root)
        
        from xml.dom import minidom
        save_path_file = self.path + "\\" + new_card + ".xml"
        xmlstr = minidom.parseString(xml.tostring(root)).toprettyxml(indent="   ")
        with open(save_path_file, "w") as f:
            f.write(xmlstr)
    
    def excel_merge_unmerge(self, new_card):
        self.reset_all('Merge Unmerge')
        self.lbl_session = Label(self.parent_frame, text="Session Name:")
        self.entry_session = Entry(self.parent_frame, width = 30)

        self.lbl_sheet_name = Label(self.parent_frame, text="Sheet Name:")
        self.entry_sheet_name = Entry(self.parent_frame, width = 30)

        self.var_merge_unmerge = StringVar()
        self.my_merge = Checkbutton(self.parent_frame, text="Merge", variable=self.var_merge_unmerge, onvalue="merge", offvalue="off", anchor = 'w')
        self.my_merge.deselect()
        self.my_unmerge = Checkbutton(self.parent_frame, text="Unmerge", variable=self.var_merge_unmerge, onvalue="unmerge", offvalue="off", anchor = 'w')
        self.my_unmerge.deselect()

        self.lbl_from_cell = Label(self.parent_frame, text="From Cell:")
        self.entry_from_cell = Entry(self.parent_frame, width = 30)

        self.lbl_to_cell = Label(self.parent_frame, text="To Cell:")
        self.entry_to_cell = Entry(self.parent_frame, width = 30)


        self.lbl_session.grid(row = 1, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_session.grid(row = 1, column = 1, padx = 10, pady = 10, sticky=NW)
        self.lbl_sheet_name.grid(row = 2, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_sheet_name.grid(row = 2, column = 1, padx = 10, pady = 10, sticky=NW)
        self.my_merge.grid(row = 3, column = 0, padx = 10, pady = 10, sticky=NW)
        self.my_unmerge.grid(row = 3, column = 1, padx = 10, pady = 10, sticky=NW)
        self.lbl_from_cell.grid(row = 4, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_from_cell.grid(row = 4, column = 1, padx = 10, pady = 10, sticky=NW)
        self.lbl_to_cell.grid(row = 5, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_to_cell.grid(row = 5, column = 1, padx = 10, pady = 10, sticky=NW)

        btn_save = Button(self.parent_frame, text="Save", command = lambda: self.save_excel_merge_unmerge(new_card))
        btn_save.grid(row = 9, column = 0, padx = 10, pady = 10, sticky=W)

        session, sheetname, merge_unmerge, from_cell, to_cell = self.read_file_from_xml.read_xml_excel_merge_unmerge(new_card)
        self.entry_session.insert(0, session)
        self.entry_sheet_name.insert(0, sheetname)
        if(merge_unmerge == "merge"):
            self.my_merge.select()
        elif(merge_unmerge == "unmerge"):
            self.my_unmerge.select()
        self.entry_from_cell.insert(0, from_cell)
        self.entry_to_cell.insert(0, to_cell)

    def save_excel_copy_paste(self, new_card):

        root = xml.Element('root')
        copy_session = xml.Element('copy_session')
        copy_sheetname = xml.Element('copy_sheetname')
        copy_start = xml.Element('copy_start')
        copy_end = xml.Element('copy_end')
        paste_session = xml.Element('paste_session')
        paste_sheetname = xml.Element('paste_sheetname')
        paste_start = xml.Element('paste_start')
        paste_end = xml.Element('paste_end')
        copy_session.text = self.entry_copy_session.get()
        copy_sheetname.text = self.entry_copy_sheet_name.get()
        copy_start.text = self.entry_copy_start.get()
        copy_end.text = self.entry_copy_end.get()
        paste_session.text = self.entry_paste_session.get()
        paste_sheetname.text = self.entry_paste_sheet_name.get()
        paste_start.text = self.entry_paste_start.get()
        paste_end.text = self.entry_paste_end.get()
        action = xml.Element('action')
        action.text = "excel_copy_paste"
        root.append(action)
        root.append(copy_session)
        root.append(copy_sheetname)
        root.append(copy_start)
        root.append(copy_end)
        root.append(paste_session)
        root.append(paste_sheetname)
        root.append(paste_start)
        root.append(paste_end)
        tree = xml.ElementTree(root)
        
        from xml.dom import minidom
        save_path_file = self.path + "\\" + new_card + ".xml"
        xmlstr = minidom.parseString(xml.tostring(root)).toprettyxml(indent="   ")
        with open(save_path_file, "w") as f:
            f.write(xmlstr)

    def excel_copy_paste(self, new_card):
        self.reset_all('Copy Paste')
        self.lbl_copy_session = Label(self.parent_frame, text="Copy Session Name:")
        self.entry_copy_session = Entry(self.parent_frame, width = 30)

        self.lbl_copy_sheet_name = Label(self.parent_frame, text="Copy Sheet Name:")
        self.entry_copy_sheet_name = Entry(self.parent_frame, width = 30)

        self.lbl_copy_start = Label(self.parent_frame, text="Copy Start:")
        self.entry_copy_start = Entry(self.parent_frame, width = 30)

        self.lbl_copy_end = Label(self.parent_frame, text="Copy End:")
        self.entry_copy_end = Entry(self.parent_frame, width = 30)

        self.lbl_paste_session = Label(self.parent_frame, text="Paste Session Name:")
        self.entry_paste_session = Entry(self.parent_frame, width = 30)

        self.lbl_paste_sheet_name = Label(self.parent_frame, text="Paste Sheet Name:")
        self.entry_paste_sheet_name = Entry(self.parent_frame, width = 30)

        self.lbl_paste_start = Label(self.parent_frame, text="Paste Start:")
        self.entry_paste_start = Entry(self.parent_frame, width = 30)

        self.lbl_paste_end = Label(self.parent_frame, text="Paste End:")
        self.entry_paste_end = Entry(self.parent_frame, width = 30)


        self.lbl_copy_session.grid(row = 1, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_copy_session.grid(row = 1, column = 1, padx = 10, pady = 10, sticky=NW)
        self.lbl_copy_sheet_name.grid(row = 2, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_copy_sheet_name.grid(row = 2, column = 1, padx = 10, pady = 10, sticky=NW)
        self.lbl_copy_start.grid(row = 3, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_copy_start.grid(row = 3, column = 1, padx = 10, pady = 10, sticky=NW)
        self.lbl_copy_end.grid(row = 4, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_copy_end.grid(row = 4, column = 1, padx = 10, pady = 10, sticky=NW)
        self.lbl_paste_session.grid(row = 5, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_paste_session.grid(row = 5, column = 1, padx = 10, pady = 10, sticky=NW)
        self.lbl_paste_sheet_name.grid(row = 6, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_paste_sheet_name.grid(row = 6, column = 1, padx = 10, pady = 10, sticky=NW)
        self.lbl_paste_start.grid(row = 7, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_paste_start.grid(row = 7, column = 1, padx = 10, pady = 10, sticky=NW)
        self.lbl_paste_end.grid(row = 8, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_paste_end.grid(row = 8, column = 1, padx = 10, pady = 10, sticky=NW)

        btn_save = Button(self.parent_frame, text="Save", command = lambda: self.save_excel_copy_paste(new_card))
        btn_save.grid(row = 9, column = 0, padx = 10, pady = 10, sticky=W)

        copy_session, copy_sheetname, copy_start, copy_end, paste_session, paste_sheetname, paste_start, paste_end = self.read_file_from_xml.read_xml_excel_copy_paste(new_card)
        self.entry_copy_session.insert(0, copy_session)
        self.entry_copy_sheet_name.insert(0, copy_sheetname)
        self.entry_copy_start.insert(0, copy_session)
        self.entry_copy_end.insert(0, copy_sheetname)
        self.entry_paste_session.insert(0, paste_session)
        self.entry_paste_sheet_name.insert(0, paste_sheetname)
        self.entry_paste_start.insert(0, paste_session)
        self.entry_paste_end.insert(0, paste_sheetname)

    def save_excel_get_value(self, new_card):

        root = xml.Element('root')
        session = xml.Element('session')
        sheetname = xml.Element('sheetname')
        cell = xml.Element('cell')
        save_to = xml.Element('variable')
        session.text = self.entry_session.get()
        sheetname.text = self.entry_sheet_name.get()
        cell.text = self.entry_cell.get()
        save_to.text = self.myVar.get()
        action = xml.Element('action')
        action.text = "excel_get_value"
        root.append(action)
        root.append(session)
        root.append(sheetname)
        root.append(cell)
        root.append(save_to)
        tree = xml.ElementTree(root)
        
        from xml.dom import minidom
        save_path_file = self.path + "\\" + new_card + ".xml"
        xmlstr = minidom.parseString(xml.tostring(root)).toprettyxml(indent="   ")
        with open(save_path_file, "w") as f:
            f.write(xmlstr)
    
    def excel_get_value(self, new_card):
        self.reset_all('Get Value')
        self.lbl_session = Label(self.parent_frame, text="Session Name:")
        self.entry_session = Entry(self.parent_frame, width = 30)

        self.lbl_sheet_name = Label(self.parent_frame, text="Sheet Name:")
        self.entry_sheet_name = Entry(self.parent_frame, width = 30)

        self.lbl_cell = Label(self.parent_frame, text="Cell:")
        self.entry_cell = Entry(self.parent_frame, width = 30)

        self.lbl_save_to = Label(self.parent_frame, text="Save To:")
        self.myVar = StringVar()
        self.my_variable = ttk.Combobox(self.parent_frame, width = 26, textvariable = self.myVar)
        self.my_variable['values'] = ('prompt-assignment')
        self.my_variable.current(0)


        self.lbl_session.grid(row = 1, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_session.grid(row = 1, column = 1, padx = 10, pady = 10, sticky=NW)
        self.lbl_sheet_name.grid(row = 2, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_sheet_name.grid(row = 2, column = 1, padx = 10, pady = 10, sticky=NW)
        self.lbl_cell.grid(row = 3, column = 0, padx = 10, pady = 10, sticky=NW)
        self.entry_cell.grid(row = 3, column = 1, padx = 10, pady = 10, sticky=NW)
        self.lbl_save_to.grid(row = 4, column = 0, padx = 10, pady = 10, sticky=NW)
        self.my_variable.grid(row = 4, column = 1, padx = 10, pady = 10, sticky=NW)

        btn_save = Button(self.parent_frame, text="Save", command = lambda: self.save_excel_get_value(new_card))
        btn_save.grid(row = 5, column = 0, padx = 10, pady = 10, sticky=W)

        session, sheetname, cell, save_to = self.read_file_from_xml.read_xml_excel_get_value(new_card)
        self.entry_session.insert(0, session)
        self.entry_sheet_name.insert(0, sheetname)
        self.entry_cell.insert(0, cell)
        self.my_variable.set(save_to)

    
    def save_open_url(self, new_card):

        root = xml.Element('root')
        session = xml.Element('session')
        link = xml.Element('link')
        browser = xml.Element('browser')
        session.text = self.entry_session.get()
        link.text = self.entry_link.get()
        browser.text = self.clicked.get()
        action = xml.Element('action')
        action.text = "openurl"
        root.append(action)
        root.append(session)
        root.append(link)
        root.append(browser)
        tree = xml.ElementTree(root)
        
        from xml.dom import minidom
        save_path_file = self.path + "\\" + new_card + ".xml"
        xmlstr = minidom.parseString(xml.tostring(root)).toprettyxml(indent="   ")
        with open(save_path_file, "w") as f:
            f.write(xmlstr)

    def open_url(self, new_card):
        self.reset_all('Open Url')
        self.lbl_session = Label(self.parent_frame, text="Session Name:")
        self.entry_session = Entry(self.parent_frame, width = 30)

        self.lbl_link = Label(self.parent_frame, text="URL:")
        self.entry_link = Entry(self.parent_frame, width = 30)

        options = [
            "Google Chrome",
            "Mozilla Firefox",
            "Internet Explorer",
            "Internet Edge",
            "Opera"
        ]

        self.clicked = StringVar()
        self.clicked.set(options[0])

        self.lbl_browser = Label(self.parent_frame, text="Select Browser:")
        self.entry_browser= OptionMenu(self.parent_frame, self.clicked, *options)  

        btn_save = Button(self.parent_frame, text="Save", command = lambda: self.save_open_url(new_card))

        self.lbl_session.grid(row = 1, column = 0, padx = 10, pady = 10, sticky=W)
        self.entry_session.grid(row = 1, column = 1, padx = 10, pady = 10, sticky=W)
        self.lbl_link.grid(row = 2, column = 0, padx = 10, pady = 10, sticky=W)
        self.entry_link.grid(row = 2, column = 1, padx = 10, pady = 10, sticky=W)
        self.lbl_browser.grid(row = 3, column = 0, padx = 10, pady = 10, sticky=W)
        self.entry_browser.grid(row = 3, column = 1, padx = 10, pady = 10, sticky=W)
        btn_save.grid(row = 4, column = 0, padx = 10, pady = 10, sticky=W)
        url, browser, session = self.read_file_from_xml.read_xml_open_url(new_card)
        self.entry_link.insert(0, url)
        self.entry_session.insert(0, session)
        self.clicked.set(browser)


        
