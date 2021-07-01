import tkinter as tk
from tkinter import ttk
import os
from global_instance import *

class Canvas():

    def __init__(self, properties, *args, **kwargs):
        self.parent_frame = None
        self.canvas_frame = None
        self.my_canvas = None
        self.xscroll = None
        self.yscroll = None
        self.cards = []
        self.cards_dict = {}
        self.var_dict = {}
        self.card_number = 0
        self.new_card = None
        self.select_x = 0
        self.select_y = 0
        self.normal_font = ("Normal")
        self.properties = properties

    def create_Ui(self, parent):
        self.parent_frame = parent
        self.canvas_frame = ttk.Frame(self.parent_frame, height = 950, width = 640)
        self.canvas_frame.pack(expand = True, fill = tk.BOTH, side = tk.TOP)
        self.my_canvas = tk.Canvas(self.canvas_frame, width = 780, height = 450, background = 'white', scrollregion = (0,0,2200, 4380))
        self.my_canvas.grid(row = 0, column = 0, columnspan=30, rowspan=30, sticky = 'nsew')

        self.xscroll = ttk.Scrollbar(self.canvas_frame, orient = tk.HORIZONTAL, command = self.my_canvas.xview)
        self.yscroll = ttk.Scrollbar(self.canvas_frame, orient = tk.VERTICAL, command = self.my_canvas.yview)
        self.my_canvas.config(xscrollcommand = self.xscroll.set, yscrollcommand = self.yscroll.set)
        self.xscroll.grid(row = 31, column = 0, columnspan = 30, sticky = 'ew')
        self.yscroll.grid(row = 1, column = 31,  rowspan = 30, sticky = 'ns')
        self.my_canvas.bind('<Button-3>', self.do_popup)

        self.right_click_menu = tk.Menu(self.my_canvas, bg="lightgrey", fg="black", tearoff=0)
        self.right_click_menu.add_command(label='Select', command=self.onclick)
        self.right_click_menu.add_command(label='Delete', command=self.delete)


    def create_card(self, card_title):
        self.new_card = 'card' + str(self.card_number)
        self.cards.append(self.new_card)
        self.cards_dict[self.new_card] = card_title
        self.card_number += 1
        self.my_canvas.create_rectangle(50, 20, 200, 100, fill='#dfcdaa', tags=[self.new_card])
        self.my_canvas.create_line(50,45, 200, 45, tags=[self.new_card])
        self.my_canvas.create_text(55,40, text=card_title, anchor = 'sw', tags=[self.new_card])
        self.my_canvas.create_text(55,50, text="Details of the card \nIt may have \nmultiple lines", anchor = 'nw', tags=[self.new_card])
        #self.my_canvas.move(a, 20, 20)
        self.my_canvas.tag_bind(self.new_card, '<B1-Motion>', self.move)
        self.my_canvas.tag_bind(self.new_card, '<ButtonPress-1>', self.savePosn)
        if(card_title == 'Open URL'):
            self.properties.open_url(self.new_card)
        elif(card_title == 'Click'):
            self.properties.click(self.new_card)
        elif(card_title == 'Clear'):
            self.properties.clear(self.new_card)
        elif(card_title == 'Read Text'):
            self.properties.read_text(self.new_card)
        elif(card_title == "Input Text"):
            self.properties.input_text(self.new_card)
        elif(card_title == "Message Box"):
            self.properties.message_box(self.new_card)
        elif(card_title == "Open Workbook"):
            self.properties.excel_open(self.new_card)
        elif(card_title == "Create Workbook"):
            self.properties.excel_create(self.new_card)
        elif(card_title == "Delete Workbook"):
            self.properties.excel_delete_file(self.new_card)
        elif(card_title == "Create Worksheet"):
            self.properties.excel_create_worksheet(self.new_card)
        elif(card_title == "Delete Worksheet"):
            self.properties.excel_delete_worksheet(self.new_card)
        elif(card_title == "Set Value"):
            self.properties.excel_set_value(self.new_card)
        elif(card_title == "Get Value"):
            self.properties.excel_get_value(self.new_card)
        elif(card_title == "Set Formula"):
            self.properties.excel_set_formula(self.new_card)
        elif(card_title == "Merge Unmerge"):
            self.properties.excel_merge_unmerge(self.new_card)
        elif(card_title == "Copy Paste"):
            self.properties.excel_copy_paste(self.new_card)
        elif(card_title == "Save Workbook"):
            self.properties.excel_save_workbook(self.new_card)


    def savePosn(self, event):
        self.lastx = event.x
        self.lasty = event.y

    def move(self, event):
        self.my_canvas.move(self.new_card, event.x-self.lastx, event.y-self.lasty)
        self.lastx = event.x
        self.lasty = event.y

    def onclick(self):
        self.my_canvas.itemconfig(self.new_card)
        if(self.cards_dict[self.new_card] == 'Open URL'):
            self.properties.open_url(self.new_card)
        elif(self.cards_dict[self.new_card] == 'Click'):
            self.properties.click(self.new_card)
        elif(self.cards_dict[self.new_card] == 'Clear'):
            self.properties.clear(self.new_card)
        elif(self.cards_dict[self.new_card] == 'Read Text'):
            self.properties.read_text(self.new_card)
        elif(self.cards_dict[self.new_card] == "Input Text"):
            self.properties.input_text(self.new_card)
        elif(self.cards_dict[self.new_card] == "Message Box"):
            self.properties.message_box(self.new_card)
        elif(self.cards_dict[self.new_card] == "Open Workbook"):
            self.properties.excel_open(self.new_card)
        elif(self.cards_dict[self.new_card] == "Create Workbook"):
            self.properties.excel_create(self.new_card)
        elif(self.cards_dict[self.new_card] == "Delete Workbook"):
            self.properties.excel_delete_file(self.new_card)
        elif(self.cards_dict[self.new_card] == "Create Worksheet"):
            self.properties.excel_create_worksheet(self.new_card)
        elif(self.cards_dict[self.new_card] == "Delete Worksheet"):
            self.properties.excel_delete_worksheet(self.new_card)
        elif(self.cards_dict[self.new_card] == "Set Value"):
            self.properties.excel_set_value(self.new_card)
        elif(self.cards_dict[self.new_card] == "Get Value"):
            self.properties.excel_get_value(self.new_card)
        elif(self.cards_dict[self.new_card] == "Set Formula"):
            self.properties.excel_set_formula(self.new_card)
        elif(self.cards_dict[self.new_card] == "Merge Unmerge"):
            self.properties.excel_merge_unmerge(self.new_card)
        elif(self.cards_dict[self.new_card] == "Copy Paste"):
            self.properties.excel_copy_paste(self.new_card)
        elif(self.cards_dict[self.new_card] == "Save Workbook"):
            self.properties.excel_save_workbook(self.new_card)


    def do_popup(self, event):
        try:
            self.select_x , self.select_y = event.x, event.y
            self.right_click_menu.tk_popup(event.x_root, event.y_root)
            item = self.my_canvas.find_closest(event.x, event.y)
            tags = self.my_canvas.itemcget(item, "tags")
            self.new_card = tags
        finally:
            self.right_click_menu.grab_release()

    def delete(self):
        self.my_canvas.delete(self.new_card)
        self.cards.remove(self.new_card)
        file_path = my_path + "files\\"
        if(os.path.isfile(file_path + self.new_card + '.xml')):
            os.remove(file_path + self.new_card + '.xml')
        if(len(self.cards) == 0):
            self.properties.reset_all("Empty")

    def get_cards(self):
        return self.cards

    def get_card_dictionary(self):
        return self.cards_dict


        

