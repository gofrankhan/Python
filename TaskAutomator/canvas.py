import tkinter as tk
from tkinter import ttk

class Canvas():

    def __init__(self, *args, **kwargs):
        self.parent_frame = None
        self.canvas_frame = None
        self.my_canvas = None
        self.xscroll = None
        self.yscroll = None
        self.cards = []
        self.card_number = 0
        self.new_card = None
        self.select_x = 0
        self.select_y = 0
        self.normal_font = ("Normal")

    def create_Ui(self, parent):
        self.parent_frame = parent
        self.canvas_frame = ttk.Frame(self.parent_frame, height = 950, width = 640)
        self.canvas_frame.pack(expand = True, fill = tk.BOTH, side = tk.TOP)
        self.my_canvas = tk.Canvas(self.canvas_frame, width = 1000, height = 600, background = 'white', scrollregion = (0,0,2200, 4380))
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

        #self.image_button = tk.Button(self, text = 'Hello')
        #self.my_canvas.create_window(10,10, window = self.image_button)

    def create_card(self):
        self.new_card = 'card' + str(self.card_number)
        self.cards.append(self.new_card)
        self.card_number += 1
        self.my_canvas.create_rectangle(50, 20, 200, 100, fill='#dfcdaa', tags=[self.new_card])
        self.my_canvas.create_line(50,45, 200, 45, tags=[self.new_card])
        self.my_canvas.create_text(55,40, text="Title of the card", anchor = 'sw', tags=[self.new_card])
        self.my_canvas.create_text(55,50, text="Details of the card \nIt may have \nmultiple lines", anchor = 'nw', tags=[self.new_card])
        #self.my_canvas.move(a, 20, 20)
        self.my_canvas.tag_bind(self.new_card, '<B1-Motion>', self.move)
        self.my_canvas.tag_bind(self.new_card, '<ButtonPress-1>', self.savePosn)

    def savePosn(self, event):
        self.lastx = event.x
        self.lasty = event.y

    def move(self, event):
        self.my_canvas.move(self.new_card, event.x-self.lastx, event.y-self.lasty)
        self.lastx = event.x
        self.lasty = event.y

    def onclick(self):
        self.my_canvas.itemconfig(self.new_card)

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

        

