import tkinter as tk
from tkinter import ttk

class Actions(ttk.Frame):
    def __init__(self, parent, canvas, properties, *args, **kwargs):
        ttk.Frame.__init__(self, parent , *args, **kwargs)
        self.parent_frame = parent
        self.my_canvas = canvas
        self.properties = properties

        self.tabControl = ttk.Notebook(self.parent_frame, height = 360)

        self.web_tab = ttk.Frame(self.tabControl)
        self.window_tab = ttk.Frame(self.tabControl)
        self.excel_tab = ttk.Frame(self.tabControl)
        self.control_tab = ttk.Frame(self.tabControl)
        self.more_tab = ttk.Frame(self.tabControl)

        self.tabControl.add(self.web_tab, text="Web")
        self.tabControl.add(self.window_tab, text="Window")
        self.tabControl.add(self.excel_tab, text="Excel")
        self.tabControl.add(self.control_tab, text="Control")
        self.tabControl.add(self.more_tab, text="More")
        self.tabControl.pack(expand = 1, fill = tk.BOTH)

        #Actions for excel automation
        btn_create_excel = tk.Button(self.excel_tab, text="Create", command=self.my_canvas.create_card)
        btn_create_excel.pack(anchor = 'nw')

        btn_create_workbook = tk.Button(self.excel_tab, text="Create Workbook", command=self.my_canvas.create_card)
        btn_create_workbook.pack(anchor = 'nw')

        btn_delete_workbook = tk.Button(self.excel_tab, text="Delete Workbook", command=self.my_canvas.create_card)
        btn_delete_workbook.pack(anchor = 'nw')

        btn_create_worksheet = tk.Button(self.excel_tab, text="Create Worksheet", command=self.my_canvas.create_card)
        btn_create_worksheet.pack(anchor = 'nw')

        btn_delete_worksheet = tk.Button(self.excel_tab, text="Delete Worksheet", command=self.my_canvas.create_card)
        btn_delete_worksheet.pack(anchor = 'nw')

        btn_save_workbook = tk.Button(self.excel_tab, text="Save Workbook", command=self.my_canvas.create_card)
        btn_save_workbook.pack(anchor = 'nw')

        btn_select_cell = tk.Button(self.excel_tab, text="Select Cell/Row/Column", command=self.my_canvas.create_card)
        btn_select_cell.pack(anchor = 'nw')

        btn_delete_cell = tk.Button(self.excel_tab, text="Delete Cell/Row/Column", command=self.my_canvas.create_card)
        btn_delete_cell.pack(anchor = 'nw')

        btn_read_cell = tk.Button(self.excel_tab, text="Read Cell/Row/Column", command=self.my_canvas.create_card)
        btn_read_cell.pack(anchor = 'nw')

        btn_find_cell = tk.Button(self.excel_tab, text="Find", command=self.my_canvas.create_card)
        btn_find_cell.pack(anchor = 'nw')


        #Actions for Web automation
        btn_open_web = tk.Button(self.web_tab, text="Open", command=self.my_canvas.create_card)
        btn_open_web.pack(anchor = 'nw')

        btn_click_web = tk.Button(self.web_tab, text="Click", command=self.my_canvas.create_card)
        btn_click_web.pack(anchor = 'nw')

        btn_input_text_web = tk.Button(self.web_tab, text="Input Text", command=self.my_canvas.create_card)
        btn_input_text_web.pack(anchor = 'nw')

        btn_dragdrop_web = tk.Button(self.web_tab, text="Drag and Drop", command=self.my_canvas.create_card)
        btn_dragdrop_web.pack(anchor = 'nw')

        #Actions for Control automation
        btn_loop_ctl = tk.Button(self.control_tab, text="Loop", command=self.my_canvas.create_card)
        btn_loop_ctl.pack(anchor = 'nw')

        btn_break_ctl = tk.Button(self.control_tab, text="Break", command=self.my_canvas.create_card)
        btn_break_ctl.pack(anchor = 'nw')

        btn_continue_ctl = tk.Button(self.control_tab, text="Continue", command=self.my_canvas.create_card)
        btn_continue_ctl.pack(anchor = 'nw')

        btn_wait_ctl = tk.Button(self.control_tab, text="Wait", command=self.my_canvas.create_card)
        btn_wait_ctl.pack(anchor = 'nw')

        btn_delay_ctl = tk.Button(self.control_tab, text="Delay", command=self.my_canvas.create_card)
        btn_delay_ctl.pack(anchor = 'nw')

        #Actions for More automation
        btn_keystroke_more = tk.Button(self.more_tab, text="Key Stroke", command=self.my_canvas.create_card)
        btn_keystroke_more.pack(anchor = 'nw')

        btn_messagebox_more = tk.Button(self.more_tab, text="Message Box", command=self.my_canvas.create_card)
        btn_messagebox_more.pack(anchor = 'nw')

        btn_leftclick_more = tk.Button(self.more_tab, text="Left Click", command=self.my_canvas.create_card)
        btn_leftclick_more.pack(anchor = 'nw')

        btn_rightclick_more = tk.Button(self.more_tab, text="Right Click", command=self.my_canvas.create_card)
        btn_rightclick_more.pack(anchor = 'nw')

        btn_scroll_more = tk.Button(self.more_tab, text="Scroll", command=self.my_canvas.create_card)
        btn_scroll_more.pack(anchor = 'nw')
    
    
