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

        btn_create_workbook = tk.Button(self.excel_tab, text="Create Workbook", command=lambda : self.my_canvas.create_card('Create Workbook'))
        btn_create_workbook.pack(anchor = 'nw')

        btn_open_workbook = tk.Button(self.excel_tab, text="Open Workbook", command=lambda : self.my_canvas.create_card('Open Workbook'))
        btn_open_workbook.pack(anchor = 'nw')

        btn_delete_workbook = tk.Button(self.excel_tab, text="Delete Workbook", command=lambda : self.my_canvas.create_card('Delete Workbook'))
        btn_delete_workbook.pack(anchor = 'nw')

        btn_create_worksheet = tk.Button(self.excel_tab, text="Create Worksheet", command=lambda : self.my_canvas.create_card('Create Worksheet'))
        btn_create_worksheet.pack(anchor = 'nw')

        btn_delete_worksheet = tk.Button(self.excel_tab, text="Delete Worksheet", command=lambda : self.my_canvas.create_card('Delete Worksheet'))
        btn_delete_worksheet.pack(anchor = 'nw')

        btn_save_workbook = tk.Button(self.excel_tab, text="Save Workbook", command=lambda : self.my_canvas.create_card('Save Workbook'))
        btn_save_workbook.pack(anchor = 'nw')

        btn_set_cell_value = tk.Button(self.excel_tab, text="Set Value", command=lambda : self.my_canvas.create_card('Set Value'))
        btn_set_cell_value.pack(anchor = 'nw')

        btn_get_cell_value = tk.Button(self.excel_tab, text="Get Value", command=lambda : self.my_canvas.create_card('Get Value'))
        btn_get_cell_value.pack(anchor = 'nw')

        btn_set_cell_formula = tk.Button(self.excel_tab, text="Set Formula", command=lambda : self.my_canvas.create_card('Set Formula'))
        btn_set_cell_formula.pack(anchor = 'nw')

        btn_merge_unmerge = tk.Button(self.excel_tab, text="Merge/Unmerge", command=lambda : self.my_canvas.create_card('Merge Unmerge'))
        btn_merge_unmerge.pack(anchor = 'nw')

        btn_copy_paste = tk.Button(self.excel_tab, text="Copy Paste", command=lambda : self.my_canvas.create_card('Copy Paste'))
        btn_copy_paste.pack(anchor = 'nw')

        btn_select_cell = tk.Button(self.excel_tab, text="Select Cell/Row/Column", command=lambda : self.my_canvas.create_card('Select Cell/Row/Column'))
        btn_select_cell.pack(anchor = 'nw')

        btn_delete_cell = tk.Button(self.excel_tab, text="Delete Cell/Row/Column", command=lambda : self.my_canvas.create_card('Delete Cell/Row/Column'))
        btn_delete_cell.pack(anchor = 'nw')

        btn_read_cell = tk.Button(self.excel_tab, text="Read Cell/Row/Column", command=lambda : self.my_canvas.create_card('Read Cell/Row/Column'))
        btn_read_cell.pack(anchor = 'nw')

        btn_find_cell = tk.Button(self.excel_tab, text="Find", command=lambda : self.my_canvas.create_card('Find'))
        btn_find_cell.pack(anchor = 'nw')


        #Actions for Web automation
        btn_open_web = tk.Button(self.web_tab, text="Open URL", command= lambda : self.my_canvas.create_card('Open URL'))
        btn_open_web.pack(anchor = 'nw')

        btn_click_web = tk.Button(self.web_tab, text="Click", command=lambda : self.my_canvas.create_card('Click'))
        btn_click_web.pack(anchor = 'nw')

        btn_input_text_web = tk.Button(self.web_tab, text="Input Text", command=lambda : self.my_canvas.create_card('Input Text'))
        btn_input_text_web.pack(anchor = 'nw')

        btn_input_text_web = tk.Button(self.web_tab, text="Read Text", command=lambda : self.my_canvas.create_card('Read Text'))
        btn_input_text_web.pack(anchor = 'nw')

        btn_clear_web = tk.Button(self.web_tab, text="Clear", command=lambda : self.my_canvas.create_card('Clear'))
        btn_clear_web.pack(anchor = 'nw')

        btn_select_web = tk.Button(self.web_tab, text="Select", command=lambda : self.my_canvas.create_card('Select'))
        btn_select_web.pack(anchor = 'nw')

        btn_switch_to_web = tk.Button(self.web_tab, text="Switch To", command=lambda : self.my_canvas.create_card('Switch To'))
        btn_switch_to_web.pack(anchor = 'nw')

        btn_navigation_web = tk.Button(self.web_tab, text="Navigation", command=lambda : self.my_canvas.create_card('Navigation'))
        btn_navigation_web.pack(anchor = 'nw')

        btn_dragdrop_web = tk.Button(self.web_tab, text="Drag and Drop", command=lambda : self.my_canvas.create_card('Drag And Drop'))
        btn_dragdrop_web.pack(anchor = 'nw')

        #Actions for Control automation
        btn_loop_ctl = tk.Button(self.control_tab, text="Loop", command=lambda : self.my_canvas.create_card('Loop'))
        btn_loop_ctl.pack(anchor = 'nw')

        btn_break_ctl = tk.Button(self.control_tab, text="Break", command=lambda : self.my_canvas.create_card('Break'))
        btn_break_ctl.pack(anchor = 'nw')

        btn_continue_ctl = tk.Button(self.control_tab, text="Continue", command=lambda : self.my_canvas.create_card('Continue'))
        btn_continue_ctl.pack(anchor = 'nw')

        btn_wait_ctl = tk.Button(self.control_tab, text="Wait", command=lambda : self.my_canvas.create_card('Wait'))
        btn_wait_ctl.pack(anchor = 'nw')

        btn_delay_ctl = tk.Button(self.control_tab, text="Delay", command=lambda : self.my_canvas.create_card('Delay'))
        btn_delay_ctl.pack(anchor = 'nw')

        #Actions for More automation
        btn_keystroke_more = tk.Button(self.more_tab, text="Key Stroke", command=lambda : self.my_canvas.create_card('Key Stroke'))
        btn_keystroke_more.pack(anchor = 'nw')

        btn_messagebox_more = tk.Button(self.more_tab, text="Message Box", command=lambda : self.my_canvas.create_card('Message Box'))
        btn_messagebox_more.pack(anchor = 'nw')

        btn_leftclick_more = tk.Button(self.more_tab, text="Left Click", command=lambda : self.my_canvas.create_card('Left Click'))
        btn_leftclick_more.pack(anchor = 'nw')

        btn_rightclick_more = tk.Button(self.more_tab, text="Right Click", command=lambda : self.my_canvas.create_card('Right Click'))
        btn_rightclick_more.pack(anchor = 'nw')

        btn_scroll_more = tk.Button(self.more_tab, text="Scroll", command=lambda : self.my_canvas.create_card('Scroll'))
        btn_scroll_more.pack(anchor = 'nw')
    
    
