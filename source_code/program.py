from parameters import *

##### MAIN PROGRAM MODULE
class Program(object):
    def __init__(self):
        self.root = tk.ThemedTk()
        self.root.get_themes()
        self.root.set_theme('clam')
        self.path = os.getcwd()
        self.icon = ICON
        
        self.root.title(PROGRAM + VERSION)
        
        self.main_frame_format()
        self.menu_format()
        self.frame_load_format()
        self.frame_fill_format()

        self.data_length = 0
        self.initial_length = 0
        self.start_err = 0
        self.csv_appended = False
        self.csv_other_list = []

        self.current_csv = ''       
        

    ## FORMATTING
    def main_frame_format(self):
        windowWidth = self.root.winfo_reqwidth()
        windowHeight = self.root.winfo_reqheight()

        self.root.geometry("+{}+{}".format(SCREEN_POS_RIGHT, SCREEN_POS_DOWN))
        self.root.iconbitmap(self.icon)

        self.main_canvas = t.Canvas(self.root, bg = BLACK, width = WIDTH, height = HEIGHT)
        self.main_canvas.pack(side = 'left', fill = 'both')

        self.menu = t.Menu(self.root)
        self.root.config(menu = self.menu)
        
        
        
    def menu_format(self):        
        file_menu = t.Menu(self.menu, tearoff = False)
        self.menu.add_cascade(label = 'File', menu = file_menu)
        
        csv_menu = t.Menu(self.menu, tearoff = False)
        self.menu.add_cascade(label = 'CSV', menu = csv_menu)
        
        report_menu = t.Menu(self.menu, tearoff = False)
        self.menu.add_cascade(label = 'Process Data', menu = report_menu)
        
        file_menu.add_command(label = "New Sheet", command = lambda: self.manual_cell_generator())
        file_menu.add_command(label = "Browse for CSV...", command = lambda: self.browse_for_csv())
        file_menu.add_separator()

        csv_menu.add_command(label = "Save New CSV...", command = lambda: self.csv_save_new())
        csv_menu.add_command(label = "Append Current CSV", command = lambda: self.csv_save_append_current())
        csv_menu.add_command(label = "Append to Other CSV...", command = lambda: self.csv_save_append_other())
        csv_menu.add_separator()

        report_menu.add_command(label = "Go To Processing", command = lambda: self.view_report())
        report_menu.add_separator()


    def frame_load_format(self):
        self.frame_load = t.Frame(self.main_canvas, bg = FORESTGREEN, bd = 5, highlightbackground = BLACK, highlightcolor = BLACK, highlightthickness = 5)
        self.frame_load.place(relwidth = 0.2, relheight = 1.0)

        load_label = t.Label(self.frame_load, font = font18Cb, bg = FORESTGREEN, text = "Data Sheet Tips")
        load_label.grid(row = 0, column = 0, columnspan = 4, sticky = "n")
        
        mrow = 1
        for i in METRIC_LIST:
            if i == 'SPECIES':
                tip_label1 = t.Label(self.frame_load, font = font12Cb, bg = FORESTGREEN, text = str(i))
                tip_label1.grid(row = mrow, column = 0, columnspan = 4, sticky = "w")
                mrow += 1
                
                menu_list = []
                for key in Timber.ALL_SPECIES_NAMES:
                    if key != 'TOTALS':
                        txt = key+":    "+Timber.ALL_SPECIES_NAMES[key]
                        menu_list.append(txt)
                        
                header = t.StringVar()
                header.set("SPECIES CODES")
                spp_menu = t.OptionMenu(self.frame_load, header , *menu_list)
                spp_menu.grid(row = mrow, column = 0, columnspan = 4, sticky = "w")
                
                mrow += 1

                blank = t.Label(self.frame_load, bg = FORESTGREEN, font = font8Cb, text = "")
                blank.grid(row = mrow, column = 0)  

                mrow += 1
                
            else:               
                tip_label = t.Label(self.frame_load, font = font12Cb, bg = FORESTGREEN, text = str(i))
                tip_label.grid(row = mrow, column = 0, columnspan = 4, sticky = "w")
                mrow += 1
                for j in METRIC_DICT[i]:
                    tip_label = t.Label(self.frame_load, font = font10Cb, bg = FORESTGREEN, text = str(j))
                    tip_label.grid(row = mrow, column = 0, columnspan = 4, sticky = "w")
                    mrow += 1
                    
                blank = t.Label(self.frame_load, bg = FORESTGREEN, font = font8Cb, text = "")
                blank.grid(row = mrow, column = 0)   

                mrow += 1
                
        

    def frame_fill_format(self):
        self.frame_fill = t.Frame(self.main_canvas, bg = FORESTGREEN, bd = 5, highlightbackground = BLACK, highlightcolor = BLACK, highlightthickness = 5)
        self.frame_fill.place(x = int(WIDTH * 0.2), relwidth = 0.8, relheight = 1.0)

        self.fill_frame_labels = t.Frame(self.frame_fill, bg = FORESTGREEN)
        self.fill_frame_labels.place(relwidth = 1.0, relheight = 0.04)

        for i in range(len(METRIC_LIST)):
            self.metric_label = t.Label(self.fill_frame_labels, font = font14Cb, bg = FORESTGREEN, text = METRIC_LIST[i], width = 14)
            self.metric_label.grid(row = 0, column = i+1, sticky = "n")           


        self.fill_frame_data = t.Frame(self.frame_fill, bg = FORESTGREEN)
        self.fill_frame_data.place(y = int(0.04 * HEIGHT), relwidth = 1.0, relheight = 0.96)        

        self.fill_canvas = t.Canvas(self.fill_frame_data, bg = FORESTGREEN, highlightthickness = 0)
        self.fill_canvas.pack(side = 'left', fill = 'both', expand = 1)
        
        self.fill_canvas.bind('<Enter>', self.canvas_bound)
        self.fill_canvas.bind('<Leave>', self.canvas_unbound)

        self.fill_scrollbar = ttk.Scrollbar(self.fill_frame_data, orient = 'vertical', command = self.fill_canvas.yview)
        self.fill_scrollbar.pack(side='right', fill='y')
        
        self.fill_canvas.config(yscrollcommand = self.fill_scrollbar.set)
        self.fill_canvas.bind('<Configure>', lambda e: self.fill_canvas.configure(scrollregion = self.fill_canvas.bbox("all")))
      
        self.fill_frame = t.Frame(self.fill_canvas, bg = FORESTGREEN)
        self.fill_frame.pack(fill = 'both', expand = 1)
        
        self.fill_canvas.create_window((0,0), window=self.fill_frame, anchor="nw")
        


    def canvas_bound(self, event):
        self.fill_canvas.bind_all('<MouseWheel>', self.on_mousewheel)
        self.fill_canvas.bind_all('<Return>', self.enter_down)
        self.fill_canvas.bind_all('<Down>', self.enter_down)
        self.fill_canvas.bind_all('<Left>', self.left)
        self.fill_canvas.bind_all('<Right>', self.right)
        self.fill_canvas.bind_all('<Up>', self.up)
        

    def canvas_unbound(self, event):
        self.fill_canvas.unbind_all('<MouseWheel>')
        self.fill_canvas.unbind_all('<Return>')
        self.fill_canvas.unbind_all('<Down>')
        self.fill_canvas.unbind_all('<Left>')
        self.fill_canvas.unbind_all('<Right>')
        self.fill_canvas.unbind_all('<Up>')
        

    def on_mousewheel(self, event):
        self.fill_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        

    def blank_label(self, Parent_Frame, Row, Column):
        self.blank = t.Label(Parent_Frame, bg = FORESTGREEN, font = font12Cb, text = "")
        self.blank.grid(row = Row, column = Column)

        

    ## USER FUNCTIONS
    def browse_for_csv(self):        
        csv_loaded = filedialog.askopenfilename(initialdir =  self.path, title = "Select a File",
                                                  filetypes = (("CSV Files", "*.csv*"), ("All Files", "*.*")))        
        if csv_loaded == '':
            return
        else:
            self.current_csv = csv_loaded
            self.frame_fill.destroy()
            self.frame_fill_format()
            with open(csv_loaded, 'r') as metric_data:    
                metric_data_reader = csv.reader(metric_data)    
                next(metric_data_reader)

                row = 0
                column = 1
                col_count = 1
            
                for line in metric_data_reader:
                    if line[0] == "":
                        break
                    else:
                        row_num = t.Label(self.fill_frame, font = font8Cb,bg = FORESTGREEN, text = str(row + 1), width = 3)
                        row_num.grid(row = row, column = 0, sticky = "w")
                        for i in line:
                            if col_count > len(METRIC_LIST):
                                break
                            else:
                                text = str(i)
                                fill_data = t.Entry(self.fill_frame, font = font14Cb, width = 14)                            
                                fill_data.grid(row = row, column = column, sticky = "w")
                                fill_data.insert(0, text)
                                col_count += 1
                                column += 1
                    row += 1                    
                    column = 1
                    col_count = 1
            
                self.data_length = row - 1
                self.initial_length = row - 1
                
            self.add_rows_button = t.Button(self.fill_frame, text = "Add Row", font = font12Cb, bd = 5, command = lambda: self.add_rows())
            self.add_rows_button.grid(row =self.data_length + 1, column = 1, sticky = "w")


    def manual_cell_generator(self):
        self.current_csv = ''
        self.data_length = 99
        self.initial_length = 99
        
        self.frame_fill.destroy()
        self.frame_fill_format()
        
        for i in range(self.data_length + 1):
            row_num = t.Label(self.fill_frame, font = font8Cb,bg = FORESTGREEN, text = str(i + 1), width = 3)
            row_num.grid(row = i, column = 0, sticky = "w")
            
            for j in range(1, len(METRIC_LIST) + 1):
                metric_entry = t.Entry(self.fill_frame, font = font14Cb, width = 14)
                metric_entry.grid(row = i, column = j, sticky = "w")

        self.add_rows_button = t.Button(self.fill_frame, text = "Add Row", font = font12Cb, bd = 5, command = lambda: self.add_rows())
        self.add_rows_button.grid(row =self.data_length + 1, column = 1, sticky = "w")
        


    def add_rows(self):
        self.add_rows_button.destroy()
        self.data_length += 1
        
        row_num = t.Label(self.fill_frame, font = font8Cb,bg = FORESTGREEN, text = str(self.data_length + 1), width = 3)
        row_num.grid(row = self.data_length, column = 0, sticky = "w")
        
        for j in range(1, len(METRIC_LIST) + 1):
                metric_entry = t.Entry(self.fill_frame, font = font14Cb, width = 14)
                metric_entry.grid(row = self.data_length, column = j, sticky = "w")

        
        self.add_rows_button = t.Button(self.fill_frame, text = "Add Row", font = font12Cb, bd = 5, command = lambda: self.add_rows())
        self.add_rows_button.grid(row =self.data_length + 1, column = 1, sticky = "w")

        self.root.update()
        self.fill_canvas.config(scrollregion=self.fill_canvas.bbox("all"))
        


    def extension_check(self, Filename, Ext_to_Check):
        check_ext = ''
        ci = -(len(Ext_to_Check))
        for i in range(len(Ext_to_Check)):
            check_ext += Filename[ci]
            ci += 1
        if check_ext != Ext_to_Check:
            Filename += Ext_to_Check
        return Filename


    def get_filename_only(self, Filename_Full):
        char_list = []
        for i in reversed(Filename_Full):
            if i == '/':
                break
            char_list.insert(0, i)
        filename = ''
        for i in char_list:
            filename += i
        return filename     



    def csv_save_new(self):
        check = self.error_check(list(self.fill_frame.children.values()))
        if check is None:
            return
        else:
            csv_saved = filedialog.asksaveasfilename(initialdir =  self.path, title = "Save CSV File",
                                                       filetypes = (("CSV Files", "*.csv"), ("All Files", "*.*")))
            if csv_saved == '':
                return
            else:
                csv_saved = self.extension_check(csv_saved, '.csv')
                self.current_csv = csv_saved
                with open(csv_saved, 'w', newline = '') as csv_file:
                    row_write = csv.writer(csv_file)
                    row_write.writerow(METRIC_LIST)
                    for i in check:
                        row_write.writerow(i)                       
    
                        
                        

    def csv_save_append_current(self):
        check = self.error_check(list(self.fill_frame.children.values()))
        if check is None:
            return
        else:
            if self.current_csv == '':
                messagebox.showwarning("Warning","No Current CSV\nSave Sheet first before Appending")
                return
            else:
                csv_saved = self.current_csv
                with open(csv_saved, 'w', newline = '') as csv_file:
                    row_write = csv.writer(csv_file)
                    row_write.writerow(METRIC_LIST)
                    for i in check:
                        row_write.writerow(i)

                messagebox.showinfo('Completed', 'Completed appending to current CSV:\n' + self.get_filename_only(csv_saved))
                        


    def csv_save_append_other(self):
        check = self.error_check(list(self.fill_frame.children.values()))
        if check is None:
            return
        else:
            csv_loaded = filedialog.askopenfilename(initialdir = self.path, title = "Select a File",
                                                      filetypes = (("CSV Files", "*.csv*"), ("All Files", "*.*")))            
            if csv_loaded == '':
                return
            else:
                temp_list = []
                with open(csv_loaded, 'r') as metric_data:    
                    metric_data_reader = csv.reader(metric_data)    
                    next(metric_data_reader)
                
                    for line in metric_data_reader:
                        if line[0] == "":
                            break
                        else:
                            temp_temp = []
                            for i in line:
                                temp_temp.append(i)
                            temp_list.append(temp_temp)

                data = temp_list + check            
                with open(csv_loaded, 'w', newline = '') as csv_file:
                    row_write = csv.writer(csv_file)
                    row_write.writerow(METRIC_LIST)
                    for i in data:
                        row_write.writerow(i)

                messagebox.showinfo('Completed', 'Completed appending to other CSV:\n' + self.get_filename_only(csv_loaded))

                

    
    def view_report(self):
        check = self.error_check(list(self.fill_frame.children.values()))
        if check is None:
            return
        else:
            if self.current_csv == '':
                messagebox.showwarning("Warning","No Current CSV\nSave Sheet first before Viewing Report")
            else:
                csv_saved = self.current_csv
                with open(csv_saved, 'w', newline = '') as csv_file:
                    row_write = csv.writer(csv_file)
                    row_write.writerow(METRIC_LIST)
                    for i in check:
                        row_write.writerow(i)
                View(self, self.current_csv)



    ## ERROR CHECKING
    def entry_error(self):
        current = self.root.focus_get()
        if current['fg'] == RED:
            current['fg'] = BLACK
        current['textvariable'] = None
        
                        

    def error_check(self, Data):
        data_sheet = Data
        
        err_list = []
        stand_height = {}

        # At least one height needed - below is error checking
        for i in range(1, len(data_sheet), (len(METRIC_LIST) + 1)):
            val = data_sheet[i].get()

            if val != '':
                if val in stand_height:
                    if data_sheet[i + 5].get() != '':
                        stand_height[val][0] += 1
                else:
                    if data_sheet[i + 5].get() != '':
                        stand_height[val] = [1, i + 5]
                    else:
                        stand_height[val] = [0, i + 5]
                        
        for key in stand_height:
            if stand_height[key][0] == 0:
                self.error_append(data_sheet[stand_height[key][1]], err_list, "Height Needed")

        if len(err_list) > 0:
            messagebox.showwarning("Warning","At least one Total Height is needed per stand")
            for i in err_list:
                i[0].config(textvariable = i[1])
                i[0].config(fg = RED)                
            return None

        # Error checking for value errors or missing values
        else:
            err_list = []
            data_list = []

            for i in range(1, len(data_sheet), (len(METRIC_LIST) + 1)):
                temp_temp = []
                
                val = data_sheet[i]

                if val.get() == '':
                    self.error_append(val, err_list, 'Missing')
                    filled_count = 0
                    
                    for j in range(i + 1, i + len(METRIC_LIST)):
                        tf, text = self.type_check(data_sheet[j], j % (len(METRIC_LIST) + 1))
                        if not tf:
                            if text == 'Missing':
                                self.error_append(data_sheet[j], err_list, text)
                            else:
                                filled_count += 1
                                self.error_append(data_sheet[j], err_list, text)
                                
                    if filled_count == 0:
                        for i in range(1, len(METRIC_LIST)):
                            err_list.pop(-1)
                        break
                    
                else:
                    temp_temp.append(val.get())
                    
                    for j in range(i + 1, i + len(METRIC_LIST)):
                        tf, text = self.type_check(data_sheet[j], j % (len(METRIC_LIST) + 1))
                        if not tf:
                            self.error_append(data_sheet[j], err_list, text)
                        else:
                            temp_temp.append(data_sheet[j].get())

                    data_list.append(temp_temp)
                    
            if len(err_list) > 0:
                messagebox.showwarning("Warning","One or more required values\nare missing or incorrect")
                for i in err_list:
                    i[0].config(textvariable = i[1])
                    i[0].config(fg = RED)                
                return None

            else:
                return data_list


    # Data type check
    def type_check(self, Entry, Column_Count):
        required_cols = [1, 2, 3, 4, 5, 7]
        int_cols = [2, 3]
        float_cols = [5, 6, 7]
        
        val = Entry.get()

        if Column_Count in required_cols:
            if val == '':
                return False, 'Missing'
            else:
                if Column_Count in int_cols:
                    try:
                        x = int(val)
                        x = math.sqrt(int(val))
                        return True, ''
                    except ValueError:
                        return False, 'Val Err'
                elif Column_Count in float_cols:
                    try:
                        x = float(val)
                        if Column_Count == 7:
                            return True, ''
                        else:
                            x = math.sqrt(float(val))
                            return True, ''
                    except ValueError:
                        return False, 'Val Err'

                elif Column_Count == 4:
                    if val.upper() in Timber.ALL_SPECIES_NAMES:
                        return True, ''
                    else:
                        return False, 'Spp Err'

        else:
            if val != '':
                try:
                    x = float(val)
                    x = math.sqrt(float(val))
                    return True, ''
                except ValueError:
                    return False, 'Val Err'
            else:
                return True, ''

        


    def error_append(self, Entry, Error_List, Error_Text):
        prev_text = Entry.get()
        if prev_text == '':
            text = Error_Text
        else:
            text = "'" + prev_text + "'" + " " + Error_Text
            
        sv = t.StringVar(master = self.root, value = text)
        sv.trace("w", lambda name, index, mode, sv=sv: self.entry_error())
        
        Error_List.append((Entry, sv))            
  
    
        
    ## BINDING EVENTS
    def enter_down(self, _event = None):           
        current = self.root.focus_get().grid_info()
        if current['row'] == self.data_length:
            return
        else:
            row = current['row'] + 1
            column = current['column']

        for children in self.fill_frame.children.values():
            if children.grid_info()['row'] == row and children.grid_info()['column'] == column:
                children.focus_set()
                if children.winfo_rooty() >= HEIGHT:
                    self.fill_canvas.yview_scroll(1, "units")
                return



    def left(self, _event = None):         
        current = self.root.focus_get().grid_info()
        if current['row'] == 0 and current['column'] == 1:
            return
        elif current['column'] == 1:
            row = current['row'] - 1
            column = len(METRIC_LIST)
        else:
            row = current['row']
            column = current['column'] - 1

        for children in self.fill_frame.children.values():
            if children.grid_info()['row'] == row and children.grid_info()['column'] == column:
                children.focus_set()
                if children.winfo_rooty() <= 120:
                    self.fill_canvas.yview_scroll(-1, "units")
                return


    def right(self, _event = None):           
        current = self.root.focus_get().grid_info()
        if current['row'] == self.data_length and current['column'] == len(METRIC_LIST):
            return
        elif current['column'] == len(METRIC_LIST):
            row = current['row'] + 1
            column = 1
        else:
            row = current['row']
            column = current['column'] + 1

        for children in self.fill_frame.children.values():
            if children.grid_info()['row'] == row and children.grid_info()['column'] == column:
                children.focus_set()
                if children.winfo_rooty() >= HEIGHT:
                    self.fill_canvas.yview_scroll(1, "units")
                return



    def up(self, _event = None):           
        current = self.root.focus_get().grid_info()
        if current['row'] == 0:
            return
        else:
            row = current['row'] - 1
            column = current['column']

        for children in self.fill_frame.children.values():
            if children.grid_info()['row'] == row and children.grid_info()['column'] == column:
                children.focus_set()
                if children.winfo_rooty() <= 120:
                    self.fill_canvas.yview_scroll(-1, "units")
                return

















##### VIEW REPORT FOR MAIN PROGRAM
class View(object):

    def __init__(self, Program, CSV):
        self.program = Program
        
        self.csv = CSV
        self.stands = []
        self.get_stands_and_plot_num()

        self.stands_access_dict = {}
        for i in self.stands:
            self.stands_access_dict[i[0]] = {}
        self.current_stand = -1            

        self.stand = ''
        self.plots = 0
        self.show_all_stands = False
        
        self.first_time_running = True
        self.stand_generated = False
        
        self.plog = 0
        self.mlog = 0

        self.plog_text = "Preferred Log Length: "
        self.mlog_text = "Minimum Log Length: "

        self.choose_access = False
        self.choose_sqlite3 = False
        self.choose_excel = False
        self.choose_multiple = False

        self.top = t.Toplevel()
        self.top.transient(self.program.root)
        
        self.width = 1250
        self.height = 750
        
        self.top.geometry("+{}+{}".format(SCREEN_POS_RIGHT * 3, SCREEN_POS_DOWN * 3))
        self.top.iconbitmap(self.program.icon)

        self.top.bind('<FocusOut>', lambda e: self.top.lower(self.program.root))

        self.menu = t.Menu(self.top)
        self.top.config(menu = self.menu)
        self.menu_format()

        self.top_mainframe = t.Frame(self.top, bg = BLACK, bd = 5, highlightcolor = BLACK, highlightbackground = BLACK, width = self.width, height = self.height)
        self.top_mainframe.pack(fill = 'both', expand = 1)

        self.top_stands_frame = t.Frame(self.top_mainframe, bg = FORESTGREEN, bd = 5, highlightcolor = BLACK, highlightbackground = BLACK)
        self.top_stands_frame.place(relwidth = 1.0, relheight = 0.25)

        self.top_report_frame = t.Frame(self.top_mainframe, bg = FORESTGREEN, bd = 5, highlightcolor = BLACK, highlightbackground = BLACK)
        self.top_report_frame.place(y = int(self.height * 0.25), relwidth = 1.0, relheight = 0.75) 

        self.stands_frame_format()

        self.top_canvas = t.Canvas(self.top_report_frame, bg = FORESTGREEN)
        self.top_canvas.pack(side = 'left', fill = 'both', expand = 1)
        
        self.top_canvas.bind("<Enter>", self.canvas_bound)
        self.top_canvas.bind("<Leave>", self.canvas_unbound)

        self.top_scrollbar = ttk.Scrollbar(self.top_report_frame, orient = 'vertical', command = self.top_canvas.yview)
        self.top_scrollbar.pack(side='right', fill='y')
        self.top_canvas.config(yscrollcommand = self.top_scrollbar.set)

        self.report_frame_format()

        self.top.mainloop()


    ## FORMATTING
    def menu_format(self):
        db_menu = t.Menu(self.menu, tearoff = False)
        self.menu.add_cascade(label = 'Export to FVS', menu = db_menu)
        
        db_menu.add_command(label = "Export to Access Database", command = lambda: self.export_db_start('Access'))
        db_menu.add_command(label="Export to SQLite3 Database", command=lambda: self.export_db_start('SQLite3'))
        db_menu.add_command(label="Export to Excel Database", command=lambda: self.export_db_start('Excel'))
        db_menu.add_command(label="Export to Multiple Databases", command=lambda: self.export_db_start('Multiple'))
        db_menu.add_separator()


    def canvas_bound(self, event):
        self.top_canvas.bind_all('<MouseWheel>', self.on_mousewheel)

    def canvas_unbound(self, event):
        self.top_canvas.unbind_all('<MouseWheel>')

    def on_mousewheel(self, event):
        self.top_canvas.yview_scroll(int(-1*(event.delta/120)), "units")


    def stands_frame_format(self):
        try:
            self.top_stands_hold_frame.destroy()
        except:
            pass
        self.current_stand = -1
        self.top_stands_hold_frame = t.Frame(self.top_stands_frame, bg = FORESTGREEN)
        self.top_stands_hold_frame.pack(fill = 'both', expand = 1) 
        
        blank = t.Label(self.top_stands_hold_frame, bg = FORESTGREEN, font = font8Cb, bd = 5, text = "")
        blank.grid(row = 0, column = 0, sticky = 'w')

        self.stands_row = 1
        label = t.Label(self.top_stands_hold_frame, bg = FORESTGREEN, font = font14Cb, bd = 5, text = "STANDS: ")
        label.grid(row = self.stands_row, column = 0, sticky = 'w')
        
        col = 1

        if len(self.stands) > 2:
            all_stand_button = t.Button(self.top_stands_hold_frame, font = font12Cb, width = 10, text = "ALL STANDS", bd = 5, command = lambda: self.choose_stand(all_stands = True))
            all_stand_button.grid(row = self.stands_row, column = col, sticky = "w")        
            col += 1
        
        for i in self.stands:
            if col > 11:
                self.stands_row += 1
                if len(self.stands) > 2:
                    col = 2
                else:
                    col = 1
            stand_button = t.Button(self.top_stands_hold_frame, font = font12Cb, width = 10, text = i[0], bd = 5, command = lambda i=i: self.choose_stand(stand = i[0], plots = i[1]))
            stand_button.grid(row = self.stands_row, column = col, sticky = "w")
            col += 1
        self.stands_row += 1

        self.filler_plog = t.Label(self.top_stands_hold_frame, bg = FORESTGREEN, font = font12Cb, bd = 5, width = (len(self.plog_text) - 3), text = "")
        self.filler_plog.grid(row = self.stands_row, column = 0, sticky = 'w')

        if self.stands_row > 3:
            rel = 0.25 + ((self.stands_row - 3) * 0.06)

            self.top_stands_frame.place(relwidth = 1.0, relheight = rel)
            self.top_report_frame.place(y = int(self.height * rel), relwidth = 1.0, relheight = 1 - rel)        


    def report_frame_format(self):       
        
        self.top_frame = t.Frame(self.top_canvas, bg = FORESTGREEN)
        self.top_frame.pack(fill = 'both', expand = 1) 

        self.top_canvas.create_window((0,0), window=self.top_frame, anchor="nw")
        self.top_canvas.bind('<Configure>', lambda e: self.top_canvas.configure(scrollregion = self.top_canvas.bbox("all")))


    def get_stands_and_plot_num(self):
        master = []
        temp_dict = {}
        with open(self.csv, 'r') as tree_data:
            tree_data_reader = csv.reader(tree_data)
            next(tree_data_reader)
            for line in tree_data_reader:
                if line[0] == "":
                    break
                else:
                    if str(line[0]) in temp_dict:
                        temp_dict[str(line[0]).upper()].append(int(line[1]))
                    else:
                        temp_dict[str(line[0]).upper()] = [int(line[1])]
        
        for key in temp_dict:
            self.stands.append([key, max(temp_dict[key])])
                    
        
        

    ## GENERATING REPORTS (ONSCREEN AND PDF)
    def choose_stand(self, stand = '', plots = 0, all_stands = False):
        if all_stands:
            self.show_all_stands = True
        else:
            self.show_all_stands = False
        self.stand = stand
        self.plots = plots
        self.choose_logs()
        
        
    def choose_logs(self):
        if not self.first_time_running:
            self.top_frame.destroy()
            self.top_stands_hold_frame.destroy()
            self.stands_frame_format()
            self.report_frame_format()
            self.top_frame.update()
            self.top_canvas['scrollregion'] = self.top_frame.bbox("all")

        self.stand_generated = False

        blank = t.Label(self.top_frame, bg = FORESTGREEN, font = font10Cb, bd = 5, text = "")
        blank.grid(row = 0, column = 0, sticky = 'w')        
        
        stand_for = t.Label(self.top_frame, bg = FORESTGREEN, font = font14Cb, bd = 5, text = 'STAND REPORT FOR:')
        stand_for.grid(row = 1, column = 1, sticky = 'w')

        if self.show_all_stands:
            stand_text = t.Label(self.top_frame, bg = FORESTGREEN, font = font14Cb, bd = 5, text = 'ALL STANDS')
            stand_text.grid(row = 1, column = 2, sticky = 'w')
        else:
            stand_text = t.Label(self.top_frame, bg = FORESTGREEN, font = font14Cb, bd = 5, text = self.stand)
            stand_text.grid(row = 1, column = 2, sticky = 'w')

        self.filler_plog.destroy()

        self.plog_label = t.Label(self.top_stands_hold_frame, bg = FORESTGREEN, font = font12Cb, bd = 5, text = self.plog_text)
        self.plog_label.grid(row = self.stands_row, column = 0, sticky = 'w')

        self.plog_entry = t.Entry(self.top_stands_hold_frame, font = font12Cb, width = 10)
        self.plog_entry.grid(row = self.stands_row, column = 1, sticky = "w")

        self.plog_button = t.Button(self.top_stands_hold_frame, font = font12Cb, width = 10, text = "Submit", bd = 2, command = lambda: self.plog_submit(self.plog_entry.get()))
        self.plog_button.grid(row = self.stands_row, column = 2, sticky = "w")

        self.stands_row += 1

        self.first_time_running = False


        
    def plog_submit(self, Plog):
        plog = Plog
        try:
            self.mlog_button.destroy()
            self.view_button.destroy()
            self.print_button.destroy()
        except:
            pass           
            
        if plog == '':
            messagebox.showwarning("Warning", "Please Enter a Log Length", parent = self.top)
            return            
        try:
            if int(plog) <= 2 or int(plog) > 120:
                messagebox.showwarning("Warning", "Invalid Log Length", parent = self.top)
                return
            else:
                self.plog = int(plog)
                
                self.mlog_label = t.Label(self.top_stands_hold_frame, bg = FORESTGREEN, font = font12Cb, bd = 5, text = "Minimum Log Length:")
                self.mlog_label.grid(row = self.stands_row, column = 0, sticky = 'w')

                self.mlog_entry = t.Entry(self.top_stands_hold_frame, font = font12Cb, width = 10)
                self.mlog_entry.grid(row = self.stands_row, column = 1, sticky = "w")

                self.mlog_button = t.Button(self.top_stands_hold_frame, font = font12Cb, width = 10, text = "Submit", bd = 2, command = lambda: self.mlog_submit(self.mlog_entry.get()))
                self.mlog_button.grid(row = self.stands_row, column = 2, sticky = "w")

                self.mlog_label = t.Label(self.top_stands_hold_frame, bg = FORESTGREEN, font = font8Cb, bd = 5, text = "*Set to Zero for full scaling*")
                self.mlog_label.grid(row = self.stands_row + 1, column = 0, sticky = 'w')
                
        except ValueError:
            messagebox.showwarning("Warning", "Log Length must be a number", parent = self.top)
            return

    
    def mlog_submit(self, Mlog):
        mlog = Mlog
        if mlog == '':
            messagebox.showwarning("Warning", "Please Enter a Log Length", parent = self.top)
            return            
        try:
            if int(mlog) < 0 or int(mlog) >= self.plog:
                messagebox.showwarning("Warning", "Invalid Log Length",parent = self.top)
                return
            else:
                self.mlog = int(mlog)
                self.mlog_button.destroy()
                self.view_button = t.Button(self.top_stands_hold_frame, font = font10Cb, width = 10, text = "View Report", bd = 2, command = lambda: self.generate_report())
                self.view_button.grid(row = self.stands_row, column = 2, sticky = "w")

                self.print_button = t.Button(self.top_stands_hold_frame, font = font10Cb, width = 10, text = "Print Report", bd = 2, command = lambda: self.print_report())
                self.print_button.grid(row = self.stands_row, column = 3, sticky = "w")
                
        except ValueError:
            messagebox.showwarning("Warning", "Log Length must be a number", parent = self.top)
            return



    def generate_report(self):
        self.top_frame.destroy()
        self.report_frame_format()

        blank = t.Label(self.top_frame, bg = FORESTGREEN, font = font12Cb, bd = 5, text = "")
        blank.grid(row = 0, column = 0, sticky = 'w')
        
        self.loading_label = t.Label(self.top_frame, bg = FORESTGREEN, font = font12Cb, bd = 5, text = "LOADING...")
        self.loading_label.grid(row = 0, column = 1, sticky = 'w')
        
        self.top_frame.update()

        if self.show_all_stands:            
            row = 0
            for i in self.stands:
                self.stand = i[0]
                self.plots = i[1]
                row = self.format_report(row)
            
        else:
            self.format_report(0)           

        self.top_frame.update()
        self.loading_label.destroy()
        self.top_canvas['scrollregion'] = self.top_frame.bbox("all")
        



    def format_report(self, row):      
        
        blank = t.Label(self.top_frame, bg = FORESTGREEN, font = font12Cb, bd = 5, text = "")
        blank.grid(row = row, column = 0, sticky = 'w')
        row += 1
        
        stand_for = t.Label(self.top_frame, bg = FORESTGREEN, font = font14Cb, bd = 5, text = 'STAND REPORT FOR:')
        stand_for.grid(row = row, column = 1, sticky = 'w')
        
        stand_text = t.Label(self.top_frame, bg = FORESTGREEN, font = font14Cb, bd = 5, text = self.stand)
        stand_text.grid(row = row, column = 2, sticky = 'w')
        row += 1

        self.stand_generated = True
        self.reporter = Report(self.csv, self.stand, self.plots, self.plog, self.mlog)
        
        blank = t.Label(self.top_frame, bg = FORESTGREEN, font = font10Cb, bd = 5, text = "")
        blank.grid(row = row, column = 0, sticky = 'w')
        row += 1

        
        #Summary Tables
        head = t.Label(self.top_frame, bg = FORESTGREEN, font = font14Cb, bd = 5, text = 'STAND SUMMARY TABLE')
        head.grid(row = row, column = 1, sticky = 'w')
        row += 1

        table_list = ['SPECIES', 'TPA', 'BA', 'QMD', 'RD', 'AVG HGT', 'HDR', 'VBAR', 'BOARD FEET', 'CUBIC FEET']

        col_width = 12
        for i in range(1, len(table_list) + 1):
            table_head = t.Label(self.top_frame, bg = FORESTGREEN, font = font12Cb, bd = 5, width = col_width, text = table_list[i - 1])
            table_head.grid(row = row, column = i, sticky = 'n')

        sep = ttk.Separator(self.top_frame, orient = 'horizontal').grid(column = 1, row = row, columnspan = len(table_list), sticky = 'nwe')

        sep_start = row
        row += 1     

        
        for key in self.reporter.conditions_dict:
            data = self.reporter.conditions_dict[key]
            
            if key == 'TOTALS':
                spp = 'TOTALS'
                blank = t.Label(self.top_frame, bg = FORESTGREEN, font = font10Cb, bd = 5, text = "")
                blank.grid(row = row, column = 0, sticky = 'w')
                row += 1
            else:
                spp = Timber.ALL_SPECIES_NAMES[key]

            spp_label = t.Label(self.top_frame, bg = FORESTGREEN, font = font10Cb, bd = 5, text = spp)
            spp_label.grid(row = row, column = 1, sticky = 'w')

            for i in range(2, len(table_list) + 1):
                data_text = self.format_data_text(data[i - 2])
                data_label = t.Label(self.top_frame, bg = FORESTGREEN, font = font10Cb, bd = 5, text = data_text)
                data_label.grid(row = row, column = i, sticky = 'n')

            
            row += 1

        sep_end = row
        for i in range(len(table_list) + 1):            
            sep = ttk.Separator(self.top_frame, orient = 'vertical').grid(column = i, row = sep_start, rowspan = (sep_end - sep_start), sticky = 'ens')
        sep = ttk.Separator(self.top_frame, orient = 'horizontal').grid(column = 1, row = sep_end, columnspan = len(table_list), sticky = 'nwe')

        for i in range(2):
            blank = t.Label(self.top_frame, bg = FORESTGREEN, font = font10Cb, bd = 5, text = "")
            blank.grid(row = row, column = 0, sticky = 'w')
            row += 1

        disclaimer = t.Label(self.top_frame, bg = FORESTGREEN, font = font10Cb, bd = 5, text = "Note: Log Grades do not consider defect or knot density")
        disclaimer.grid(row = row, column = 1, columnspan = 2, sticky = 'w')
        row += 1
        

        #Logs Tables
        header11 = 'LOGS COUNT SUMMARY TABLE'
        header12 = 'LOGS PER ACRE'

        header21 = 'LOG BOARD FEET SUMMARY TABLE'
        header22 = 'LOG GRADE BF PER ACRE'

        header31 = 'LOG CUBIC FEET SUMMARY TABLE'
        header32 = 'LOG GRADE CF PER ACRE'

        row = self.format_logs(row, header11, header12, self.reporter.logs_dict, 0)
        row = self.format_logs(row, header21, header22, self.reporter.logs_dict, 1)
        row = self.format_logs(row, header31, header32, self.reporter.logs_dict, 2)

        for i in range(row + 1, row + 10):
            blank = t.Label(self.top_frame, bg = FORESTGREEN, font = font10Cb, bd = 5, text = "")
            blank.grid(row = i, column = 0, sticky = 'w')
            row += 1        

        return row
            




    def format_data_text(self, Data_to_Format):
        data = str(round(Data_to_Format, 1))
        if data == '0':
            data = '-'
            return data
        
        if len(data) > 5:
            char_list = []
            for i in data:
                char_list.append(i)               
            char_list.insert(-5, ',')
            data = ''
            for i in char_list:
                data += i            
            return data
        
        else:
            return data
        


    def format_logs(self, Start_Row, Header1, Header2, Dict_to_Format, Index):
        row = Start_Row
        data_dict = Dict_to_Format
        
        blank = t.Label(self.top_frame, bg = FORESTGREEN, font = font10Cb, bd = 5, text = "")
        blank.grid(row = row, column = 0, sticky = 'w')
        row += 1

        log_lengths_head = ['LOG LENGTHS', '41+ ft', '31-40 ft', '21-30 ft', '11-20 ft', '1-10 ft', 'TOTALS', 'BY GRADE']

        head1 = t.Label(self.top_frame, bg = FORESTGREEN, font = font14Cb, bd = 5, text = Header1)
        head1.grid(row = row, column = 1, columnspan = 2, sticky = 'w')
        row += 1

        sep_start = row
        
        for i in range(1, len(log_lengths_head) - 1):
            head = t.Label(self.top_frame, bg = FORESTGREEN, font = font12Cb, bd = 5, text = log_lengths_head[i - 1])
            head.grid(row = row, column = i, sticky = 'n')

        totals_head = t.Label(self.top_frame, bg = FORESTGREEN, font = font12Cb, bd = 5, text = log_lengths_head[-2])
        totals_head.grid(row = row, column = len(log_lengths_head) - 1, sticky = 'n')
        
        bygrade_head = t.Label(self.top_frame, bg = FORESTGREEN, font = font12Cb, bd = 5, text = log_lengths_head[-1])
        bygrade_head.grid(row = row + 1, column = len(log_lengths_head) - 1, sticky = 'n')

        sep = ttk.Separator(self.top_frame, orient = 'horizontal').grid(column = 1, row = row, columnspan = 7, sticky = 'nwe')
        row += 1

        head2 = t.Label(self.top_frame, bg = FORESTGREEN, font = font10Cb, bd = 5, text = Header2)
        head2.grid(row = row, column = 1, sticky = 'w')
        row += 1
        
        lrl = ["40+ ft", "31-40 ft", "21-30 ft", "11-20 ft", "1-10 ft", 'TGRD']
        
        for spp in data_dict:

            if spp == 'TOTALS':
                spp_label = t.Label(self.top_frame, bg = FORESTGREEN, font = font12Cb, bd = 5, text = spp)
                spp_label.grid(row = row, column = 1, sticky = 'w')
            else:            
                spp_label = t.Label(self.top_frame, bg = FORESTGREEN, font = font12Cb, bd = 5, text = Timber.ALL_SPECIES_NAMES[spp])
                spp_label.grid(row = row, column = 1, sticky = 'w')
                
            sep = ttk.Separator(self.top_frame, orient = 'horizontal').grid(column = 1, row = row, columnspan = len(lrl) + 1, sticky = 'nwe')
            
            row += 1
            
            for grade in data_dict[spp]:
                data = data_dict[spp][grade]

                if grade == 'TTL':
                    txt = 'TOTALS BY LENGTH'
                else:
                    txt = Timber.GRADE_NAMES[grade]
                
                grade_label = t.Label(self.top_frame, bg = FORESTGREEN, font = font10Cb, bd = 5, text = txt)
                grade_label.grid(row = row, column = 1, sticky = 'w')
                
                for i in range(2, len(lrl) + 2):
                    
                    text = self.format_data_text(data[lrl[i - 2]][Index])
                    data_label = t.Label(self.top_frame, bg = FORESTGREEN, font = font10Cb, bd = 5, text = text)
                    data_label.grid(row = row, column = i, sticky = 'n')


                row += 1

            
        sep_end = row
        for i in range(8):            
            sep = ttk.Separator(self.top_frame, orient = 'vertical').grid(column = i, row = sep_start, rowspan = (sep_end - sep_start), sticky = 'ens')
        sep = ttk.Separator(self.top_frame, orient = 'horizontal').grid(column = 1, row = sep_end, columnspan = len(lrl) + 1, sticky = 'nwe')

        return row + 1



    ## PRINT REPORT TO PDF
    def print_report(self):
        pdf_save = filedialog.asksaveasfilename(initialdir =  self.program.path, title = "Save PDF File",
                                                  filetypes = (("PDF", "*.pdf"), ("All Files", "*.*")))

        if pdf_save == '':
            return
        else:           
            pdf_save = self.program.extension_check(pdf_save, '.pdf')
            
            pdf = PDF()
            pdf.alias_nb_pages()

            if self.show_all_stands:
                for i in self.stands:
                    pdf.add_page()
                    self.stand = i[0]
                    self.plots = i[1]
                    self.format_print_report(pdf)
            else:
                pdf.add_page()
                self.format_print_report(pdf)           
            

            filename = pdf_save        
            pdf.output(filename, 'F')

            question = messagebox.askyesno('Open PDF Report', 'Report Generated.\n\nWould you like open the Report?', parent = self.top)
            if question:
                os.startfile(filename)
                return
            else:
                return



    def format_print_report(self, pdf):
        self.reporter = Report(self.csv, self.stand, self.plots, self.plog, self.mlog)
        
        pdf.set_font('Times', 'B', 14)
        height = 8

        pdf.cell(150, height, 'STAND REPORT FOR: ' + self.stand, 0, 0, align = 'L')
        pdf.ln(height)

        pdf.set_font('Times', 'B', 12)
        plog_text = 'Preferred Log Length: ' + str(self.plog) + ' feet'
        pdf.cell(100, height, plog_text, 0, 0, align = 'L')
        pdf.ln(height)
        
        mlog_text = 'Minimum Log Length: ' + str(self.mlog) + ' feet'
        pdf.cell(100, height, mlog_text, 0, 0, align = 'L')
        pdf.ln(height * 2)

        # Summary Table
        pdf.cell(50, height, 'STAND SUMMARY TABLE', 0, 0, align = 'L')

        table_list = ['SPECIES', 'TPA', 'BA', 'QMD', 'RD', 'AVG HGT', 'HDR', 'VBAR', 'BOARD FEET', 'CUBIC FEET']
        pdf.ln(height)


        spp_col = 38
        end_col = 21
        dat_col = 16

        for i in range(len(table_list)):
            pdf.set_font('Times', 'B', 9)
            if i == 0:
                pdf.cell(spp_col, height, table_list[i], 1, 0, align = 'L')
            elif i == len(table_list) - 1 or i == len(table_list) - 2:
                pdf.set_font('Times', 'B', 8)
                pdf.cell(end_col, height, table_list[i], 1, 0, align = 'C')
            else:
                pdf.cell(dat_col, height, table_list[i], 1, 0, align = 'C')

        pdf.ln(height)
        pdf.set_font('Times', '', 9)
        
        for key in self.reporter.conditions_dict:
            data = self.reporter.conditions_dict[key]
            if key == 'TOTALS':
                pdf.set_font('Times', 'B', 9)
                pdf.cell(spp_col, height, 'TOTALS', 1, 0, align='L')
            else:
                pdf.set_font('Times', '', 8)
                pdf.cell(spp_col, height, Timber.ALL_SPECIES_NAMES[key], 1, 0, align = 'L')

            pdf.set_font('Times', '', 9)
            for i in range(len(data)):
                show_text = self.format_data_text(data[i])
                if i == len(data) - 1 or i == len(data) - 2:
                    pdf.cell(end_col, height, show_text, 1, 0, align = 'C')
                else:
                    pdf.cell(dat_col, height, show_text, 1, 0, align = 'C')
            pdf.ln(height)
            pdf.set_font('Times', '', 9)
            

        pdf.ln(height * 2)
        pdf.cell(50, height, "Note: Log Grades do not consider defect or knot density", 0, 0, align = 'L')


        # Logs Summary
        pdf.ln(height)
        pdf.set_font('Times', 'B', 12)
        pdf.cell(75, height, 'LOGS AND GRADES SUMMARY TABLES', 0, 0, align = 'L')

        header1 = 'Log Count per Acre by Grade'
        header2 = 'Log Board Feet per Acre by Grade'
        header3 = 'Log Cubic Feet per Acre by Grade'

        self.format_logs_print_report(pdf, header1, self.reporter.logs_dict, 0)
        pdf.ln(height * 2)
        self.format_logs_print_report(pdf, header2, self.reporter.logs_dict, 1)
        pdf.ln(height * 2)
        self.format_logs_print_report(pdf, header3, self.reporter.logs_dict, 2)

        

    def format_logs_print_report(self, PDF, Header, Log_Dict, Index):
        pdf = PDF
        height = 8
        pdf.set_font('Times', 'B', 11)
        pdf.ln(height)
        pdf.cell(50, height, Header, 0, 0, align = 'L')
        pdf.set_font('Times', 'B', 10)
        pdf.ln(height)
        
        log_lengths_head = ['LOG LENGTHS', '41+ ft', '31-40 ft', '21-30 ft', '11-20 ft', '1-10 ft', 'TOTALS BY GRADE']

        grd_tot_col = 40
        rng_col = 22
        spp_span = 190

        for i in range(len(log_lengths_head)):
            if i == 0:
                pdf.cell(grd_tot_col, height, log_lengths_head[i], 1, 0, align = 'L')
            elif i == len(log_lengths_head) - 1:
                pdf.cell(grd_tot_col, height, log_lengths_head[i], 1, 0, align = 'C')
            else:
                pdf.cell(rng_col, height, log_lengths_head[i], 1, 0, align = 'C')

        pdf.set_font('Times', '', 9)
        pdf.ln(height)
        
        lrl = ["40+ ft", "31-40 ft", "21-30 ft", "11-20 ft", "1-10 ft", 'TGRD']

        log_dict = Log_Dict
        for spp in log_dict:

            if spp== 'TOTALS':
                pdf.set_font('Times', 'B', 9)
                pdf.cell(spp_span, height, spp, 1, 0, align = 'C')
                pdf.ln(height)
            else:
                pdf.set_font('Times', 'B', 9)
                pdf.cell(spp_span, height, Timber.ALL_SPECIES_NAMES[spp], 1, 0, align = 'C')
                pdf.ln(height)
                pdf.set_font('Times', '', 9)           

            
            for grade in log_dict[spp]:
                data = log_dict[spp][grade]
                pdf.set_font('Times', '', 9)

                if grade == 'TTL':
                    pdf.cell(grd_tot_col, height, 'TOTALS BY LENGTH', 1, 0, align = 'L')
                else:
                    pdf.set_font('Times', '', 8)
                    pdf.cell(grd_tot_col, height, Timber.GRADE_NAMES[grade], 1, 0, align = 'L')
                
                for i in range(len(lrl)):
                    show_text = self.format_data_text(data[lrl[i]][Index])
                    if i == len(lrl) - 1:
                        pdf.cell(grd_tot_col, height, show_text, 1, 0, align='C')
                    else:
                        pdf.cell(rng_col, height, show_text, 1, 0, align = 'C')

                pdf.ln(height)

            #pdf.ln(height)################################################
            pdf.set_font('Times', '', 9)




    ## TO DATABASE FUNCTIONS
    def export_db_start(self, Database_Type):
        if Database_Type == 'Access':
            self.choose_access = True
            self.choose_sqlite3 = False
            self.choose_excel = False
            self.choose_multiple = False

        elif Database_Type == 'SQLite3':
            self.choose_access = False
            self.choose_sqlite3 = True
            self.choose_excel = False
            self.choose_multiple = False

        elif Database_Type == 'Excel':
            self.choose_access = False
            self.choose_sqlite3 = False
            self.choose_excel = True
            self.choose_multiple = False

        elif Database_Type == 'Multiple':
            self.choose_access = False
            self.choose_sqlite3 = False
            self.choose_excel = False
            self.choose_multiple = True

        self.current_stand = -1
        self.to_db()
        

    def to_db(self):
        self.current_stand += 1
        self.top_stands_hold_frame.destroy()
        self.top_stands_hold_frame = t.Frame(self.top_stands_frame, bg = FORESTGREEN)
        self.top_stands_hold_frame.pack(fill = 'both', expand = 1)

        back_button = t.Button(self.top_stands_hold_frame, font = font8Cb, text = "Back to Stands", bd = 4, command = lambda: self.stands_frame_format())
        back_button.grid(row = 0, column = 0, sticky = "w")

        label = t.Label(self.top_stands_hold_frame, bg = FORESTGREEN, font = font14Cb, bd = 5, text = 'INPUTS FOR:  ' + self.stands[self.current_stand][0])
        label.grid(row = 0, column = 1, sticky = 'w')
        
        self.stands_row = 1

        if self.current_stand == len(self.stands) - 1:
            if self.choose_access:
                txt = 'Export to Access'
            elif self.choose_sqlite3:
                txt = 'Export to SQLite3'
            elif self.choose_excel:
                txt = 'Export to Excel'
            elif self.choose_multiple:
                txt = 'Export to Databases'
                self.acv = t.IntVar()
                self.sqv = t.IntVar()
                self.exv = t.IntVar()

                ac_chk = t.Checkbutton(self.top_stands_hold_frame, text="Access DB", font = font10Cb, bg=FORESTGREEN, justify='left', variable=self.acv)
                ac_chk.grid(row=self.stands_row, column=0, sticky='w')

                sq_chk = t.Checkbutton(self.top_stands_hold_frame, text="SQLite3 DB", font = font10Cb, bg=FORESTGREEN, justify='left', variable=self.sqv)
                sq_chk.grid(row=self.stands_row + 1, column=0, sticky='w')

                ex_chk = t.Checkbutton(self.top_stands_hold_frame, text="Excel DB", font = font10Cb, bg=FORESTGREEN, justify='left', variable=self.exv)
                ex_chk.grid(row=self.stands_row + 2, column=0, sticky='w')

            db_button = t.Button(self.top_stands_hold_frame, font = font10Cb, text = txt, bd = 4, command = lambda: self.create_db())
            db_button.grid(row = self.stands_row + 3, column = 0, sticky = "w")


        else:
            next_button = t.Button(self.top_stands_hold_frame, font = font10Cb, text = "Next Stand\t", bd = 4, command = lambda: self.create_db())
            next_button.grid(row = self.stands_row + 3, column = 0, sticky = "w")
        
        col = 1
        self.access_data_dict = {}
        for i in STANDS_INPUT:
            for j in i:
                self.access_data_dict[j] =  ''
        
        for i in STANDS_INPUT[0]:
            if i == 'Variant':
                stands_label = t.Label(self.top_stands_hold_frame, bg = FORESTGREEN, font = font12Cb, bd = 5, text = i, width = 15)
                stands_label.grid(row = self.stands_row, column = col, sticky = 'n')
                
                var_menu_list = []
                for key in VARIANTS_LOCS:
                    text = key + ": " + VARIANTS_LOCS[key][0]
                    var_menu_list.append(text)
                        
                var_header = t.StringVar()
                var_header.set("Variants")
                
                self.var_menu = t.OptionMenu(self.top_stands_hold_frame, var_header, (*var_menu_list), command = self.variant_menu)
                self.var_menu.grid(row = self.stands_row + 1, column = col, sticky = "we")
                
            elif i == 'Forest Code':
                stands_label = t.Label(self.top_stands_hold_frame, bg = FORESTGREEN, font = font12Cb, bd = 5, text = i)
                stands_label.grid(row = self.stands_row, column = col, sticky = 'n')
                        
                self.for_header = t.StringVar()
                self.for_header.set(i + 's')

                self.for_menu_list = ['']
                
                self.for_menu = t.OptionMenu(self.top_stands_hold_frame, self.for_header, (*self.for_menu_list), command = self.forest_menu)
                self.for_menu.grid(row = self.stands_row + 1, column = col, sticky = "we")

            else:
                stands_label = t.Label(self.top_stands_hold_frame, bg = FORESTGREEN, font = font12Cb, bd = 5, text = i)
                stands_label.grid(row = self.stands_row, column = col, sticky = 'n')
                        
                self.reg_header = t.StringVar()
                self.reg_header.set(i + 's')

                self.reg_menu_list = ['']
                
                self.reg_menu = t.OptionMenu(self.top_stands_hold_frame, self.reg_header, (*self.reg_menu_list), command = self.region_menu)
                self.reg_menu.grid(row = self.stands_row + 1, column = col, sticky = "we")
                
            col += 1

        col = 1
        self.stands_row += 2
        
        for i in STANDS_INPUT[1]:
            stands_label = t.Label(self.top_stands_hold_frame, bg = FORESTGREEN, font = font12Cb, bd = 5, text = i)
            stands_label.grid(row = self.stands_row, column = col, sticky = 'n')

            if i == 'Inventory Year':
                stands_entry = t.Entry(self.top_stands_hold_frame, font = font14Cb, width = 20)
                stands_entry.grid(row = self.stands_row + 1, column = col, sticky = "w")
            else:
                stands_entry = t.Entry(self.top_stands_hold_frame, font = font14Cb, width = 12)
                stands_entry.grid(row = self.stands_row + 1, column = col, sticky = "w")
                
            col += 1


    def variant_menu(self, Selection):
        self.var_menu['fg'] = BLACK
        key = ''
        for i in range(2):
            key += Selection[i]

        self.access_data_dict['Variant'] = key

        self.for_menu['menu'].delete(0, 'end')
        self.reg_menu['menu'].delete(0, 'end')
        
        for i in VARIANTS_LOCS[key][1]:
            self.for_menu['menu'].add_command(label = i, command = t._setit(self.for_header, i, self.forest_menu))                


    def forest_menu(self, Selection):
        self.for_menu['fg'] = BLACK
        self.access_data_dict['Forest Code'] = Selection
        
        if Selection in VARIANTS_LOCS['AK'][1]:
            reg = 10
        else:
            reg = int(str(Selection)[0])
            
        self.reg_menu['menu'].delete(0, 'end')
        self.reg_menu['menu'].add_command(label = reg, command = t._setit(self.reg_header, reg, self.region_menu))


    def region_menu(self, Selection):
        self.reg_menu['fg'] = BLACK
        self.access_data_dict['Region Code'] = Selection
        

    ## ERROR CHECKING
    def entry_error(self):
        current = self.top.focus_get()
        if current['fg'] == RED:
            current['fg'] = BLACK
        current['textvariable'] = None
        

    def error_check(self):
        data_sheet = list(self.top_stands_hold_frame.children.values())
        
        col_count = 1
        
        var_err_code = 0
        for_err_code = 0
        reg_err_code = 0

        #ERROR CODES =  0: No Error
        #               1: Missing
        #               2: Not In Variant or Forest

        
        err_list = []
        err_count = 0
        
        if self.access_data_dict['Variant'] == '':            
            var_err_code = 1
            err_count += 1
            
        if self.access_data_dict['Forest Code'] == '':
            for_err_code = 1
            err_count += 1
            
        elif self.access_data_dict['Forest Code'] not in VARIANTS_LOCS[self.access_data_dict['Variant']][1]:
            for_err_code = 2
            err_count += 1

        if self.access_data_dict['Region Code'] == '':
            reg_err_code = 1
            err_count += 1
            
        elif self.access_data_dict['Region Code'] != 10 and str(self.access_data_dict['Region Code']) != str(self.access_data_dict['Forest Code'])[0]:
            reg_err_code = 2
            err_count += 1

        elif self.access_data_dict['Region Code'] == 10 and str(self.access_data_dict['Region Code']) != (str(self.access_data_dict['Forest Code'])[0] + str(self.access_data_dict['Forest Code'])[1]): 
            reg_err_code = 2
            err_count += 1

        if err_count > 0:
            warn_text = ''
            if var_err_code == 1:
                warn_text += 'Variant needs to be selected\n\n'

            if for_err_code == 1:
                warn_text += 'Forest needs to be selected\n\n'
            elif for_err_code == 2:
                warn_text += 'Forest Code not in selected Variant\n\n'

            if reg_err_code == 1:
                warn_text += 'Region needs to be selected'
            elif reg_err_code == 2:
                warn_text += 'Region Code does not match Forest Code'
            
            messagebox.showwarning("Warning", warn_text, parent = self.top)
            if var_err_code > 0:            
                self.var_menu['fg'] = RED
                
            if for_err_code > 0:            
                self.for_menu['fg'] = RED

            if reg_err_code > 0:
                self.reg_menu['fg'] = RED
                
            return

        else:       
            for i in range(len(data_sheet)):            
                val = data_sheet[i]
                required_cols = [1, 4, 5, 6, 7]
                
                if type(val) == t.Entry:
                    if val.get() == '' and col_count in required_cols:
                        self.program.error_append(val, err_list, "Missing")
                    else:
                        if not self.type_check(val, col_count):
                            self.program.error_append(val, err_list, "Val Err")
                        else:
                            if col_count == 1:
                                self.access_data_dict['Inventory Year'] = int(val.get())
                                
                            elif col_count == 2:
                                if val.get() == '':
                                    default = self.access_data_dict['Variant']
                                    self.access_data_dict['Latitude'] = LAT_LONG_DEFAULT[default][0]
                                else:
                                    self.access_data_dict['Latitude'] = float(val.get())
                                    
                            elif col_count == 3:
                                if val.get() == '':
                                    default = self.access_data_dict['Variant']
                                    self.access_data_dict['Longitude'] = LAT_LONG_DEFAULT[default][1]
                                else:
                                    self.access_data_dict['Longitude'] = float(val.get())
                                    
                            elif col_count == 4:
                                self.access_data_dict['Age'] = int(val.get())

                            elif col_count == 5:
                                self.access_data_dict['Elevation'] = int(val.get())

                            elif col_count == 6:
                                self.access_data_dict['Site Species'] = val.get().upper()

                            elif col_count == 7:
                                self.access_data_dict['Site Index'] = int(val.get()) 
                                
                    col_count += 1
                    
            if len(err_list) > 0:
                messagebox.showwarning("Warning", "One or more required values\nare missing or incorrect", parent = self.top)
                for i in err_list:
                    i[0].config(textvariable = i[1])
                    i[0].config(fg = RED)                
                return False
            else:
                self.stands_access_dict[self.stands[self.current_stand][0]] = self.access_data_dict
                return True

    # Data type check    
    def type_check(self, Entry, Column_Count):
        int_cols = [1, 4, 5, 7]
        float_cols = [2, 3]

        val = Entry.get()
        if Column_Count in int_cols:
            try:
                x = int(val)
                x = math.sqrt(int(val))
                return True                
            except:
                return False 
            
        elif Column_Count in float_cols and val != '':
            try:
                x = float(val)
                x = math.sqrt(float(val))
                return True
            except:
                return False

        elif Column_Count == 6:
            if val.upper() in Timber.ALL_SPECIES_NAMES:
                return True
            else:
                return False
        else:
            return True
        


    ## EXPORTING TO FVS DATABASE
    def get_plot_factor(self):
        factor_dict = {}
            
        with open(self.csv, 'r') as csv_data:    
            csv_data_reader = csv.reader(csv_data)    
            next(csv_data_reader)

            for line in csv_data_reader:
                if line[0].upper() not in factor_dict:
                    factor_dict[line[0].upper()] = float(line[6])

        return factor_dict
    


    def create_db_table(self, Table_Dict, Index):
        text = """CREATE TABLE """
        for key in Table_Dict:
            text += key + " ("
            for i, data in enumerate(Table_Dict[key]):
                if i == len(Table_Dict[key]) - 1:
                    text += data[0] + " " + data[Index] + ");"
                else:
                    text += data[0] + " " + data[Index] + ", "
        return text



    def create_db_stands(self):
        master = []
        factor_dict = self.get_plot_factor()       
        
        for stand in self.stands:
            temp_list = []
            temp_list.append(stand[0])
            temp_list.append(self.stands_access_dict[stand[0]]['Variant'])
            temp_list.append(self.stands_access_dict[stand[0]]['Inventory Year'])
            temp_list.append('All_Stands')
            temp_list.append(self.stands_access_dict[stand[0]]['Latitude'])
            temp_list.append(self.stands_access_dict[stand[0]]['Longitude'])
            temp_list.append(self.stands_access_dict[stand[0]]['Region Code'])
            temp_list.append(self.stands_access_dict[stand[0]]['Forest Code'])
            temp_list.append(self.stands_access_dict[stand[0]]['Age'])
            temp_list.append(self.stands_access_dict[stand[0]]['Elevation'])
            temp_list.append(factor_dict[stand[0]])                
            temp_list.append(0)
            temp_list.append(stand[1])
            temp_list.append(self.stands_access_dict[stand[0]]['Site Species'])
            temp_list.append(self.stands_access_dict[stand[0]]['Site Index'])
            master.append(temp_list)
            
        return master


    def create_db_trees(self):
        master = []
        with open(self.csv, 'r') as csv_data:    
            csv_data_reader = csv.reader(csv_data)    
            next(csv_data_reader)
            
            for line in csv_data_reader:
                temp_list = []
                if line[0] == '':
                    break
                else:
                    temp_list.append(line[0].upper())
                    temp_list.append(line[0].upper() + "_" + line[1])
                    temp_list.append(int(line[1]))
                    temp_list.append(int(line[2]))
                    temp_list.append(1)
                    temp_list.append(1)
                    temp_list.append(line[3].upper())
                    temp_list.append(float(line[4]))
                    if line[5] == '':
                        temp_list.append(None)
                    else:
                        temp_list.append(float(line[5]))
                    master.append(temp_list)
                    
        return master           
            


    def create_db(self):
        if not self.error_check():
            return
        else:
            if self.current_stand < len(self.stands) - 1:
                self.to_db()
            else:
                if self.choose_access:
                    db_save = filedialog.asksaveasfilename(initialdir =  self.program.path, title = "Save Access Database File",
                                                             filetypes = (("Microsoft Access", "*.accdb"), ("All Files", "*.*")))

                elif self.choose_sqlite3:
                    db_save = filedialog.asksaveasfilename(initialdir = self.program.path, title = "Save SQLite3 Database File",
                                                             filetypes = (("Database File", "*.db"), ("All Files", "*.*")))

                elif self.choose_excel:
                    db_save = filedialog.asksaveasfilename(initialdir = self.program.path, title = "Save Excel Database File",
                                                             filetypes= (("Excel File", "*.xlsx"), ("All Files", "*.*")))

                elif self.choose_multiple:
                    if not bool(self.acv.get()) and not bool(self.sqv.get()) and not bool(self.exv.get()):
                        messagebox.showwarning("Warning", "Please choose at least one Database",
                                               parent=self.top)
                        return
                    else:
                        db_save = filedialog.asksaveasfilename(initialdir = self.program.path, title = "Save Database Filename",
                                                               filetypes= (("File Name", "*.file"), ("All Files", "*.*")))

                if db_save == '':
                    return
                else:
                    stands_list = self.create_db_stands()
                    trees_list = self.create_db_trees()


                    if self.choose_access:
                        self.generate_access(stands_list, trees_list, db_save)

                    elif self.choose_sqlite3:
                        self.generate_sqlite3(stands_list, trees_list, db_save)

                    elif self.choose_excel:
                        self.generate_excel(stands_list, trees_list, db_save)

                    elif self.choose_multiple:
                        if bool(self.acv.get()):
                            self.generate_access(stands_list, trees_list, db_save)
                        if bool(self.sqv.get()):
                            self.generate_sqlite3(stands_list, trees_list, db_save)
                        if bool(self.exv.get()):
                            self.generate_excel(stands_list, trees_list, db_save)


    # Generating Access database for FVS
    def generate_access(self, stands_list, trees_list, db_name):
        db_save = self.program.extension_check(db_name, '.accdb')
        db_short = self.program.get_filename_only(db_save)

        groups_table = self.create_db_table(FVS_GROUPS, 1)
        stands_table = self.create_db_table(FVS_STANDS, 1)
        trees_table = self.create_db_table(FVS_TREES, 1)

        database_exists = True

        try:
            accApp = Dispatch("Access.Application")
            dbEngine = accApp.DBEngine
            workspace = dbEngine.Workspaces(0)

            dbLangGeneral = ';LANGID=0x0409;CP=1252;COUNTRY=0'
            newdb = workspace.CreateDatabase(db_save, dbLangGeneral, 64)

            newdb.Execute(groups_table)
            newdb.Execute(stands_table)
            newdb.Execute(trees_table)

            database_exists = False

        except Exception as e:
            pass

        finally:
            accApp.DoCmd.CloseDatabase
            accApp.Quit
            newdb = None
            workspace = None
            dbEngine = None
            accApp = None

        drive = '{Microsoft Access Driver (*.mdb, *.accdb)}'

        connection_text = r'Driver={driver};DBQ={filename};'.format(driver = drive, filename = db_save)
        connection = db.connect(connection_text)
        insert = connection.cursor()

        if not database_exists:
            keyword_text = """ INSERT INTO FVS_GroupAddFilesAndKeywords (Groups, FVSKeywords) VALUES (?, ?)"""

            formatted = FVS_KEYWORDS.format(FILLER_DB_NAME = db_short)
            FVS_GROUPS[FVS_G][2][3] = formatted

            insert.execute(keyword_text,
                           FVS_GROUPS[FVS_G][0][3], FVS_GROUPS[FVS_G][2][3])

        for i in stands_list:
            stand_text = self.insert_table_text(FVS_STANDS, i)
            insert.execute(stand_text)


        for i in trees_list:
            tree_text = self.insert_table_text(FVS_TREES, i)
            insert.execute(tree_text)


        connection.commit()
        connection.close()

        self.generate_loc(db_save)



    # Generating SQLite3 database for FVS
    def generate_sqlite3(self, stands_list, trees_list, db_name):
        db_save = self.program.extension_check(db_name, '.db')
        db_short = self.program.get_filename_only(db_save)


        if not os.path.isfile(db_save):

            groups_table = self.create_db_table(FVS_GROUPS, 2)
            stands_table = self.create_db_table(FVS_STANDS, 2)
            trees_table = self.create_db_table(FVS_TREES, 2)

            connection = lite.connect(db_save)
            insert = connection.cursor()

            connection.execute(groups_table)
            connection.execute(stands_table)
            connection.execute(trees_table)
            connection.commit()

            keyword_text = """ INSERT INTO FVS_GroupAddFilesAndKeywords (Groups, FVSKeywords) VALUES (?, ?)"""

            formatted = FVS_KEYWORDS.format(FILLER_DB_NAME = db_short)
            FVS_GROUPS[FVS_G][2][3] = formatted

            insert.execute(keyword_text,
                           [FVS_GROUPS[FVS_G][0][3],
                           FVS_GROUPS[FVS_G][2][3]])

            connection.commit()
            connection.close()

        connection = lite.connect(db_save)
        insert = connection.cursor()
        for i in stands_list:
            stand_text = self.insert_table_text(FVS_STANDS, i)
            insert.execute(stand_text)

        for i in trees_list:
            tree_text = self.insert_table_text(FVS_TREES, i)
            insert.execute(tree_text)

        connection.commit()
        connection.close()

        messagebox.showinfo('Completed',
                              'Completed exporting to SQLite3 Database',
                              parent=self.top)



    # Generating Excel database for FVS
    def generate_excel(self, stands_list, trees_list, db_name):
        db_save = self.program.extension_check(db_name, '.xlsx')

        if not os.path.isfile(db_save):
            self.insert_excel(db_save, stands_list, trees_list)

        else:
            try:
                data = xl.load_workbook(db_save)
                stands_sheet = data.get_sheet_by_name(FVS_S)
                for i in range(2, stands_sheet.max_row + 1):
                    temp_list = [stands_sheet.cell(i, j).value for j in range(1, stands_sheet.max_column + 1)]
                    stands_list.insert(i - 2, temp_list)

                trees_sheet = data.get_sheet_by_name(FVS_T)
                for i in range(2, trees_sheet.max_row + 1):
                    temp_list = [trees_sheet.cell(i, j).value for j in range(1, trees_sheet.max_column + 1)]
                    trees_list.insert(i - 2, temp_list)

                data.close()

            except:
                pass

            self.insert_excel(db_save, stands_list, trees_list)

        messagebox.showinfo('Completed',
                            'Completed exporting to Excel Database',
                            parent=self.top)



    # Inserting data into excel sheets
    def insert_excel(self, save_file, stands_list, trees_list):
        wb = xl.Workbook()

        groups = wb.create_sheet(FVS_G)

        db_short = self.program.get_filename_only(save_file)
        formatted = EXCEL_KEYWORDS.format(FILLER_DB_NAME = db_short)
        FVS_GROUPS[FVS_G][2][3] = formatted

        for i, fill in enumerate(FVS_GROUPS[FVS_G]):
            if i == len(FVS_GROUPS[FVS_G]) - 1:
                groups.cell(1, i + 1, fill[0])
                groups.cell(2, i + 1, fill[-1])
                groups.cell(2, i + 1).alignment = Alignment(wrapText = True)
            else:
                groups.cell(1, i+1, fill[0])
                groups.cell(2, i+1, fill[-1])


        stands = wb.create_sheet(FVS_S)
        for i, fill in enumerate(FVS_STANDS[FVS_S]):
            stands.cell(1, i+1, fill[0])

        for data in stands_list:
            stands.append(data)

        trees = wb.create_sheet(FVS_T)
        for i, fill in enumerate(FVS_TREES[FVS_T]):
            trees.cell(1, i+1, fill[0])

        for data in trees_list:
            trees.append(data)

        try:
            del wb['Sheet']
        except:
            pass

        wb.save(save_file)
        wb.close()




    # Formatting SQL statements for insertion into database
    def insert_table_text(self, Table_Dict, Value_List):
        text = """INSERT INTO """
        for key in Table_Dict:
            text += key + " ("
            for i, data in enumerate(Table_Dict[key]):
                if i == len(Table_Dict[key]) - 1:
                    text += data[0] + ') '
                else:
                    text += data[0] + ', '

        text += 'VALUES('

        val = Value_List
        for i, data in enumerate(val):
            if i == len(val) - 1:
                if data is None:
                    text += 'NULL);'
                elif isinstance(data, str):
                    text += "'" + data + "');"
                else:
                    text += '{var});'.format(var = data)
            else:
                if data is None:
                    text += 'NULL,'
                elif isinstance(data, str):
                    text += "'" + data + "',"
                else:
                    text += '{var},'.format(var = data)

        return text






    # Create or append suppose locations file for FVS
    def generate_loc(self, Access_Full_Filename):
        char_list = []
        inserting = False
        for i in reversed(Access_Full_Filename):
            if not inserting:
                if i == "/":
                    char_list.insert(0, i)
                    inserting = True
                else:
                    next
            else:
                char_list.insert(0, i)
        path_name = ''
        for i in char_list:
            path_name += i
                
        filename = path_name + 'Suppose.loc'
        name = self.program.get_filename_only(Access_Full_Filename)

        fill_text = 'C "'

        for i in name:
            if i == '.':
                break
            else:
                fill_text += i
        fill_text += '" ' + name

        try:
            f = open(filename, 'r')
            lines = f.readlines()
            in_loc = False
            for line in lines:
                if fill_text == str(line).strip('\n'):
                    in_loc = True
                    break

            if in_loc:
                f.close()
                messagebox.showinfo('Completed', 'Completed exporting to Access Database\n\nDatabase already in Suppose.loc file',
                                      parent = self.top)
                return
            else:
                f.close()
                f = open(filename, 'a')
                f.write('\n' + fill_text)
                f.close()
                messagebox.showinfo('Completed', 'Completed exporting to Access Database\n\nDatabase appended to Suppose.loc file',
                                      parent = self.top)
                return
        except:
            f = open(filename, 'w')
            f.write(fill_text)
            f.close()
            messagebox.showinfo('Completed', 'Completed exporting to Access Database\n\nCompleted creation of Suppose.loc file for use in FVS',
                                  parent = self.top)
            return




















