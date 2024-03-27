import tkinter.filedialog as filedialog
import tkinter as tk
from configparser import ConfigParser
from tkinter import ttk

from tools import Tools

# XMLData.Get_Analysis_Link(b)
# XMLData.Running()

class Interface_XML2Excel():
    def __init__(self):
        self.XML_file = None
        self.Excel_file = None

        self.root = tk.Tk()
        self.root.title('ReportAnalysis')
        self.root.geometry('600x200+450+200')
        # Init Frame for all sub-windows
        container = tk.Frame(self.root)
        container.pack(side="top", fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        # set up all sub-windows use the main frame and arrange them on the main object
        self.sub_windows = {}
        for sub_window in (Choose_Funcktion_Window, Choose_File_Window, Converting_Excel, Show_Progress_Window,
                           Choose_Measurement_Window):
            window_name = sub_window.__name__
            window = sub_window(parent=container, controller=self)
            self.sub_windows[window_name] = window
            window.grid(row=0, column=0, sticky="nsew")
        self.show_window("Choose_Funcktion_Window")

    def show_window(self, page_name):
        '''Show a frame for the given page name'''
        sub_window = self.sub_windows[page_name]
        sub_window.tkraise()

class Choose_Funcktion_Window(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        FrameTable = tk.Frame(self)
        frm_Top = tk.Frame(FrameTable)
        tk.Label(frm_Top, text='Porsche', font=('Porsche Design Font', 14)).pack(pady=15)
        frm_Top.pack(side=tk.TOP)
        frm_Bottom = tk.Frame(FrameTable)
        tk.Button(frm_Bottom, text='Report Generator', command=self.Report_Generator).pack(side=tk.LEFT, padx=10, pady=50)
        # tk.Button(frm_Bottom, text='Merge FSM and FEM', command=self.Merge_FSM_and_FEM).pack(side=tk.LEFT, padx=10, pady=50)
        frm_Bottom.pack(side=tk.BOTTOM)
        FrameTable.pack()

    def Report_Generator(self):
        self.controller.show_window('Choose_File_Window')

class Choose_File_Window(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        FrameTable = tk.Frame(self)
        frm_Top = tk.Frame(FrameTable)
        tk.Label(frm_Top, text='Porsche', font=('Porsche Design Font', 14)).pack(pady=15)
        frm_Top.pack(side=tk.TOP)
        frm_Bottom = tk.Frame(FrameTable)
        tk.Button(frm_Bottom, text='Choose XML Report', command=self.Choosefile).pack(side=tk.LEFT, pady=50)
        frm_Bottom.pack(side=tk.BOTTOM)
        FrameTable.pack()

    def Choosefile(self):
        config = ConfigParser()
        config.read(Tools.get_path('/Config.ini'))
        path = Tools.get_path('/Config.ini')
        print path
        XML_import_path = config.get('Section_ini', 'XML_import_path')
        XML_import_path = Tools.path_slash2backslash(XML_import_path)
        # print XML_import_path
        get_XML_file = filedialog.askopenfilename(title='Choose a XML Report', initialdir=XML_import_path,
                                        filetypes=[('XML Report (.xml)', '.xml')])
        self.controller.XML_file = Tools.path_slash2backslash(get_XML_file)
        if get_XML_file != '':
            config.set('Section_ini', 'XML_import_path', get_XML_file)
            config.write(open(Tools.get_path('/Config.ini'), 'w'))
            self.controller.show_window('Choose_Measurement_Window')

class Choose_Measurement_Window(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        FrameTable = tk.Frame(self)
        frm_Top = tk.Frame(FrameTable)
        tk.Label(frm_Top, text='Porsche', font=('Porsche Design Font', 14)).pack(pady=15)
        frm_Top.pack(side=tk.TOP)
        frm_Bottom = tk.Frame(FrameTable)
        tk.Button(frm_Bottom, text='Choose XML Report', command=self.Choosefile).pack(side=tk.LEFT, pady=50)
        frm_Bottom.pack(side=tk.BOTTOM)
        FrameTable.pack()

    def Choosefile(self):
        config = ConfigParser()
        config.read(Tools.get_path('/Config.ini'))
        path = Tools.get_path('/Config.ini')
        print path
        XML_import_path = config.get('Section_ini', 'XML_import_path')
        XML_import_path = Tools.path_slash2backslash(XML_import_path)
        # print XML_import_path
        get_XML_file = filedialog.askopenfilename(title='Choose a XML Report', initialdir=XML_import_path,
                                        filetypes=[('XML Report (.xml)', '.xml')])
        self.controller.XML_file = Tools.path_slash2backslash(get_XML_file)
        if get_XML_file != '':
            config.set('Section_ini', 'XML_import_path', get_XML_file)
            config.write(open(Tools.get_path('/Config.ini'), 'w'))
            self.controller.show_window('Converting_Excel')

class Converting_Excel(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        FrameTable = tk.Frame(self)
        frm_Top = tk.Frame(FrameTable)
        tk.Label(frm_Top, text='Porsche', font=('Porsche Design Font', 14)).pack(pady=15)
        frm_Top.pack(side=tk.TOP)
        frm_Bottom = tk.Frame(FrameTable)
        tk.Button(frm_Bottom, text='Get Excel Report', command=lambda: self.running_XML_GiL(self.controller.XML_file)).pack(side=tk.LEFT, pady=50)
        frm_Bottom.pack(side=tk.BOTTOM)
        FrameTable.pack()

    def running_XML_GiL(self, XML_file):
        from XML_Data import XMLData
        print XML_file
        self.controller.show_window('Show_Progress_Window')
        XMLData.Get_Analysis_Link(XML_file)
        XMLData.Running()
        self.controller.show_window("Choose_Funcktion_Window")
        print 'Finishing'

class Show_Progress_Window(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        Show_Progress_Window.root_window = controller.root
        self.controller = controller
        FrameTable = tk.Frame(self)
        frm_Top = tk.Frame(FrameTable)
        tk.Label(frm_Top, text='Porsche', font=('Porsche Design Font', 14)).pack(pady=15)
        frm_Top.pack(side=tk.TOP)
        frm_Bottom = tk.Frame(FrameTable)
        Show_Progress_Window.progress = ttk.Progressbar(frm_Bottom, orient="horizontal",
                                        length=200, mode="determinate")
        Show_Progress_Window.progress.pack(side=tk.TOP, pady=(30, 10))
        Show_Progress_Window.progress['value'] = 0
        Show_Progress_Window.progress['maximum'] = 50000
        frm_Bottom.pack(side=tk.TOP)
        Show_Progress_Window.root_window.update()
        FrameTable.pack()

class Program_GiL():
    def __init__(self):
        Interface_XML2Excel()
        tk.mainloop()