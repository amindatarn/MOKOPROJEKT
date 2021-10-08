import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import ImageTk, Image
import getpass
import xml.etree.ElementTree as ET
import xlrd
from pathlib import Path


def resource_path(relative_path):
    """ Get absolute path to resource, to get logo"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def get_user_name():

    anv = getpass.getuser()
    anv = ("".join([(" "+i if i.isupper() else i) for i in anv]).strip().upper())
    return anv

def loop_trough_row(var, name_col, list_error, list_of_project_info):
    """"function looping through other functions to build tree and then save it"""

    sheet = get_excel(exceldokument)

    i = 3
    while i < sheet.nrows:
        try:
            file_name = str(sheet.cell_value(i, name_col))
            if file_name != "":
                tree, safecookie, steps, prev = createxmlmall()
                list_error = loop_through_col(steps, safecookie, i, file_name, var, list_error, list_of_project_info)

                # save_xml(tree, (file_name) +  ".xml", folder_name)
                for errors in list_error:
                    if errors.error_type == "4":
                        return list_error
                save_xml(tree, file_name + ".tcs", folder_name)

            else:
                for l in range(sheet.ncols):
                    if str(sheet.cell_value(i, l)) != "":
                        p = AddFileWithError(i + 1, "3")
                        list_error = p.add_el(list_error, p)
            i += 1

        except:
            p = AddFileWithError(i + 1, "1")
            list_error = p.add_el(list_error, p)
            return list_error

    return list_error


def get_excel(exceldocument):
    """opens the chosen exceldocument and returns it as sheet"""

    sheet = xlrd.open_workbook(exceldocument).sheet_by_index(0)
    return sheet


def createxmlmall():
    """creates an xml structure with root and motherelements"""

    root = ET.Element("state")
    model = ET.SubElement(root, "model")
    model.text = r""

    dataid = ET.SubElement(root, "dataids")
    application = ET.SubElement(root, "application")

    application.text = "SIBS Configurator"
    safecookie = ET.SubElement(root, "safecookie")
    steps = ET.SubElement(root, "steps")
    prev = ET.SubElement(steps, "prev")

    lastproxy = ET.SubElement(root, "last-proxy").text = "tcserver0"

    tree = ET.ElementTree(root)                                     # saves tree in variable "tree"
    return tree, safecookie, steps, prev


class Steps:
    """adds each column not empty in excelsheet to tree"""
    def __init__(self, name):
        self.name = name

    def addtoxml(self, i, n, steps, safecookie):
        safestep2 = ET.SubElement(safecookie, "safe-step", name=self.name)
        commit = ET.SubElement(safestep2, "commits")

        if i == str(1):
            prev = ET.SubElement(steps, "prev")
            step = ET.SubElement(prev, "step").text = self.name

        elif i == str(n):
            next = ET.SubElement(steps, "next")
            step = ET.SubElement(next, "step").text = self.name

        else:
            current = ET.SubElement(steps, "current").text = self.name
            # step = ET.SubElement(current, "step").text = (self.name)

        return commit


def add_sub(mother_element, mother_element_text, sub_element_name, sub_element_text):           # skapar alla attribut för commits1
    ET.SubElement(mother_element, mother_element_text, name=sub_element_name).text = sub_element_text


def save_xml(tree, file_name, folder_name):
    """Creates a folder and saves xml tree in a specific path"""
    import os             # ändrar plats för filer
    os.chdir(folder_name)
    tree.write(file_name)  # Namnet på ny fil


def makeinputstring(variabel):
    """takes input and returns a string"""
    if type(variabel) == int:
        return str(variabel)
    elif type(variabel) == float:
        return str(int(float(variabel)))
    else:
        return str(variabel)


class AddAttrToTree:

    """Creates a commit for each variable in excelsheet and adds it to the tree. Each indata from excel are named "text" and added to its own element and name"""
    def __init__(self, element, name, text):
        self.element = element
        self.name = name
        self.text = text

    def addtoxml(self):
        add_sub(self.element, "committed", self.name, self.text)


def loop_through_col(steps, safecookie, b, file_name, var, list_error, list_of_project_info):
    """This function loops through the excel and sorts out elements, names, texts and where in the tree they should be added."""

    col, k, j, g = 4, 1, 0, 0
    sheet = get_excel(exceldokument)
    row_for_commitname = 2      #nya
    #row_for_commitname = 1     #gamla


    while col < sheet.ncols:
        if sheet.cell_type(0, col) != 0:
            j += 1
        if sheet.cell_type(row_for_commitname, col)!= 0:
            g += 1
        col += 1

    if j == 0 or g == 0:
        p = AddFileWithError(file_name, "4")
        list_error = p.add_el(list_error, p)
        return list_error

    col = 4

    while col < sheet.ncols:


        if sheet.cell_type(0, col) != 0:


            name = (sheet.cell_value(0, col))
            commit = Steps(name).addtoxml(k, j, steps, safecookie)
            commit_name = (sheet.cell_value(row_for_commitname, col))
            list_error = check_cell_error(b, col, sheet, list_error, file_name)
            if commit_name.lower() == "littra":
                commit_name = "littra"
            if k == 1 and sheet.cell_type(b, col) != 0:
                c_name = (sheet.cell_value(b, col))
                AddAttrToTree(commit, commit_name, makeinputstring(c_name)).addtoxml()
                for x in range(0, 5):
                    AddAttrToTree(commit, list_of_project_info[x], makeinputstring(var[x])).addtoxml()

            elif k != 1 and sheet.cell_type(b, col) != 0:
                c_name = (sheet.cell_value(b, col))
                AddAttrToTree(commit, commit_name, makeinputstring(c_name)).addtoxml()

            k += 1
            n = col + 1

            if sheet.cell_type(0, n) != 0:
                if sheet.cell_type(b, col) != 0:
                    commit_name = (sheet.cell_value(row_for_commitname, col))
                    c_name = (sheet.cell_value(b, col))
                    AddAttrToTree(commit, commit_name, makeinputstring(c_name)).addtoxml()

            elif sheet.cell_type(0, n) == 0:
                try:
                    while sheet.cell_type(0, n) == 0 and n - 1 < sheet.ncols:
                        if sheet.cell_type(b, n) != 0:
                            list_error = check_cell_error(b, n, sheet, list_error, file_name)
                            commit_name = (sheet.cell_value(row_for_commitname, n))
                            c_name = (sheet.cell_value(b, n))
                            AddAttrToTree(commit, commit_name, makeinputstring(c_name)).addtoxml()

                        n += 1
                except:
                    pass


        col += 1





    return list_error

class AddFileWithError:
    def __init__(self, file_name, error_type):
        self.file_name = file_name
        self.error_type = error_type

    def add_el(self, list_error,p):
        list_error.append(p)
        return list_error


def check_cell_error(b, n, sheet, list_error, file_name):
    if sheet.cell(b, n).ctype == 5:  # check excel cell for error
        p = AddFileWithError(file_name, "2")
        list_error = p.add_el(list_error, p)
    return list_error


class Window(tk.Tk):
    """This is the main window"""
    def __init__(self):
        tk.Tk.__init__(self)
        colour = "white"
        container = ttk.Frame(self)
        container.pack(side="top", fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self.frames = {}
        self.title("Tacton Configurator States Generator")

        try:
            load = Image.open(resource_path('moko_logo_green_rgb.png'))
            load = load.resize((2, 2), Image.ANTIALIAS)
            self.iconphoto(False, ImageTk.PhotoImage(load))
        except:
           pass

        frame = StartPage(container, self)
        frame.configure(background=colour)
        self.frames[StartPage] = frame
        frame.grid(row=0, column=0, sticky="nsew")

    def entry_set_excel(self, entry):
        """Creates variable of chosen exceldocument"""
        global exceldokument
        exceldokument = filedialog.askopenfilename(filetypes=[("Excel file","*.xlsx"),("Excel file", "*.xlsm")])
        entry.delete(0, 'end')
        entry.insert(tk.END, exceldokument)

    def entry_set_folder(self, entry):
        """Creates variable of chosen folder"""
        global folder_name
        folder_name = filedialog.askdirectory()
        entry.delete(0, 'end')
        entry.insert(tk.END, folder_name)

    def start_config(self, entries, name_col, list_error, list_of_project_info):
        """Sends all the inputinformation to loop"""
        var = []
        for x in entries[3:]:
            var.append(x.get())
        list_error = loop_trough_row(var, name_col, list_error, list_of_project_info)
        return list_error


class StartPage(tk.Frame):
    """This is the main page that is in the mainwindow."""
    def __init__(self, parent, controller):
        small_font = ("Verdana", 11)
        large_font = ("Verdana", 14)

        tk.Frame.__init__(self, parent)
        colour = "white"



        dir_path = Path('C:\Temp')
        file_name = 'indata.txt'

        # check if directory exists

        try:
            f = open("indata.txt", "r")
            file_lines = f.readlines()
            f.close()
        except:
            if my_path.is_dir():
                f = open(dir_path.joinpath(file_name), 'w')
                print('File created')
                for i in ("B", "2020-02-20", "AMINDA TÄRN", "PATRIK JENSEN", "FOR PRODUCTION", "001001"):
                    f.write(i + "\n")
                f.close()
            else:
                print('Directory doesn texist')
            #f = open("indata.txt", "w")


        f = open("indata.txt", "r")
        file_lines = f.readlines()

        try:
            line5 = file_lines[2].strip()  # First Line
        except:
            anv = get_user_name()
            line5 = anv
        f.close()

        frame1 = tk.Frame(self)
        frame1.configure(background=colour)
        frame1.grid(row=0, sticky="nsew")

        frame1.columnconfigure(0, weight=1)
        frame1.columnconfigure(1, weight=2)
        frame1.columnconfigure(3, weight=1)

        frame1.rowconfigure(0, weight=1)
        frame1.rowconfigure(1, weight=1)
        frame1.rowconfigure(15, weight=3)

        self.columnconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        self.rowconfigure(13, weight=1)
        self.rowconfigure(14, weight=1)

        def make_label(self, text, col, row):
            label = tk.Label(self, text=text, font=small_font)
            label.configure(background=colour)
            label.grid(column=col, row=row, ipadx=100, pady=10, sticky='w')

        def make_entry(self, default, row, col):
            entry = ttk.Entry(self)
            entry.insert(tk.END, default)
            entry.grid(column=col, row=row, pady=10, sticky='nsew')
            return entry

        try:
            load = Image.open(resource_path('moko_logo_green_rgb.png'))
            load = load.resize((200, 50), Image.ANTIALIAS)

            render = ImageTk.PhotoImage(load)
            img = tk.Label(frame1, image=render)
            img.image = render
            img.configure(background=colour)
            img.grid(column=0, row=0, ipadx=100, pady=10, sticky='w')
        except:
            pass

        label0 = tk.Label(frame1, text="Tacton Configurator States Generator", font=large_font)
        label0.configure(background=colour)
        label0.grid(column=1, row=0, ipadx=100, pady=10, sticky='nsew')

        make_label(frame1, "Exceldocument :", "0", "2")

        file1 = tk.StringVar()
        entry1 = ttk.Entry(frame1, textvariable=file1)
        entry1.grid(column=1, row=2, pady=10, sticky='nsew')

        button1 = ttk.Button(frame1, text="Browse...", command=lambda: controller.entry_set_excel(entry1))
        button1.grid(column=2, row=2, padx=10, pady=10, sticky='nsew')

        make_label(frame1, "Folder :", "0", "3")

        file2 = tk.StringVar()
        entry2 = ttk.Entry(frame1, textvariable=file2)
        entry2.grid(column=1, row=3, pady=10, sticky='nsew')

        button2 = ttk.Button(frame1, text="Browse...",
                            command=lambda: controller.entry_set_folder(entry2))
        button2.grid(column=2, row=3, padx=10, pady=10, sticky='nsew')

        make_label(frame1, "Filename from column (STATE FILE) :", "0", "4")
        entry21 = ttk.Entry(frame1)
        entry21.insert(tk.END, file_lines[0].strip())
        entry21.grid(column="1", row="4", pady=10, sticky='nsew')

        label00 = tk.Label(frame1, text="", font=small_font)
        label00.configure(background=colour)
        label00.grid(column=1, row=5, padx=10, pady=10)

        frame2 = tk.Frame(self)
        frame2.configure(background=colour)
        frame2.grid(row=6, sticky="nsew")

        frame2.columnconfigure(2, weight=2)
        frame2.columnconfigure(1, weight=1)

        frame2.rowconfigure(0, weight=1)
        frame2.rowconfigure(1, weight=2)

        list_of_p_info = ["Date", "Drafter", "Project_manager", "Status", "Regulations"]    #nya
        #list_of_p_info = ["Datum", "Ritad_Av", "Uppdragsansvarig", "Status", "Regulations"]    #gamla
        x = 6
        list_of_project_info = []
        for i in list_of_p_info:
            make_label(frame1, i, "0", x)
            i = (i.lower())
            list_of_project_info.append(i)
            x += 1

        entry3 = make_entry(frame1, file_lines[1].strip(), "6", "1")        # 2020-02-20
        entry4 = make_entry(frame1, line5.upper(), "7", "1")                        # Aminda Tärn
        entry5 = make_entry(frame1, file_lines[3].strip().upper(), "8", "1")        # PATRIK jENSEN
        entry6 = make_entry(frame1, file_lines[4].strip().upper(), "9", "1")        # For production
        entry7 = make_entry(frame1, file_lines[5].strip(), "10", "1")       # 001001

        frame3 = tk.Frame(self)
        frame3.configure(background=colour)
        frame3.grid(row=12, sticky="nsew")

        label4 = tk.Label(self, text="")
        label4.configure(background=colour)
        label4.grid(row=13)

        entries = [entry1, entry2, entry21, entry3, entry4, entry5,
                   entry6, entry7]

        error_label = tk.Label(self, text=(""))
        error_label.configure(background=colour)
        error_label.grid(row=14)

        button3 = ttk.Button(self, text="Generate State Files", command=lambda: self.check_entry(controller, entries, list_of_project_info, error_label))
        button3.grid(row=15, padx=10, pady=10, ipadx=30, ipady=10)

        label5 = tk.Label(self, text="@ Aminda Tärn")
        label5.configure(background=colour)
        label5.grid(row=16, pady=10, padx=10, sticky='se')


    def col_to_num(self, entry):
        import string
        try:
            num = 0
            for c in entry:
                if c in string.ascii_letters:
                    name_col = num * 26 + (ord(c.upper()) - ord('A'))

            return name_col
        except:
            return "Error"

    def write_to_indata(self, entries):
        try:
            f = open("indata.txt", "w")
            text_list = []
            for x in range(2, len(entries)):
                text_list.append(str(entries[x].get()) + "\n")
            f.writelines(text_list)
            f.close()
        except:
            pass

    def check_entry(self, controller, entries, list_of_project_info, error_label):
        """Contoll of inputs and try/except for mainloop"""

        for x in range(0, len(entries)):
            if entries[x].get() == "":
                messagebox.showerror("Error", "Expected no empty fields")
                return
            if not entries[2].get().isalpha():
                messagebox.showerror("Error", "Expected column in letter not number, e.g. 'B' ")
                return
        name_col = self.col_to_num(entries[2].get())
        self.write_to_indata(entries)

        list_error,error_present  = [], []
        list_error = controller.start_config(entries, name_col, list_error, list_of_project_info)
        if len(list_error) == 0:
            message = "Successfully generated all state files"
            error_present.append(message)
            error_label.config(text="Successfully generated all state files")
        else:
            for element in list_error:
                if element.error_type == "1":                                                               # error in loop_trough_row
                    message = "expected error in excel spreadsheet at row" + str(element.file_name) + "\n"
                elif element.error_type == "2":                                                             #filname missing
                    message = "expected error in file " + str(element.file_name)+ "\n"
                elif element.error_type == "3":                                                             # Filename error
                    message = "expected error in file name at row " + str(element.file_name) + "\n"
                elif element.error_type == "4":                                                             # "Seems like error in 1:st or 3:rd line in excel sheet"
                    message = "expected error in excel spreadsheet on 1:st or 3:rd row " + "\n"
                error_present.append(message)
            error_report = open("error_report.txt", "w+")
            error_report.write(''.join(error_present))
            error_report.close()
            error_label.config(text="Error occured, check error report in "+ entries[1].get())
        # error_label.config(text=(''.join(error_present)))



app = Window()
app.mainloop()
