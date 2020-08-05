
import tkinter as tk
from tkinter import filedialog, font
from tkinter import ttk
from tkinter import messagebox
import WallLoops

LARGE_FONT = ("Verdana", 12)

class Window(tk.Tk):

    def __init__(self):
        tk.Tk.__init__(self)
        container = tk.Frame(self)

        container.pack(side ="top", fill = "both", expand = True)

        container.grid_rowconfigure(0, weight =1)
        container.grid_columnconfigure(0, weight=1)

        self.frames = {}

        for F in (StartPage, PageOne, PageTwo, PageThree, PageFour):

            frame = F(container,self)

            self.frames[F] = frame

            frame.grid(row=0, column = 0, sticky = "nsew")

        self.show_frame(StartPage)

    def show_frame(self, cont):
        frame = self.frames[cont]
        frame.tkraise()

    def entry_set_excel(self, entry):
        global exceldokument
        exceldokument = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])

        entry.delete(0, 'end')
        entry.insert(tk.END, exceldokument)

    def entry_set_folder(self, entry):
        global foldername
        foldername = filedialog.askdirectory()

        entry.delete(0, 'end')
        entry.insert(tk.END, foldername)


    def StartOuterwallConfig(self, entry, conf):
        Datum = entry.get()
        Uppdragsansvarig = "PATRIK JENSEN"
        Status = "FOR PRODUCTION"
        Regulations = "001001"
        if conf == "outer":
            print("outer")
            WallLoops.OuterWalls(Uppdragsansvarig, Status, Datum, Regulations, foldername, exceldokument)
        elif conf == "facade":
            print("facade")
            WallLoops.FasadWalls(Uppdragsansvarig, Status, Datum, Regulations, foldername, exceldokument)
        else:
            print("fel")

class StartPage(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self,parent)


        label = tk.Label(self, text = "Wallconfigurator", font = LARGE_FONT)
        label.pack(pady = 10, padx = 10)

        label = tk.Label(self, text="Välj den typ av vägg du vill konfigurera:")
        label.pack(pady=10, padx=10)

        img = tk.PhotoImage(file = r"C:\Users\AmindaTärn\Desktop\Python\Python program\Namnlös.png")
        button = tk.Button(self, image = img,  width="100", height="25",
                            command = lambda: controller.show_frame(PageOne))

        button.image =img
        button.pack(pady=10, padx=10)

        button2 = tk.Button(self, text="Facadewall Configurator",
                           command=lambda: controller.show_frame(PageThree))

        button2.pack(pady=10, padx=10)

class PageOne(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)

        label = tk.Label(self, text="Outerwall configurator", font=LARGE_FONT)
        label.grid(column=1, row=0, padx= 10, pady=10)

        label1 = tk.Label(self, text="Välj den excelfil du vill hämta datan från :")
        label1.grid(column=0, row=1, padx= 10, pady=10, sticky = tk.W)

        file1 = tk.StringVar()
        entry1 = tk.Entry(self, textvariable=file1)
        entry1.grid(column=1, row=1, pady=10,ipadx =140)

        button3 = tk.Button(self, text="Browse..." ,
                             command = lambda: controller.entry_set_excel(entry1))
        button3.grid(column=2, row=1, padx= 10, pady=10)

        label2 = tk.Label(self, text="Välj vart du vill spara dina filer :")
        label2.grid(column=0, row=2, padx= 10, pady=10, sticky = tk.W)

        file2 = tk.StringVar()
        entry2 = tk.Entry(self, textvariable=file2)
        entry2.grid(column=1, row=2, pady=10, ipadx = 140)

        button3 = tk.Button(self, text="Browse...",
                            command=lambda: controller.entry_set_folder(entry2))
        button3.grid(column=2, row=2, padx= 10, pady=10)

        button2 = tk.Button(self, text="Next",
                             command=lambda: self.checkEntry(controller, file1, file2))

        button2.grid(column=3, row=3, padx=10, pady=10, sticky=tk.E)

        button1 = tk.Button(self, text="Back to startpage",
                           command=lambda: controller.show_frame(StartPage))
        button1.grid(column=0, row=3, sticky = tk.W+tk.E, padx= 10, pady=10)

    def checkEntry(self, controller, file1, file2):
        if file1.get() == "" or file2.get() == "":
            print("empty")
            messagebox.showerror("Error", "Du måste fylla i båda fälten" )
        else:
            controller.show_frame(PageTwo)

class PageTwo(tk.Frame):

    def __init__(self, parent, controller):

        tk.Frame.__init__(self, parent)

        label = tk.Label(self, text="Outerwall configurator", font=LARGE_FONT)
        label.grid(column=2, row=0, padx= 10, pady=10)

        label1 = tk.Label(self, text="Välj datum :")
        label1.grid(column=1, row=1, pady=10, sticky = tk.E)

        file1 = tk.StringVar()
        entry1 = tk.Entry(self, textvariable=file1)
        entry1.delete(0, tk.END)
        entry1.insert(0, "2020-09-11")
        entry1.grid(column=2, row=1,padx= 10, pady=10)

        button3 = tk.Button(self, text="Start Configurator",
                             command=lambda: controller.StartOuterwallConfig(entry1, "outer"))
        button3.grid(column=2, row=2,padx= 10, pady=10, ipadx = 30,ipady = 10)

        button2 = tk.Button(self, text="Back to outerwallconfigurator",
                             command=lambda: controller.show_frame(PageOne))
        button2.grid(column=1, row=4, padx= 10, sticky = tk.W+ tk.S)

        button1 = tk.Button(self, text="Back to wall configurator",
                             command=lambda: controller.show_frame(StartPage))
        button1.grid(column=0, row=4,padx= 10, sticky = tk.W+ tk.S)

class PageThree(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)

        label = tk.Label(self, text="Facadewall configurator", font=LARGE_FONT)
        label.grid(column=1, row=0, padx= 10, pady=10)

        label1 = tk.Label(self, text="Välj den excelfil du vill hämta datan från :")
        label1.grid(column=0, row=1, padx= 10, pady=10, sticky = tk.W)

        file1 = tk.StringVar()
        entry1 = tk.Entry(self, textvariable=file1)
        entry1.grid(column=1, row=1, pady=10,ipadx =140)

        button3 = tk.Button(self, text="Browse..." ,
                             command = lambda: controller.entry_set_excel(entry1))
        button3.grid(column=2, row=1, padx= 10, pady=10)

        label2 = tk.Label(self, text="Välj vart du vill spara dina filer :")
        label2.grid(column=0, row=2, padx= 10, pady=10, sticky = tk.W)

        file2 = tk.StringVar()
        entry2 = tk.Entry(self, textvariable=file2)
        entry2.grid(column=1, row=2, pady=10, ipadx = 140)

        button3 = tk.Button(self, text="Browse...",
                            command=lambda: controller.entry_set_folder(entry2))
        button3.grid(column=2, row=2, padx= 10, pady=10)

        button2 = tk.Button(self, text="Next",
                             command=lambda: self.checkEntry(controller, file1, file2))

        button2.grid(column=3, row=3, padx=10, pady=10, sticky=tk.E)

        button1 = tk.Button(self, text="Back to startpage",
                           command=lambda: controller.show_frame(StartPage))

        button1.grid(column=0, row=3, sticky = tk.W+tk.E, padx= 10, pady=10)

    def checkEntry(self, controller, file1, file2):
        if file1.get() == "" or file2.get() == "":
            print("empty")
            messagebox.showerror("Error", "Du måste fylla i båda fälten" )
        else:
            controller.show_frame(PageFour)

class PageFour(tk.Frame):

    def __init__(self, parent, controller):

        tk.Frame.__init__(self, parent)

        label = tk.Label(self, text="Facadewall configurator", font=LARGE_FONT)
        label.grid(column=2, row=0, padx= 10, pady=10)

        label1 = tk.Label(self, text="Välj datum :")
        label1.grid(column=1, row=1, pady=10, sticky = tk.E)

        file1 = tk.StringVar()
        entry1 = tk.Entry(self, textvariable=file1)
        entry1.delete(0, tk.END)
        entry1.insert(0, "2020-09-11")
        entry1.grid(column=2, row=1,padx= 10, pady=10)

        button3 = tk.Button(self, text="Start Configurator",
                             command=lambda: controller.StartOuterwallConfig(entry1, "facade"))
        button3.grid(column=2, row=2,padx= 10, pady=10, ipadx = 30,ipady = 10)

        button2 = tk.Button(self, text="Back to facadewallconfigurator",
                             command=lambda: controller.show_frame(PageThree))
        button2.grid(column=1, row=4, padx= 10, sticky = tk.W+ tk.S)

        button1 = tk.Button(self, text="Back to wall configurator",
                             command=lambda: controller.show_frame(StartPage))
        button1.grid(column=0, row=4,padx= 10, sticky = tk.W+ tk.S)

app = Window()
app.mainloop()