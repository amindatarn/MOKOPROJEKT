
import tkinter as tk
from tkinter import filedialog,messagebox,ttk
import WallLoops
import tkinter.ttk as ttk
from ttkthemes import ThemedStyle
from PIL import ImageTk, Image

LARGE_FONT = ("Verdana", 12)


class Window(tk.Tk):


    def __init__(self):
        tk.Tk.__init__(self)

        container = ttk.Frame(self)

        #style = ThemedStyle(container)
        #style.set_theme("black")

        container.pack(side ="top", fill = "both", expand = True)

        container.grid_rowconfigure(0, weight =1)
        container.grid_columnconfigure(0, weight=1)

        self.frames = {}

        for F in (StartPage, PageOne, PageTwo, PageThree, PageFour, PageFive, PageSix, PageSeven, PageEight):

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
        elif conf == "balcony":
            print("balcony")
            WallLoops.Balcony(Uppdragsansvarig, Status, Datum, Regulations, foldername, exceldokument)
        elif conf == "inner":
            print("inner")
            WallLoops.InnerWalls(Uppdragsansvarig, Status, Datum, Regulations, foldername, exceldokument)
        else:
            print("fel")

class StartPage(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self,parent)

        load = Image.open(r"C:\Users\AmindaTärn\Desktop\MOKOPROJEKT\moko_logo_green_rgb.png")
        load = load.resize((200, 50), Image.ANTIALIAS)
        render = ImageTk.PhotoImage(load)
        img = tk.Label(self, image=render)
        img.image = render
        img.pack()

        label1 = tk.Label(self, text = "Modulconfigurator", font = LARGE_FONT)
        label1.pack(pady = 10, padx = 10)

        label2 = tk.Label(self, text="Välj den typ du vill konfigurera:")
        label2.pack(pady=10, padx=10)


        button = ttk.Button(self, text="OuterWall Configurator",
                            command = lambda: controller.show_frame(PageOne))
        button.pack(pady=10, padx=10)

        button2 = ttk.Button(self, text="Facadewall Configurator",
                           command=lambda: controller.show_frame(PageThree))

        button2.pack(pady=10, padx=10)

        button3 = ttk.Button(self, text="Balcony Configurator",
                            command=lambda: controller.show_frame(PageFive))

        button3.pack(pady=10, padx=10)

        button4 = ttk.Button(self, text="InnerWalls Configurator",
                             command=lambda: controller.show_frame(PageSeven))

        button4.pack(pady=10, padx=10)

        label3 = tk.Label(self, text="@ Aminda Tärn")
        label3.pack(pady=10, padx=10, anchor="se")

class PageOne(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)

        label = tk.Label(self, text="Outerwall configurator", font=LARGE_FONT)
        label.grid(column=1, row=0, padx= 10, pady=10)

        label1 = tk.Label(self, text="Välj den excelfil du vill hämta datan från :")
        label1.grid(column=0, row=1, padx= 10, pady=10, sticky = tk.W)

        file1 = tk.StringVar()
        entry1 = ttk.Entry(self, textvariable=file1)
        entry1.grid(column=1, row=1, pady=10,ipadx =140)

        button3 = ttk.Button(self, text="Browse..." ,
                             command = lambda: controller.entry_set_excel(entry1))
        button3.grid(column=2, row=1, padx= 10, pady=10)

        label2 = tk.Label(self, text="Välj vart du vill spara dina filer :")
        label2.grid(column=0, row=2, padx= 10, pady=10, sticky = tk.W)

        file2 = tk.StringVar()
        entry2 = ttk.Entry(self, textvariable=file2)
        entry2.grid(column=1, row=2, pady=10, ipadx = 140)

        button3 = ttk.Button(self, text="Browse...",
                            command=lambda: controller.entry_set_folder(entry2))
        button3.grid(column=2, row=2, padx= 10, pady=10)

        button2 = ttk.Button(self, text="Next",
                             command=lambda: self.checkEntry(controller, file1, file2))

        button2.grid(column=3, row=3, padx=10, pady=10, sticky=tk.E)

        button1 = ttk.Button(self, text="Back to startpage",
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
        entry1 = ttk.Entry(self, textvariable=file1)
        entry1.delete(0, tk.END)
        entry1.insert(0, "2020-09-11")
        entry1.grid(column=2, row=1,padx= 10, pady=10)

        button3 = ttk.Button(self, text="Start Configurator",
                             command=lambda: controller.StartOuterwallConfig(entry1, "outer"))
        button3.grid(column=2, row=2,padx= 10, pady=10, ipadx = 30,ipady = 10)

        button2 = ttk.Button(self, text="Back to outerwallconfigurator",
                             command=lambda: controller.show_frame(PageOne))
        button2.grid(column=1, row=4, padx= 10, sticky = tk.W+ tk.S)

        button1 = ttk.Button(self, text="Back to wall configurator",
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
        entry1 = ttk.Entry(self, textvariable=file1)
        entry1.grid(column=1, row=1, pady=10,ipadx =140)

        button3 = ttk.Button(self, text="Browse..." ,
                             command = lambda: controller.entry_set_excel(entry1))
        button3.grid(column=2, row=1, padx= 10, pady=10)

        label2 = tk.Label(self, text="Välj vart du vill spara dina filer :")
        label2.grid(column=0, row=2, padx= 10, pady=10, sticky = tk.W)

        file2 = tk.StringVar()
        entry2 = ttk.Entry(self, textvariable=file2)
        entry2.grid(column=1, row=2, pady=10, ipadx = 140)

        button3 = ttk.Button(self, text="Browse...",
                            command=lambda: controller.entry_set_folder(entry2))
        button3.grid(column=2, row=2, padx= 10, pady=10)

        button2 = ttk.Button(self, text="Next",
                             command=lambda: self.checkEntry(controller, file1, file2))

        button2.grid(column=3, row=3, padx=10, pady=10, sticky=tk.E)

        button1 = ttk.Button(self, text="Back to startpage",
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
        entry1 = ttk.Entry(self, textvariable=file1)
        entry1.delete(0, tk.END)
        entry1.insert(0, "2020-09-11")
        entry1.grid(column=2, row=1,padx= 10, pady=10)

        button3 = ttk.Button(self, text="Start Configurator",
                             command=lambda: controller.StartOuterwallConfig(entry1, "facade"))
        button3.grid(column=2, row=2,padx= 10, pady=10, ipadx = 30,ipady = 10)

        button2 = ttk.Button(self, text="Back to facadewallconfigurator",
                             command=lambda: controller.show_frame(PageThree))
        button2.grid(column=1, row=4, padx= 10, sticky = tk.W+ tk.S)

        button1 = ttk.Button(self, text="Back to wall configurator",
                             command=lambda: controller.show_frame(StartPage))
        button1.grid(column=0, row=4,padx= 10, sticky = tk.W+ tk.S)

class PageFive(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)

        label = tk.Label(self, text="Balcony configurator", font=LARGE_FONT)
        label.grid(column=1, row=0, padx= 10, pady=10)

        label1 = tk.Label(self, text="Välj den excelfil du vill hämta datan från :")
        label1.grid(column=0, row=1, padx= 10, pady=10, sticky = tk.W)

        file1 = tk.StringVar()
        entry1 = ttk.Entry(self, textvariable=file1)
        entry1.grid(column=1, row=1, pady=10,ipadx =140)

        button3 = ttk.Button(self, text="Browse..." ,
                             command = lambda: controller.entry_set_excel(entry1))
        button3.grid(column=2, row=1, padx= 10, pady=10)

        label2 = tk.Label(self, text="Välj vart du vill spara dina filer :")
        label2.grid(column=0, row=2, padx= 10, pady=10, sticky = tk.W)

        file2 = tk.StringVar()
        entry2 = ttk.Entry(self, textvariable=file2)
        entry2.grid(column=1, row=2, pady=10, ipadx = 140)

        button3 = ttk.Button(self, text="Browse...",
                            command=lambda: controller.entry_set_folder(entry2))
        button3.grid(column=2, row=2, padx= 10, pady=10)

        button2 = ttk.Button(self, text="Next",
                             command=lambda: self.checkEntry(controller, file1, file2))

        button2.grid(column=3, row=3, padx=10, pady=10, sticky=tk.E)

        button1 = ttk.Button(self, text="Back to startpage",
                           command=lambda: controller.show_frame(StartPage))

        button1.grid(column=0, row=3, sticky = tk.W+tk.E, padx= 10, pady=10)

    def checkEntry(self, controller, file1, file2):
        if file1.get() == "" or file2.get() == "":
            print("empty")
            messagebox.showerror("Error", "Du måste fylla i båda fälten" )
        else:
            controller.show_frame(PageSix)

class PageSix(tk.Frame):

    def __init__(self, parent, controller):

        tk.Frame.__init__(self, parent)

        label = tk.Label(self, text=" Balcony configurator", font=LARGE_FONT)
        label.grid(column=2, row=0, padx= 10, pady=10)

        label1 = tk.Label(self, text="Välj datum :")
        label1.grid(column=1, row=1, pady=10, sticky = tk.E)

        file1 = tk.StringVar()
        entry1 = ttk.Entry(self, textvariable=file1)
        entry1.delete(0, tk.END)
        entry1.insert(0, "2020-09-11")
        entry1.grid(column=2, row=1,padx= 10, pady=10)

        button3 = ttk.Button(self, text="Start Configurator",
                             command=lambda: controller.StartOuterwallConfig(entry1, "balcony"))
        button3.grid(column=2, row=2,padx= 10, pady=10, ipadx = 30,ipady = 10)

        button2 = ttk.Button(self, text="Back to facadewallconfigurator",
                             command=lambda: controller.show_frame(PageThree))
        button2.grid(column=1, row=4, padx= 10, sticky = tk.W+ tk.S)

        button1 = ttk.Button(self, text="Back to wall configurator",
                             command=lambda: controller.show_frame(StartPage))
        button1.grid(column=0, row=4,padx= 10, sticky = tk.W+ tk.S)



class PageSeven(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)

        label = tk.Label(self, text="InnerWall configurator", font=LARGE_FONT)
        label.grid(column=1, row=0, padx= 10, pady=10)

        label1 = tk.Label(self, text="Välj den excelfil du vill hämta datan från :")
        label1.grid(column=0, row=1, padx= 10, pady=10, sticky = tk.W)

        file1 = tk.StringVar()
        entry1 = ttk.Entry(self, textvariable=file1)
        entry1.grid(column=1, row=1, pady=10,ipadx =140)

        button3 = ttk.Button(self, text="Browse..." ,
                             command = lambda: controller.entry_set_excel(entry1))
        button3.grid(column=2, row=1, padx= 10, pady=10)

        label2 = tk.Label(self, text="Välj vart du vill spara dina filer :")
        label2.grid(column=0, row=2, padx= 10, pady=10, sticky = tk.W)

        file2 = tk.StringVar()
        entry2 = ttk.Entry(self, textvariable=file2)
        entry2.grid(column=1, row=2, pady=10, ipadx = 140)

        button3 = ttk.Button(self, text="Browse...",
                            command=lambda: controller.entry_set_folder(entry2))
        button3.grid(column=2, row=2, padx= 10, pady=10)

        button2 = ttk.Button(self, text="Next",
                             command=lambda: self.checkEntry(controller, file1, file2))

        button2.grid(column=3, row=3, padx=10, pady=10, sticky=tk.E)

        button1 = ttk.Button(self, text="Back to startpage",
                           command=lambda: controller.show_frame(StartPage))

        button1.grid(column=0, row=3, sticky = tk.W+tk.E, padx= 10, pady=10)

    def checkEntry(self, controller, file1, file2):
        if file1.get() == "" or file2.get() == "":
            print("empty")
            messagebox.showerror("Error", "Du måste fylla i båda fälten" )
        else:
            controller.show_frame(PageEight)

class PageEight(tk.Frame):

    def __init__(self, parent, controller):

        tk.Frame.__init__(self, parent)

        label = tk.Label(self, text=" Balcony configurator", font=LARGE_FONT)
        label.grid(column=2, row=0, padx= 10, pady=10)

        label8 = tk.Label(self, text="Välj datum :")
        label8.grid(column=1, row=1, pady=10, sticky = tk.E)

        file4 = tk.StringVar()
        entry8 = ttk.Entry(self, textvariable=file4)
        entry8.delete(0, tk.END)
        entry8.insert(0, "2020-09-11")
        entry8.grid(column=2, row=1,padx= 10, pady=10)

        button3 = ttk.Button(self, text="Start Configurator",
                             command=lambda: controller.StartOuterwallConfig(entry8, "inner"))
        button3.grid(column=2, row=2,padx= 10, pady=10, ipadx = 30,ipady = 10)

        button2 = ttk.Button(self, text="Back to facadewallconfigurator",
                             command=lambda: controller.show_frame(PageThree))
        button2.grid(column=1, row=4, padx= 10, sticky = tk.W+ tk.S)

        button1 = ttk.Button(self, text="Back to wall configurator",
                             command=lambda: controller.show_frame(StartPage))
        button1.grid(column=0, row=4,padx= 10, sticky = tk.W+ tk.S)



app = Window()
app.mainloop()