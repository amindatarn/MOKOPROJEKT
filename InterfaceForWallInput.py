import WallLoops
from tkinter import *
from tkinter import Tk,filedialog,font

window = Tk()
window.geometry("600x400")
window.configure(bg="#FFFFFF")
myFont = font.Font(family="Arial", size=11)

frame1 = Frame(window)
frame2 = Frame(window)

labelFrame = Frame(frame1)
entryFrame = Frame(frame1)
startConfiguration = Frame(frame2)

def setText():
    Projekt_number = entry1.get()
    Ritad_Av = entry2.get()
    Uppdragsansvarig = "PATRIK JENSEN"
    Status = "FOR PRODUCTION"
    Datum = entry5.get()
    Regulations = "001001"
    WallLoops.OuterWalls(Projekt_number, Ritad_Av, Uppdragsansvarig, Status, Datum, Regulations, foldername, exceldokument)
    exit()

def fileDialog():
    global foldername
    foldername = filedialog.askdirectory()
    createButton(startConfiguration, "Start Configuration", setText)

def createLabel(master, labelText):
    label = Label(master=master, text=labelText, anchor='w')
    label['font'] = myFont
    label.configure(bg="#FFFFFF")
    label.pack(fill=X)

def createEntry(master, defaultValue):
    entry = Entry(master=master, width=70, bg="#FFFFFF")
    entry.configure(bg="#FFFFFF")
    entry['font'] = myFont
    entry.pack(fill=X)

    entry.delete(0, END)
    entry.insert(0, defaultValue)

    return entry

def createButton(master, text, command):
    button = Button(master, command=command, text=text, width=70, bg="#FFFFFF")
    button['font'] = myFont
    button.pack()


def browseForFile():
    global exceldokument
    exceldokument = filedialog.askopenfilename()


labelFrame.configure(bg="#FFFFFF")
entryFrame.configure(bg="#FFFFFF")

createLabel(labelFrame, "Project Number:")
entry1 = createEntry(entryFrame, "007")

createLabel(labelFrame, "Ritad Av:")
entry2 = createEntry(entryFrame, "AMINDA TÄRN")

createLabel(labelFrame, "Uppdragsansvarig:")
entry3 = createEntry(entryFrame, "PATRIK JENSEN")

createLabel(labelFrame, "Status:")
entry4 = createEntry(entryFrame, "FOR PRODUCTION")

createLabel(labelFrame, "Datum:")
entry5 = createEntry(entryFrame, "2020-09-11")

createLabel(labelFrame, "Regulations:")
entry6 = createEntry(entryFrame, "001001")

createLabel(labelFrame, "Spara filer som:")
createButton(entryFrame, "Browse", fileDialog)

createLabel(labelFrame, "använd denna excelfil")
createButton(entryFrame, "Browse", browseForFile)

def browseForFile():
    global exceldokument
    exceldokument = filedialog.askopenfilename()

frame1.pack()
frame2.pack()
labelFrame.pack(side=LEFT, fill=BOTH)
entryFrame.pack(side=LEFT, fill=BOTH)
startConfiguration.pack(side=BOTTOM)

window.mainloop()