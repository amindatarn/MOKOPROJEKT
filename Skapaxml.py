import xml.etree.ElementTree as ET

import xlrd
sheet = xlrd.open_workbook(r"C:\Users\AmindaTärn\Desktop\Python\Filmeddata(KOPIA).xlsx").sheet_by_index(0)

def createxmlmall(i,Projekt_number, Ritad_Av, Uppdragsansvarig, Status, Datum, Regulations):
    """creates an xml structure with root and motherelemnts"""

    root = "root"
    root = ET.Element("state")

    model = ET.SubElement(root, "model")
    model.text = (r"C:\Users\Linn Arstad\AppData\Local\Temp\Temporary\Deployed\Files\For\DA\tmpFF28.tmp")

    dataid = ET.SubElement(root, "dataids")

    application = ET.SubElement(root, "application")
    application.text = ("SIBS Configurator")

    safecookie = ET.SubElement(root, "safecookie")

    safestep1 = ET.SubElement(safecookie, "safe-step", name= "Project Information")

    global commits1
    commits1 = ET.SubElement(safestep1, "commits")

    steps = ET.SubElement(root, "steps")
    prev = ET.SubElement(steps, "prev")

    Commit2("Project_number", Projekt_number).addtoxml(commits1)
    Commit2("Ritad_Av", Ritad_Av).addtoxml(commits1)
    Commit2("Uppdragsansvarig", Uppdragsansvarig).addtoxml(commits1)
    Commit2("Status", Status).addtoxml(commits1)
    Commit2("Datum", Datum).addtoxml(commits1)
    Commit2("Regulations", Regulations).addtoxml(commits1)

    lastproxy = ET.SubElement(root, "last-proxy").text="tcserver0"

    filnamn = commit1(i)  # adds subElements to tree in first commit and returns filname

    tree = ET.ElementTree(root)                                     # saves tree in variable "tree"
    return tree, commits1, filnamn, safecookie, steps, prev



def createxmlmalldefault(steps,prev):
    """adds subElements to defaultxml"""
    current = ET.SubElement(steps, "current").text = ("Project Information")
    next = ET.SubElement(steps, "next")
    step = ET.SubElement(next, "step").text = ("Indata")

def createxmlmallejdefault(i,safecookie,steps,prev):
    """adds subElements to the uniqe walls"""
    safestep2 = ET.SubElement(safecookie, "safe-step", name="Indata")
    commits2 = ET.SubElement(safestep2, "commits")

    step = ET.SubElement(prev, "step").text = ("Project Information")
    current = ET.SubElement(steps, "current").text = ("Indata")
    next = ET.SubElement(steps, "next")
    return commits2                                                     # returns commits2 where all the uniqe wall information adds

def addSubWithNameAndText(motherElementText, subElementName, subElementText): #skapar alla attribut för commits1
    ET.SubElement(commits1, motherElementText, name=subElementName).text = subElementText

def addSubWithNameAndTextCommit2(motherElement, motherElementText, subElementName, subElementText): #skapar alla attribut för commits1
    ET.SubElement(motherElement, motherElementText, name=subElementName).text = subElementText
def addSub(motherElement, motherElementText):
    ET.SubElement(motherElement, motherElementText)
def addSubWithName(motherElement, motherElementText, subElementName):
    ET.SubElement(motherElement, motherElementText, name=subElementName)

def addSubWithText(motherElementText, subElementText):
    ET.SubElement(commits1, motherElementText).text = subElementText


def savexml(tree,filnamn,foldername):
    """Creates a folder and saves xml tree in a specific path"""
    import os             # ändrar plats för filer
    mappnamn = "stateexempel.xml"

    try:
        os.chdir(foldername)
        tree.write(filnamn)  # Namnet på ny fil

    except:
        os.mkdir(mappnamn)                                                      # Namnet på ny mapp
        os.chdir(r"C:\Users\AmindaTärn\Desktop\Python\xmlfiler\\" + mappnamn)
        tree.write(filnamn)                                                     # Namnet på ny fil


#EXCEL
def makeinputstring(variabel, i, f):
    """takes input and returns a string"""

    variabel = sheet.cell_value(i,f)
    if type(variabel) == int:
        return str(variabel)
    elif type(variabel) == float:
        return str(int(float(variabel)))
    else:
        return str(variabel)

def formatStiffener(stiffener):
    """gives stiffenes the right format"""
    if stiffener == "Right":
        return "RS"
    elif stiffener == "Left":
        return "LS"
    else:
        return "NS"

def commit1(i): #läser raderna i vald excelfil
    """creates and adds all the information in first commit,different for every module"""
    modulenumber = sheet.cell_value(i, 0)

    addSubWithNameAndText( "committed", "Module_number", modulenumber[2:])                     # adds modulnumer
    addSubWithNameAndText( "committed", "Module_Type", modulenumber[0:2])                      # adds modultypen


    columnsformat =  makeinputstring("columns",i,1) + "P"                                               # format of columns
    addSubWithNameAndText("committed", "Columns", columnsformat)  # adds columns

    addSubWithNameAndText( "committed", "Stiffener", formatStiffener(sheet.cell_value(i, 2)))  # adds stiffener

    addSubWithNameAndText( "committed", "X1", makeinputstring("X1", i, 3))
    addSubWithNameAndText( "committed", "Y1", makeinputstring("Y1", i, 4))
    addSubWithNameAndText("committed", "Z1", makeinputstring("Z1", i, 7))

    if makeinputstring("columns",i,1) == "6":  # lägger in värdet för olika antal dörrar                #different number of columns makes different number of walls
        addSubWithNameAndText("committed", "Y2", makeinputstring("Y2", i, 5))
    elif makeinputstring("columns",i,1) == "8":
        addSubWithNameAndText("committed","Y2", makeinputstring("Y2", i, 5))
        addSubWithNameAndText( "committed","Y3", makeinputstring("Y3", i, 6))

    filnamn = "007" + modulenumber.zfill(3) + "_"

    return filnamn

def ventholecheck(i,commits2): #ger värde till de olika ventholesen
    """Checks if there is any ventholes and adds them"""
    n = 12
    k = 1
    while n <=14:
        v = sheet.cell_value(i, n)
        if isinstance(v, float) == True:
            Commit2(str("X_VH" + str(k)),str(int(v))).addtoxml(commits2)
            Commit2(str("VH" + str(k) + "_qty"),"Yes").addtoxml(commits2)
            k += 1
        n +=1
    else:
        return

def doorsetup(i,l):
    """Checks what doorsetup there is and returns the right format"""
    p = (sheet.cell_value(i, l)).lower()
    if p == "ytterdörr":  # döper de olika dörrarna till rätt benämning
        return "DS1"
    elif p == "inv lgh-dörr":
        return "DS2"
    elif p == "passage":
        return "DS3"
    else:
        print("Fel i doorsetup")
        return None

def trappning(i,l,commits2,trappning):
    """Checks if there is trappning and adds it"""
    p = (sheet.cell_value(i, l)).lower()
    if p == "ja":
        Commit2(trappning, "Yes").addtoxml(commits2)
    elif p == "nej":
        Commit2(trappning, "No").addtoxml(commits2)

class Commit2:
    """puts all the information from each door into two variebles and adds it in second commit"""
    def __init__(self, name,text):
        self.name = name
        self.text = text

    def addtoxml(self,commit):
        addSubWithNameAndTextCommit2(commit, "committed" , self.name ,self.text)

def commit2(i,x, commits2):

    Commit2("Wall_Number", str(x)).addtoxml(commits2)

    ventholecheck(i,commits2)

    trappning(i,10,commits2,"trappningunder")
    trappning(i,11,commits2,"trappningöver")

    wallSetup(i, (sheet.cell_value(i, 9)).lower(), commits2) #adds wallsetup, door/window type,size and setup

def wallSetup(i,s,commits2):
    """"This is where each wallinformation about setup adds"""
    if s == "straight":
        Commit2("Wall_Setup", "S").addtoxml(commits2)

    elif s == "door":
        Commit2("Wall_Setup", "D").addtoxml(commits2)
        makeadoor(i, 15, 16, 17,1,commits2,"Door_Type","X_Door","Door1 Serup")

    elif s == "window":
        Commit2("Wall_Setup", "W").addtoxml(commits2)
        makewindow(i,21,22,23,1,commits2,"Window_Type", "X_Window", "Window_Sill")

    elif s == "window and door":
        Commit2("Wall_Setup", "DRWL").addtoxml(commits2)
        makewindow(i, 21, 22, 23, 1, commits2,"Window_Type", "X_Window", "Window_Sill")
        makeadoor(i, 15, 16, 17, 1, commits2,"Door_Type","X_Door","Door1 Serup")

    elif s == "door and window":
        Commit2("Wall_Setup", "DLWR").addtoxml(commits2)
        makewindow(i, 21, 22, 23, 1, commits2,"Window_Type", "X_Window", "Window_Sill")
        makeadoor(i, 15, 16, 17, 1, commits2,"Door_Type","X_Door","Door1 Serup")

    elif s == "door and door":
        wallsetup = Commit2("Wall_Setup", "DD").addtoxml(commits2)
        makeadoor(i, 15, 16, 17, 1, commits2,"Door_Type","X_Door","Door1 Serup")
        makeadoor(i, 18, 19, 20, 2, commits2,"Door2_Type","X_Door2","Door2 Setup")

    elif s == "window and window":
        wallsetup = Commit2("Wall_Setup", "WW").addtoxml(commits2)
        makewindow(i, 21, 22, 23, 1, commits2,"Window_Type", "X_Window", "Window_Sill")
        makewindow(i, 24, 25, 26, 2, commits2,"Window2_Type", "X_Window_2","Window_Sill_2")

    else:
        wallsetup = "FELLLLLL"

def rounding(x):
    """rounds down the indata of size"""
    if len(x) == 3:
        return int(round(int(x) / 100.0)) * 100
    elif len(x) == 4:
        return int(round(int(x) / 100.0)) * 100

def windowanddoorsize(a):
    """splits up the size and puts it in roundfunction"""
    list = []
    for b in a.split("x"):
        list.append(str(rounding(b)))

    return "x".join(list)

def makeadoor(i,a,b,c,n,commits2, doortype, xdoor, doorserup):
    """"puts values into the variables doors - type,size,setup"""
    avrundadstrl = windowanddoorsize(sheet.cell_value(i, a))
    Commit2(doortype, ("D_" + avrundadstrl)).addtoxml(commits2)
    Commit2(xdoor, str(int(sheet.cell_value(i, b)))).addtoxml(commits2)
    Commit2(doorserup, str(doorsetup(i, c))).addtoxml(commits2)

def makewindow(i,a,b,c,n,commits2,windowtype,xwindow,windowsill):
    """put values into the variebles winfow - size,type and sill"""
    avrundadstrl = windowanddoorsize(sheet.cell_value(i, a))
    Commit2(windowtype, "W_" + avrundadstrl).addtoxml(commits2)
    Commit2(xwindow, str(int(sheet.cell_value(i, b)))).addtoxml(commits2)
    Commit2(windowsill,str(int(sheet.cell_value(i, c)))).addtoxml(commits2)

def gothroughexcel(Projekt_number, Ritad_Av, Uppdragsansvarig, Status, Datum, Regulations,foldername):
    """"main function creating xml from exceldata and then saving it"""
    i=1
    while i < (sheet.nrows):
        if len(sheet.cell_value(i, 0)) == 5 and sheet.cell_value(i, 0)[0:2] == "VA": #and sheet.cell_value(i,9) != "":  # sorterar ut väggar som som har commits1data
            columns = makeinputstring("columns", i, 1)  # kontroll av columns
            tree, commits1, filnamn,safecookie,steps,prev = createxmlmall(i,Projekt_number, Ritad_Av, Uppdragsansvarig, Status, Datum, Regulations)
            createxmlmalldefault(steps, prev)
            savexml(tree, filnamn + "default" + ".xml",foldername)
            x=1
            k=i
            while x<= int(columns):
                if not sheet.cell_value(k, 15)[0:2].lower() == "sa": #sorterar bort alla väggar som är "samma som"
                    tree, commits1, filnamn,safecookie,steps,prev = createxmlmall(i,Projekt_number, Ritad_Av, Uppdragsansvarig, Status, Datum, Regulations)
                    commits2 = createxmlmallejdefault(i, safecookie, steps, prev)
                    commit2(k, x, commits2)
                    savexml(tree, filnamn + "20300" + str(x).zfill(3) + ".tcs",foldername)
                k+=1
                x+=1
        else:
            pass
        i+=1
