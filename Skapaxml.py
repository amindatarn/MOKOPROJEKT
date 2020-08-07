import xml.etree.ElementTree as ET
import xlrd

def getExcel(exceldokument):
    global sheet
    sheet = xlrd.open_workbook(exceldokument).sheet_by_index(0)

    return sheet
    # sheet = xlrd.open_workbook(r"C:\Users\AmindaTärn\Desktop\Python\Filmeddata(KOPIA) .xlsx").sheet_by_index(0)
    # C:/Users/AmindaTärn/Desktop/Python/Filmeddata(KOPIA) .xlsx


def createxmlmall(Ritad_Av, Uppdragsansvarig, Status, Datum, Regulations,projectNumber):
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

    addSubWithNameAndText("committed", "Project_number", projectNumber)

    Commit2("Ritad_Av", Ritad_Av).addtoxml(commits1)
    Commit2("Uppdragsansvarig", Uppdragsansvarig).addtoxml(commits1)
    Commit2("Status", Status).addtoxml(commits1)
    Commit2("Datum", Datum).addtoxml(commits1)
    Commit2("Regulations", Regulations).addtoxml(commits1)

    lastproxy = ET.SubElement(root, "last-proxy").text="tcserver0"

    tree = ET.ElementTree(root)                                     # saves tree in variable "tree"
    return tree, commits1, safecookie, steps, prev

def createxmlmalldefault(steps):
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

def savexml(tree,filnamn,foldername):
    """Creates a folder and saves xml tree in a specific path"""
    import os             # ändrar plats för filer

    os.chdir(foldername)
    tree.write(filnamn)  # Namnet på ny fil

    #except:
       # os.mkdir(r"C:\Users\AmindaTärn\Desktop\Python\xmlfiler\\" + foldername)                                                      # Namnet på ny mapp
       ## os.chdir(r"C:\Users\AmindaTärn\Desktop\Python\xmlfiler\\" + foldername)
       # tree.write(filnamn)                                                     # Namnet på ny fil

#EXCEL
def makeinputstring(i, f):
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
def innerWalls(i,commits2):
    modulenumber = sheet.cell_value(i, 0)

    addSubWithNameAndText("committed", "Module_number", modulenumber[5:])  # adds modulnumer
    addSubWithNameAndText("committed", "Module_Type", modulenumber[3:5])  # adds modultypen

    Commit2("Wall_Number",sheet.cell_value(i, 1)[4:]).addtoxml(commits2)
    Commit2("Length", makeinputstring(i, 4)).addtoxml(commits2)

    Commit2("Fire_Gypsum", sheet.cell_value(i, 1)).addtoxml(commits2)

    s = sheet.cell_value(i, 2).lower()

    if s =="straight":
        Commit2("Wall_Setup", "S").addtoxml(commits2)

    elif s == "door":
        makedoorInner(i, 5, 6, commits2, "Door_Type","X_Door")

    elif s == "door and door":
        makedoorInner(i, 5, 6, commits2, "Door_Type", "X_Door")
        makedoorInner(i, 7, 8, commits2, "Door2_Type", "X_Door2")

    filnamn = modulenumber.zfill(3) + "_"

    return filnamn


def outerWallsCommit1(i): #läser raderna i vald excelfil
    """creates and adds all the information in first commit,different for every module"""
    modulenumber = sheet.cell_value(i, 0)

    addSubWithNameAndText( "committed", "Module_number", modulenumber[2:])                     # adds modulnumer
    addSubWithNameAndText( "committed", "Module_Type", modulenumber[0:2])                      # adds modultypen

    columnsformat =  makeinputstring(i,1) + "P"                                               # format of columns
    addSubWithNameAndText("committed", "Columns", columnsformat)  # adds columns

    addSubWithNameAndText( "committed", "Stiffener", formatStiffener(sheet.cell_value(i, 2)))  # adds stiffener

    addSubWithNameAndText( "committed", "X1", makeinputstring(i, 3))
    addSubWithNameAndText( "committed", "Y1", makeinputstring( i, 4))
    addSubWithNameAndText("committed", "Z1", makeinputstring(i, 7))

    if makeinputstring(i,1) == "6":  # lägger in värdet för olika antal dörrar                #different number of columns makes different number of walls
        addSubWithNameAndText("committed", "Y2", makeinputstring( i, 5))
    elif makeinputstring(i,1) == "8":
        addSubWithNameAndText("committed","Y2", makeinputstring( i, 5))
        addSubWithNameAndText( "committed","Y3", makeinputstring( i, 6))

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

def outerWallsCommit2(i,x, commits2):

    Commit2("Wall_Number", str(x)).addtoxml(commits2)

    ventholecheck(i,commits2)

    trappning(i,10,commits2,"trappningunder")
    trappning(i,11,commits2,"trappningöver")

    wallSetup(i, (sheet.cell_value(i, 9)).lower(), commits2) #adds wallsetup, door/window type,size and setup

def facadeWallsCommits(i, commits2):

    filnamn = (sheet.cell_value(i, 1)[0:8])

    modulenumber = str(sheet.cell_value(i, 1))

    addSubWithNameAndText("committed", "Module_Type", modulenumber[3:5])
    addSubWithNameAndText("committed", "Module_number", modulenumber[5:8])
    Commit2("Length", makeinputstring( i, 2)).addtoxml(commits2)
    Commit2("Offset",makeinputstring( i, 3)).addtoxml(commits2)
    Commit2("Floor", checkFloor(i)).addtoxml(commits2)

    #Commit2("Facade Material", "Stående Träpanel").addtoxml(commits2)

    wallSetupFacade(i, (sheet.cell_value(i, 5)).lower(), commits2)

    return filnamn

def BalconyCommits(i, commits2):

    filnamn = (sheet.cell_value(i, 1)[0:8])

    modulenumber = str(sheet.cell_value(i, 1))

    addSubWithNameAndText("committed", "Module_Type", modulenumber[3:5])
    addSubWithNameAndText("committed", "Module_number", modulenumber[5:8])


    Commit2("Length", makeinputstring( i, 2)).addtoxml(commits2)
    Commit2("Depth",makeinputstring( i, 3)).addtoxml(commits2)
    Commit2("Height", "2900").addtoxml(commits2)
    Commit2("Tension_rod", checkTension(i)).addtoxml(commits2)
    Commit2("CornerAB", checkCornerAB(i)).addtoxml(commits2)
    Commit2("CornerBC", checkCornerBC(i)).addtoxml(commits2)
    Commit2("CornerCD", checkCornerCD(i)).addtoxml(commits2)
    Commit2("CornerAD", checkCornerDA(i)).addtoxml(commits2)

    SideInput(i, "A", commits2)
    SideInput(i, "B", commits2)
    SideInput(i, "C", commits2)
    SideInput(i, "D", commits2)
    rail(i, commits2)

    return filnamn

def rail(i,commits2):
    if (sheet.cell_value(i, 30)).lower() == "a":
        Commit2("Rail_SideA_Length", makeinputstring(i, 32)).addtoxml(commits2)
        Commit2("Offset_SideA", makeinputstring(i, 31)).addtoxml(commits2)
        Commit2("Rail_SideB", "No").addtoxml(commits2)
        Commit2("Rail_SideC", "No").addtoxml(commits2)
        Commit2("Rail_SideD", "No").addtoxml(commits2)

    elif (sheet.cell_value(i, 30)).lower() == "b":
        Commit2("Rail_SideB_Length", makeinputstring(i, 32)).addtoxml(commits2)
        Commit2("Offset_SideB", makeinputstring(i, 31)).addtoxml(commits2)
        Commit2("Rail_SideA", "No").addtoxml(commits2)
        Commit2("Rail_SideC", "No").addtoxml(commits2)
        Commit2("Rail_SideD", "No").addtoxml(commits2)

    elif (sheet.cell_value(i, 30)).lower() == "c":
        Commit2("Rail_SideC_Length", makeinputstringe(i, 32)).addtoxml(commits2)
        Commit2("Offset_SideC", makeinputstring(i, 31)).addtoxml(commits2)
        Commit2("Rail_SideB", "No").addtoxml(commits2)
        Commit2("Rail_SideA", "No").addtoxml(commits2)
        Commit2("Rail_SideD", "No").addtoxml(commits2)

    elif (sheet.cell_value(i, 30)).lower() == "d":
        Commit2("Rail_SideD_Length", sheet.cell_value(i, 32)).addtoxml(commits2)
        Commit2("Offset_SideD", sheet.cell_value(i, 31)).addtoxml(commits2)
        Commit2("Rail_SideB", "No").addtoxml(commits2)
        Commit2("Rail_SideC", "No").addtoxml(commits2)
        Commit2("Rail_SideA", "No").addtoxml(commits2)

    else:
        Commit2("Rail_SideA", "No").addtoxml(commits2)
        Commit2("Rail_SideB", "No").addtoxml(commits2)
        Commit2("Rail_SideC", "No").addtoxml(commits2)
        Commit2("Rail_SideD", "No").addtoxml(commits2)

def SideInput(i,side,commits2):
    if side == "A":
        n = 9
        k = 10
        j = 11
    elif side == "B":
        n = 15
        k = 16
        j = 17

    elif side == "C":
        n = 21
        k = 22
        j = 23

    elif side == "D":
        n = 27
        k = 28
        j = 29

    Commit2("Water Proofing Plates Side " + str(side), WaterProof(i,n)).addtoxml(commits2)
    Commit2("Facade_plate_Side_" + str(side) + "_Length" , makeinputstring(i, j)).addtoxml(commits2)
    Commit2("Offset_FP_Side" + str(side), makeinputstring(i, k)).addtoxml(commits2)


def WaterProof(i,n):
    if (sheet.cell_value(i, n)).lower() == "drip":
        return "WPQ1"

    elif (sheet.cell_value(i, n)).lower() == "side":
        return "WPQ2"

    elif (sheet.cell_value(i, n)).lower() == "facade":
        return "WPQ3"

def checkCornerAB(i):
    x = (sheet.cell_value(i, 5)).lower()
    if x == "a":
        return "CAB1"
    elif x == "b":
        return "CAB2"
    else:
        print("fel i corner AB")

def checkCornerBC(i):
    x = (sheet.cell_value(i, 6)).lower()
    if x == "b":
        return "CBC1"
    elif x == "c":
        return "CBC2"
    elif x == "-":
        return "CBC4"
    else:
        print("fel i corner BC")

def checkCornerCD(i):
    x = (sheet.cell_value(i, 7)).lower()
    if x == "c":
        return "CCD1"
    elif x == "d":
        return "CCD2"
    elif x == "-":
        return "CCD4"
    else:
        print("fel i corner CD")
def checkCornerDA(i):
    x = (sheet.cell_value(i, 8)).lower()
    if x == "d":
        return "CAD1"
    elif x == "a":
        return "CAD2"
    else:
        print("fel i corner DA")

def checkTension(i):
    x = (sheet.cell_value(i, 4)).lower()
    if x == "yes":
        return "TR4"
    elif x == "no":
        return "TR1"
    else:
        print("Fel i tensionrod")


def checkFloor(i):
    x = (sheet.cell_value(i, 4)).lower()
    if x == "yes":
        return "Floor 1"
    elif x== "no":
        return "Floor>1"
    else:
        print("Fel i tätskicktsanslutning")


def wallSetupFacade(i,s, commits2):
    if s == "straight":
        Commit2("Wall_Setup", "S").addtoxml(commits2)

    elif s == "door":
        Commit2("Wall_Setup", "D").addtoxml(commits2)
        makeadoorFacade(i,6,7,8,commits2, "Door_Type", "X_Door1", "Door_Mark")

    elif s == "window":
        Commit2("Wall_Setup", "W").addtoxml(commits2)
        makewindowFacade(i, 12, 13, 14,15, commits2, "Window_Size", "X_Window", "Window_Sill","Window_Mark")

    elif s == "window and door":
        Commit2("Wall_Setup", "DRWL").addtoxml(commits2)
        makewindowFacade(i, 12, 13, 14, 15, commits2, "Window_Size", "X_Window", "Window_Sill", "Window_Mark")
        makeadoorFacade(i,6,7,8,commits2, "Door_Type", "X_Door1", "Door_Mark")

    elif s == "door and window":
        Commit2("Wall_Setup", "DLWR").addtoxml(commits2)
        makewindowFacade(i, 12, 13, 14, 15, commits2, "Window_Size", "X_Window", "Window_Sill", "Window_Mark")
        makeadoorFacade(i, 6, 7, 8, commits2, "Door_Type", "X_Door1", "Door_Mark")

    elif s == "door and door":
        Commit2("Wall_Setup", "DD").addtoxml(commits2)
        makeadoorFacade(i, 6, 7, 8, commits2,"Door_Type","X_Door1","Door_Mark")
        makeadoorFacade(i, 9, 10, 11, commits2,"Door_Type2","X_Door2","Door_Mark2")

    elif s == "window and window":
        Commit2("Wall_Setup", "WW").addtoxml(commits2)
        makewindowFacade(i,12, 13, 14, 15, commits2, "Window_Size", "X_Window", "Window_Sill", "Window_Mark")
        makewindowFacade(i,16, 17, 18, 19, commits2, "Window_Size2", "X_Window2", "Window_Sill2", "Window_Mark2")

def wallSetup(i,s,commits2):
    """"This is where each wallinformation about setup adds"""
    if s == "straight":
        Commit2("Wall_Setup", "S").addtoxml(commits2)

    elif s == "door":
        Commit2("Wall_Setup", "D").addtoxml(commits2)
        makeadoor(i, 15, 16, 17,commits2,"Door_Type","X_Door","Door1 Serup")

    elif s == "window":
        Commit2("Wall_Setup", "W").addtoxml(commits2)
        makewindow(i,21,22,23,commits2,"Window_Type", "X_Window", "Window_Sill")

    elif s == "window and door":
        Commit2("Wall_Setup", "DRWL").addtoxml(commits2)
        makewindow(i, 21, 22, 23, commits2,"Window_Type", "X_Window", "Window_Sill")
        makeadoor(i, 15, 16, 17, commits2,"Door_Type","X_Door","Door1 Serup")

    elif s == "door and window":
        Commit2("Wall_Setup", "DLWR").addtoxml(commits2)
        makewindow(i, 21, 22, 23, commits2,"Window_Type", "X_Window", "Window_Sill")
        makeadoor(i, 15, 16, 17, commits2,"Door_Type","X_Door","Door1 Serup")

    elif s == "door and door":
        wallsetup = Commit2("Wall_Setup", "DD").addtoxml(commits2)
        makeadoor(i, 15, 16, 17, commits2,"Door_Type","X_Door","Door1 Serup")
        makeadoor(i, 18, 19, 20, commits2,"Door2_Type","X_Door2","Door2 Setup")

    elif s == "window and window":
        wallsetup = Commit2("Wall_Setup", "WW").addtoxml(commits2)
        makewindow(i, 21, 22, 23, commits2,"Window_Type", "X_Window", "Window_Sill")
        makewindow(i, 24, 25, 26, commits2,"Window2_Type", "X_Window_2","Window_Sill_2")

    else:
        print("fel i wallsetup")

def rounding(x):
    """rounds down the indata of size"""
    try:
        return int(round(int(x) / 100.0)) * 100
    except:
        print("fel i dörr/fönsterstrl")
        return None

def windowanddoorsize(a):
    """splits up the size and puts it in roundfunction"""
    list = []
    for b in a.split("x"):
        list.append(str(rounding(b)))
    return "x".join(list)

def makeadoorFacade(i,a,b,c,commits2, doortype, xdoor, doormark):
    avrundadstrl = windowanddoorsize(sheet.cell_value(i, a))
    Commit2(doortype, ("D_" + avrundadstrl)).addtoxml(commits2)
    Commit2(xdoor, str(int(sheet.cell_value(i, b)))).addtoxml(commits2)
    Commit2(doormark, str(sheet.cell_value(i, c))).addtoxml(commits2)

def makewindowFacade(i,a,b,c,d,commits2,windowsize,xwindow,windowsill,windowmark):
    """put values into the variebles winfow - size,type and sill"""
    avrundadstrl = windowanddoorsize(sheet.cell_value(i, a))
    Commit2(windowsize, "W_" + avrundadstrl).addtoxml(commits2)
    Commit2(xwindow, str(int(sheet.cell_value(i, b)))).addtoxml(commits2)
    if len(makeinputstring(i, c)) < 1: #Checks if Sill is different or same as window 1
        Sill = (makeinputstring(i, 14))
    else:
        Sill = (makeinputstring(i, c))
    Commit2(windowsill, Sill).addtoxml(commits2)
    Commit2(windowmark, str(sheet.cell_value(i, d))).addtoxml(commits2)


def makeadoor(i,a,b,c, commits2, doortype, xdoor, doorserup):
    """"puts values into the variables doors - type,size,setup"""
    avrundadstrl = windowanddoorsize(sheet.cell_value(i, a))
    Commit2(doortype, ("D_" + avrundadstrl)).addtoxml(commits2)
    Commit2(xdoor, makeinputstring(i,b)).addtoxml(commits2)
    Commit2(doorserup, str(doorsetup(i, c))).addtoxml(commits2)

def makedoorInner(i,a,b, commits2, doortype, xdoor):
    avrundadstrl = windowanddoorsize(sheet.cell_value(i, a))
    Commit2(doortype, ("D_" + avrundadstrl)).addtoxml(commits2)
    Commit2(xdoor, makeinputstring(i, b)).addtoxml(commits2)


def makewindow(i,a,b,c,commits2,windowtype,xwindow,windowsill):
    """put values into the variebles winfow - size,type and sill"""
    avrundadstrl = windowanddoorsize(sheet.cell_value(i, a))
    Commit2(windowtype, "W_" + avrundadstrl).addtoxml(commits2)
    Commit2(xwindow, str(int(sheet.cell_value(i, b)))).addtoxml(commits2)
    Commit2(windowsill,str(int(sheet.cell_value(i, c)))).addtoxml(commits2)
