import xml.etree.ElementTree as ET
import xlrd
loc=(r"C:\Users\AmindaTärn\Desktop\Python\Filmeddata(KOPIA).xlsx")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

def createxmlmall(i):
    # build a tree structure
    root = "root" + str(i)
    root = ET.Element("state")

    model = ET.SubElement(root, "model")
    model.text = (r"C:\Users\Linn Arstad\AppData\Local\Temp\Temporary\Deployed\Files\For\DA\tmpFF28.tmp")

    dataid = ET.SubElement(root, "dataids")

    application = ET.SubElement(root, "application")
    application.text = ("SIBS Configurator")

    safecookie = ET.SubElement(root, "safecookie")

    safestep1 = ET.SubElement(safecookie, "safe-step", name= "Project Information")
    safestep2 = ET.SubElement(safecookie, "safe-step", name="Indata")

    commits1 = ET.SubElement(safestep1, "commits")
    commits2 = ET.SubElement(safestep2, "commits")

    steps = ET.SubElement(root, "steps")
    prev = ET.SubElement(steps, "prev")
    step = ET.SubElement(prev, "step").text =("Project Information")

    current = ET.SubElement(steps, "current").text=("Indata")
    next = ET.SubElement(steps, "next")

    lastproxy = ET.SubElement(root, "last-proxy").text="tcserver0"

    filnamn = commit1(i, commits1)

    tree = ET.ElementTree(root)  # sparar hela trädet i variabeln tree

    return tree, commits1, commits2,filnamn

def getdatafromexcel(commits1, projectnumber, moduletype, columnsformat, stiffener, X1, Y1, Z1, Y2, Y3,): #lägger in varibler

    indatatillxml(commits1, "Project_number", "007")
    indatatillxml(commits1, "Module_Type", str(moduletype))
    indatatillxml(commits1, "Module_number",str(projectnumber))
    indatatillxml(commits1, "Ritad_Av", "AMINDA TÄRN")
    indatatillxml(commits1, "Uppdragsansvarig","PATRIK JENSEN")
    indatatillxml(commits1, "Status", "FOR PRODUCTION")
    indatatillxml(commits1, "Datum", "2020-09-11")
    indatatillxml(commits1, "Regulations", "001001")
    indatatillxml(commits1, "Columns", columnsformat)
    indatatillxml(commits1, "Stiffener", stiffener)

    indatatillxml(commits1, "X1", str(X1))
    indatatillxml(commits1, "Y1", str(Y1))
    indatatillxml(commits1, "Y2", str(Y2))
    indatatillxml(commits1, "Y3", str(Y3))
    indatatillxml(commits1, "Z1", str(Z1))


def indatatillxml(commit ,variabel1,variabel2): #skapar alla attribut för commits1
    ET.SubElement(commit, "committed", name=variabel1).text = variabel2

def savexml(tree,filnamn):
    import os
    os.chdir(r"C:\Users\AmindaTärn\Desktop\Python\xmlfiler")  # ändrar plats för filer
    mappnamn = "stateexempel.xml"
    try:
        os.chdir(r"C:\Users\AmindaTärn\Desktop\Python\xmlfiler\\" + mappnamn)
        tree.write(filnamn) #, encoding="UTF-8")  # Namnet på ny fil
    except:
        os.mkdir(mappnamn)  # Namnet på ny mapp
        os.chdir(r"C:\Users\AmindaTärn\Desktop\Python\xmlfiler\\" + mappnamn)
        tree.write(filnamn) #, encoding="UTF-8")  # Namnet på ny fil

#EXCEL
def makeinputstring(variabel, i, f):

    variabel = sheet.cell_value(i,f)
    if type(variabel) == int:
        return str(variabel)
    elif type(variabel) == float:
        return str(int(float(variabel)))
    else:
        return str(variabel)


def commit1(i, commits1): #läser raderna i vald excelfil
    modulenumber = sheet.cell_value(i, 0)
    projectnumber = modulenumber[2:] #läser av modulnumer
    moduletype = modulenumber[0:2] #läser av modultypen

    columns = makeinputstring("columns",i,1)
    columnsformat = columns + "P"

    stiffener = sheet.cell_value(i,2) #läser av stiffener och gör om det till rätt format.
    if stiffener == "Right":
        stiffener = "RS"
    elif stiffener == "Left":
        stiffener = "LS"
    else:
        stiffener = "NS"

    X1 = makeinputstring("X1",i,3)
    Y1 = makeinputstring("Y1",i,4)
    Z1 = makeinputstring("Z1",i,7)

    if columns == "6":  #lägger in värdet för olika antal dörrar
        Y2 = makeinputstring("Y2",i,5)
        Y3 = ""
    elif columns == "8":
        Y2 = makeinputstring("Y2",i,5)
        Y3 = makeinputstring("Y3",i,6)
    else:
        Y2 = ""
        Y3 = ""

    filnamn = "007" + modulenumber.zfill(3) + "_"
    getdatafromexcel(commits1, projectnumber, moduletype, columnsformat, stiffener, X1, Y1, Z1, Y2, Y3)
    print(projectnumber, moduletype, columnsformat, stiffener, X1, Y1, Y2, Y3, Z1)
    return filnamn


def ventholecheck(i,commits2): #ger värde till de olika ventholesen
    n = 12
    k = 1
    while n <=14:
        v = sheet.cell_value(i, n)
        if isinstance(v, float) == True:
            attra = Commit2(str("X_VH" + str(k)),str(int(v)))
            attrb = Commit2(str("VH" + str(k) + "_qty"),"Yes")
            attra.addtoxml(commits2)
            attrb.addtoxml(commits2)
            k += 1
        n +=1
    else:
        return

def doorsetup(i,l):
    p = (sheet.cell_value(i, l)).lower()
    if p == "ytterdörr":  # döper de olika dörrarna till rätt benämning
        door1setup = "DS1"
    elif p == "inv lgh-dörr":
        door1setup = "DS2"
    elif p == "passage":
        door1setup = "DS3"
    else:
        door1setup = None
        print("Fel i doorsetup")
    return door1setup

def trappning(i,l,commits2,trappning):
    p = (sheet.cell_value(i, l)).lower()
    if p == "ja":
        trapp = Commit2(trappning, "Yes")
        trapp.addtoxml(commits2)
    elif p == "nej":
        trapp = Commit2(trappning, "No")
        trapp.addtoxml(commits2)

class Commit2:
    def __init__(self, name,text):
        self.name = name
        self.text = text

    def addtoxml(self,commits2):
        indatatillxml(commits2,self.name ,self.text)

def commit2(i,commits2,x):

    wall = Commit2("Wall_number", str(x))
    wall.addtoxml(commits2)

    ventholecheck(i,commits2)

    trappning(i,10,commits2,"trappningunder")
    trappning(i,11,commits2,"trappningöver")

    s = (sheet.cell_value(i, 9)).lower()  # sänker all wallsetup till små bokstäver

    if s == "straight":
        wallsetup = Commit2("Wall_setup", "S")
        wallsetup.addtoxml(commits2)

    elif s == "door":
        wallsetup = Commit2("Wall_setup", "D")
        wallsetup.addtoxml(commits2)
        makeadoor(i, 15, 16, 17,1,commits2)

    elif s == "window":
        wallsetup = Commit2("Wall_setup", "W")
        wallsetup.addtoxml(commits2)
        makewindow(i,21,22,23,1,commits2)

    elif s == "window and door":
        wallsetup = Commit2("Wall_setup", "DRWL")
        wallsetup.addtoxml(commits2)
        makewindow(i, 21, 22, 23, 1, commits2)
        makeadoor(i, 15, 16, 17, 1, commits2)

    elif s == "door and window":
        wallsetup = Commit2("Wall_setup", "DLWR")
        wallsetup.addtoxml(commits2)
        makewindow(i, 21, 22, 23, 1, commits2)
        makeadoor(i, 15, 16, 17, 1, commits2)

    elif s == "door and door":
        wallsetup = Commit2("Wall_setup", "DD")
        wallsetup.addtoxml(commits2)
        makeadoor(i, 15, 16, 17, 1, commits2)
        makeadoor(i, 18, 19, 20, 2, commits2)

    elif s == "window and window":
        wallsetup = Commit2("Wall_setup", "WW")
        wallsetup.addtoxml(commits2)
        makewindow(i, 21, 22, 23, 1, commits2)
        makewindow(i, 24, 25, 26, 2, commits2)

    else:
        wallsetup = "FELLLLLL"

def makeadoor(i,a,b,c,n,commits2):
    doorsize = Commit2("Door" + str(n) + "_Type", ("D_" + (sheet.cell_value(i, a))))
    placementdoor = Commit2("X_Door" + str(n), str(int(sheet.cell_value(i, b))))
    dsetup = Commit2("Door" + str(n) + "setup", str(doorsetup(i, c)))
    doorsize.addtoxml(commits2)
    placementdoor.addtoxml(commits2)
    dsetup.addtoxml(commits2)

def makewindow(i,a,b,c,n,commits2):
    windowsize = Commit2("Window" + str(n) + "_Type", "W_" + (sheet.cell_value(i, a)))
    placementwindow = Commit2( "X_Window", str(int(sheet.cell_value(i, b))))
    windowsill = Commit2("Window_sill",str(sheet.cell_value(i, c)))
    windowsize.addtoxml(commits2)
    placementwindow.addtoxml(commits2)
    windowsill.addtoxml(commits2)


def gothroughexcel():
    i=1
    while i < (sheet.nrows):
        if len(sheet.cell_value(i, 0)) == 5 and sheet.cell_value(i, 0)[0:2] == "VA": #and sheet.cell_value(i,9) != "":  # sorterar ut väggar som som har commits1data
            columns = makeinputstring("columns", i, 1)  # kontroll av columns
            tree, commits1, commits2, filnamn = createxmlmall(i)
            savexml(tree, filnamn + "default" + ".xml")
            x=1
            k=i
            while x<= int(columns):
                if not sheet.cell_value(k, 15)[0:2].lower() == "sa": #sorterar bort alla väggar som är "samma som"
                    tree, commits1, commits2, filnamn = createxmlmall(i)
                    commit2(k, commits2, x)
                    savexml(tree, filnamn + "20300" + str(x).zfill(3) + ".xml")
                k+=1
                x+=1
        else:
            print(i)
        i+=1


gothroughexcel()