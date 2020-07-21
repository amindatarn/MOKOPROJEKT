import xml.etree.ElementTree as ET

def createxmlmall():
    # build a tree structure
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

    getdatafromexcel(commits1,commits2)

    steps = ET.SubElement(root, "steps")
    prev = ET.SubElement(steps, "prev")
    step = ET.SubElement(prev, "step").text =("Project Information")

    current = ET.SubElement(steps, "current").text=("Indata")
    next = ET.SubElement(steps, "next")

    lastproxy = ET.SubElement(root, "last-proxy").text="tcserver0"


    tree = ET.ElementTree(root)  # sparar hela trädet i variabeln tree
    savexml(tree)

def getdatafromexcel(commits1, commits2): #lägger in varibler
    indatatillxml(commits1, "Project_number","007")
    indatatillxml(commits1, "Module_Type","VA")
    indatatillxml(commits1, "Module_number","001")
    indatatillxml(commits1, "Ritad_Av", "AMINDA TÄRN")
    indatatillxml(commits1, "Uppdragsansvarig","PATRIK JENSEN")
    indatatillxml(commits1, "Status", "FOR PRODUCTION")
    indatatillxml(commits1, "Datum", "2020-09-11")
    indatatillxml(commits1, "Regulations", "001001")
    indatatillxml(commits1, "Columns", "8P")
    indatatillxml(commits1, "Stiffener", "RS")
    indatatillxml(commits1, "X1", "1800")
    indatatillxml(commits1, "Y1", "3000")
    indatatillxml(commits1, "Y2", "7000")
    indatatillxml(commits1, "Y3", "8900")
    indatatillxml(commits1, "Z1", "2900")

    indatatillxml(commits2, "Wall_Number", "5")
    indatatillxml(commits2, "Wall_Setup", "S")


def indatatillxml(commit ,variabel1,variabel2): #skapar alla attribut för commits1
    ET.SubElement(commit, "committed", name=variabel1).text = variabel2

def savexml(tree):
    import os
    os.chdir(r"C:\Users\AmindaTärn\Desktop\Python\xmlfiler")  # ändrar plats för filer
    mappnamn = "testmapp"
    n = 1
    while True:
        print(n)
        try:
            mappnamn = mappnamn + str(n)
            os.mkdir(mappnamn)  # Namnet på ny mapp
            os.chdir(r"C:\Users\AmindaTärn\Desktop\Python\xmlfiler\\" + mappnamn)
            tree.write("nyfil.xml", encoding = "UTF-8")  # Namnet på ny fil
            break
        except:
            print("ehj")
            n = int(n) + 1
            print(n)

createxmlmall()
