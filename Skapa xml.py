

def createxml():
    indata1 = ("modulnumer")
    indata2 = ("columns")
    import xml.etree.ElementTree as ET

    # build a tree structure
    root = ET.Element("state")

    model = ET.SubElement(root, "model")
    model.text = (r"C:\Users\Linn Arstad\AppData\Local\Temp\Temporary\Deployed\Files\For\DA\tmpFF28.tmp")

    dataid = ET.SubElement(root, "dataids")

    application = ET.SubElement(root, "application")
    application.text = ("SIBS Configurator")

    safecookie = ET.SubElement(root, "safecookie")
    safestep = ET.SubElement(safecookie, "safecookie", name= "Project Inormation")
    safestep = ET.SubElement(safecookie, "Project Information", attr="007")

    commits = ET.SubElement(safestep, "commits")

    commitedname02 = ET.SubElement(commits, "committed name", name="Module_Type", number="VA")

    tree = ET.ElementTree(root)  # sparar hela trädet i variabeln tree
    savexml(tree)

""" commitedname03 = ET.SubElement(commits, "committed name", name="Module_number", number=indata1)
commitedname04 = ET.SubElement(commits, "committed name", name="Ritad_Av", number="AMINDA TÄRN")
commitedname05 = ET.SubElement(commits, "committed name", name="Uppdragsansvarig", number="PATRIK JENSEN")
commitedname06 = ET.SubElement(commits, "committed name", name="Status", number="FOR PRUDUCTION")
commitedname07 = ET.SubElement(commits, "committed name", name="Datum", date="0097")
commitedname08 = ET.SubElement(commits, "committed name", name="Regulations", number="0097")
commitedname09 = ET.SubElement(commits, "committed name", name="Columns", number=indata2)
commitedname10 = ET.SubElement(commits, "committed name", name="Project_number", number="0097")"""


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
            tree.write("nyfil.xml")  # Namnet på ny fil
            break
        except:
            print("ehj")
            n = int(n) + 1
            print(n)

createxml()
