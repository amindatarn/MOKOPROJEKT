import xlrd

class Commit1:
    def __init__(self, modulenumber):
        modulenumber = self.modulenumber

    def getmodulenumber(self):
        return excelPathCellValue(i,0)

    def createCommit1(self):
        putInTree("Module_number", modulenumber[2:])            # läser av modulnumer
        putInTree("Mudule_Type", modulenumber[0:2])                  # läser av modultypen
        putInTree("Columns", makeinputstring("columns", i, 1))      # görcolumns till string

        columnsformat = columns + "P"                   # formaterar columns

        putInTree("Stiffener", formatStiffener(excelPathCellValue(i, 2))) # läser av stiffener och gör om det till rätt format.

        putInTree("X1",makeinputstring("X1", i, 3))
        putInTree("Y1" ,makeinputstring("Y1", i, 4))
        putInTree("Z1",makeinputstring("Z1", i, 7))

        #projectNumber, ritadAv, uppdragsansvarig, status, datum, regulations = getDataFromInput()

        if columns == "6":  # lägger in värdet för olika antal dörrar
            putInTree("Y2",makeinputstring("Y2", i, 5))
        elif columns == "8":
            putInTree("Y2" ,makeinputstring("Y2", i, 5))
            putInTree("Y3", makeinputstring("Y3", i, 6))

        filnamn = "007" + modulenumber.zfill(3) + "_"

    def getexcelParh(self):
        loc = (r"C:\Users\AmindaTärn\Desktop\Python\Filmeddata(KOPIA).xlsx")
        wb = xlrd.open_workbook(loc)
        sheet = wb.sheet_by_index(0)

    def formatStiffener(stiffener):
        if stiffener == "Right":
            return "RS"
        elif stiffener == "Left":
            return "LS"
        else:
            return "NS"

    def makeinputstring(variabel, i, f):

        variabel = sheet.cell_value(i, f)
        if type(variabel) == int:
            return str(variabel)
        elif type(variabel) == float:
            return str(int(float(variabel)))
        else:
            return str(variabel)

    def putInTree(textName, elementText):
        Classxmlmall.SubElement3("commits1", "Committed",textName, elementText)
