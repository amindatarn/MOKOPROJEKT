
import Skapaxml
import getpass

global username
username = getpass.getuser()
username = ("".join([(" "+i if i.isupper() else i) for i in username]).strip().upper())

def FasadWalls(Uppdragsansvarig, Status, Datum, Regulations,foldername, exceldokument):

    sheet = Skapaxml.getExcel(exceldokument)
    i = 1
    while i < (sheet.nrows):
        if i >1 and len(sheet.cell_value(i, 1)) > 0:
            projectNumber = str(sheet.cell_value(i, 1))[0:3]
            tree, commits1, safecookie, steps, prev = Skapaxml.createxmlmall(username,Uppdragsansvarig, Status, Datum, Regulations,projectNumber)
            commits2 = Skapaxml.createxmlmallejdefault(i, safecookie, steps, prev)
            filnamn = Skapaxml.facadeWallsCommits(i, commits2)
            Skapaxml.savexml(tree, filnamn + ".tcs", foldername)
        else:
            print("FEL")
            pass
        i += 1

    quit()


def OuterWalls(Uppdragsansvarig, Status, Datum, Regulations, foldername, exceldokument):
    """"main function creating xml from exceldata and then saving it"""
    sheet = Skapaxml.getExcel(exceldokument)

    i = 1
    while i < (sheet.nrows):
        if len(sheet.cell_value(i, 0)) == 5 and sheet.cell_value(i, 0)[0:2] == "VA":  # and sheet.cell_value(i,9) != "":  # sorterar ut väggar som som har commits1data
            projectNumber =  str(sheet.cell_value(i, 0))[0:2]
            columns = Skapaxml.makeinputstring( i, 1)  # kontroll av columns
            tree, commits1, safecookie, steps,prev = Skapaxml.createxmlmall(username,Uppdragsansvarig, Status, Datum,Regulations,projectNumber)
            filnamn = Skapaxml.outerWallsCommit1(i)  # adds subElements to tree in first commit and returns filname
            Skapaxml.createxmlmalldefault(steps)
            Skapaxml.savexml(tree, filnamn + "default" + ".xml", foldername)
            x = 1
            k = i
            while x <= int(columns):
                if not sheet.cell_value(k, 15)[0:2].lower() == "sa":  # sorterar bort alla väggar som är "samma som"
                    tree, commits1, safecookie, steps, prev = Skapaxml.createxmlmall(username,Uppdragsansvarig, Status, Datum,Regulations,projectNumber)
                    filnamn = Skapaxml.outerWallsCommit1(i)  # adds subElements to tree in first commit and returns filname
                    commits2 = Skapaxml.createxmlmallejdefault(i, safecookie, steps, prev)
                    Skapaxml.outerWallsCommit2(k, x, commits2)
                    Skapaxml.savexml(tree, filnamn + "20300" + str(x).zfill(3) + ".xml", foldername)
                k += 1
                x += 1
        else:
            pass
        i += 1

    quit()

def Balcony(Uppdragsansvarig, Status, Datum, Regulations,foldername, exceldokument):

    sheet = Skapaxml.getExcel(exceldokument)
    i = 1
    while i < (sheet.nrows):
        if i >1 and len(sheet.cell_value(i, 1)) > 0:
            projectNumber = str(sheet.cell_value(i, 1))[0:3]
            tree, commits1, safecookie, steps, prev = Skapaxml.createxmlmall(username,Uppdragsansvarig, Status, Datum, Regulations,projectNumber)
            commits2 = Skapaxml.createxmlmallejdefault(i, safecookie, steps, prev)
            filnamn = Skapaxml.BalconyCommits(i, commits2)
            Skapaxml.savexml(tree, filnamn + ".xml", foldername)
            Skapaxml.savexml(tree, filnamn + ".tcs", foldername)
        else:
            print("FEL")
            pass
        i += 1

    quit()


def InnerWalls(Uppdragsansvarig, Status, Datum, Regulations, foldername, exceldokument):
    """"main function creating xml from exceldata and then saving it"""

    def loop(i):
        tree, commits1, safecookie, steps, prev = Skapaxml.createxmlmall(username, Uppdragsansvarig, Status, Datum,
                                                                         Regulations, projectNumber)
        commits2 = Skapaxml.createxmlmallejdefault(i, safecookie, steps, prev)
        filnamn = Skapaxml.innerWalls(i, commits2)  # adds subElements to tree in first commit and returns filname
        Skapaxml.savexml(tree, filnamn + "204001".zfill(3) + ".xml", foldername)


    sheet = Skapaxml.getExcel(exceldokument)
    i = 1
    while i < (sheet.nrows):
        if len(sheet.cell_value(i, 0)) == 8 and sheet.cell_value(i, 0)[3:5] == "VA":  # and sheet.cell_value(i,9) != "":  # sorterar ut väggar som som har commits1data
            print(sheet.cell_value(i, 0))
            projectNumber = str(sheet.cell_value(i, 0))[0:3]
            loop(i)
            k = i
            k += 1
            while (sheet.cell_value(k, 0)) == "":
                print(sheet.cell_value(k, 1))
                loop(i)

                k += 1



    quit()
