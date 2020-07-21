
import xlrd
loc=(r"C:\Users\AmindaTärn\Desktop\Python\Filmeddata(KOPIA).xlsx")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

def checkvalueofcolumns(i):

    columns = str(int(sheet.cell_value(i, 1)))  # läser av antal columns och gör om det till rätt format
    return (columns)

def commit1(i, columns): #läser raderna i vald excelfil
    modulenumber = sheet.cell_value(i, 0)
    projectnumber = modulenumber[2:] #läser av modulnumer
    moduletype = modulenumber[0:2] #läser av modultypen

    columnsformat = columns + "P"

    stiffener = sheet.cell_value(i,2) #läser av stiffener och gör om det till rätt format.
    if stiffener == "Right":
        stiffener = "RS"
    elif stiffener == "Left":
        stiffener = "LS"
    else:
        stiffener = "NS"

    X1 = int(sheet.cell_value(i, 3))
    Y1 = int(sheet.cell_value(i, 4))
    Z1 = int(sheet.cell_value(i, 7))

    if columns == "6":  #lägger in värdet för olika antal dörrar
        Y2 = int(sheet.cell_value(i, 5))
        Y3 = ""
    elif columns == "8":
        Y2 = int(sheet.cell_value(i, 5))
        Y3 = int(sheet.cell_value(i, 6))
    else:
        Y2 = ""
        Y3 = ""

    #print(projectnumber, moduletype, columnsformat, stiffener, X1, Y1, Y2, Y3, Z1)

def ventholecheck(i,n): #ger värde till de olika ventholesen
    v = sheet.cell_value(i, n)
    if isinstance(v, float) == True:
        VH1 = "Yes"
        X = v
        return VH1, X
    else:
        return None, None

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

def typedoor(i,k,j,l):
    door1size = ("D_" + (sheet.cell_value(i, k)))
    placementdoor1 = str(int(sheet.cell_value(i, j)))
    door1setup = doorsetup(i,l)  # funktion som döper de olika dörröppningarna till variabler

    return door1size, placementdoor1, door1setup
def typewindow(i,m,n,o):

    windowsize = "W_" + (sheet.cell_value(i, m))
    placementwindow = int(sheet.cell_value(i, n))
    windowsill = int(sheet.cell_value(i, o))

    return windowsize,  placementwindow, windowsill

def trappning(i,l):
    p = (sheet.cell_value(i, l)).lower()
    if p == "ja":
        return "Yes"
    elif p == "nej":
        return "No"
    else:
        return None

def commit2(i,j):
    VH1_qty, X_VH1 = ventholecheck(i,12)
    VH2_qty, X_VH2 = ventholecheck(i, 13)
    VH3_qty, X_VH3 = ventholecheck(i, 14)

    #print (VH1_qty, X_VH1,VH2_qty, X_VH2,VH3_qty, X_VH3)

    trappningunder = trappning(i,10)
    trappningöver = trappning(i,11)

    s = (sheet.cell_value(i, 9)).lower()  # sänker all wallsetup till små bokstäver
    if s == "straight":
        wallsetup = "S"
    elif s == "door":
        wallsetup = "D"
        (door1size, placementdoor1, door1setup) = typedoor(i,15,16,17)
    elif s == "window":
        wallsetup = "W"
        (window1size, placementwindow, windowsill) = typewindow(i,21,22,23)
    elif s == "window and door":
        wallsetup = "DRWL"
        (window1size, placementwindow, windowsill) = typewindow(i,21,22,23)
        (door1size, placementdoor1, door1setup) = typedoor(i,15,16,17)
    elif s == "door and window":
        wallsetup = "DLWR"
        (window1size, placementwindow, windowsill) = typewindow(i,21,22,23)
        (door1size, placementdoor1, door1setup) = typedoor(i,15,16,17)
    elif s == "door and door":
        wallsetup = "DD"
        (door1size, placementdoor1, door1setup) = typedoor(i,15,16,17)
        (door2size, placementdoor2, door2setup) = typedoor(i, 18, 19,20)
    elif s == "window and window":
        wallsetup = "WW"
        (window1size, placementwindow, windowsill) = typewindow(i,21,22,23)
        (window2size, placement2window, window2sill) = typewindow(i, 24, 25, 26)
    else:
        wallsetup = "FELLLLLL"
def gothroughexcel():

    for i in range(sheet.nrows):
        if len(sheet.cell_value(i, 0)) == 5 and sheet.cell_value(i, 0)[0:2] == "VA" and sheet.cell_value(i, 9) != "" : #sorterar ut väggar som som har commits1data
            columns = checkvalueofcolumns(i) #kontroll av columns
            commit1(i,columns) #ger värden till variablerna
            j=1
            while j <= int(columns):
                #print(sheet.row_values(i))
                commit2(i,j)
                j +=1
                i +=1

gothroughexcel()
