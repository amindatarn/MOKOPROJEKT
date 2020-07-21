
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

def commit2(i,j):
    wallnumber = j
    s = (sheet.cell_value(i, 9)).lower() #sänker all wallsetup till små bokstäver
    if s == "straight":
        wallsetup = "S"
    elif s == "door":
        wallsetup = "D"
        door1size = "D_"+ (sheet.cell_value(i, 15))
        placementdoor1 = 
        print(door1size)
    elif s == "window":
        wallsetup = "W"
        window1size = "W_" + (sheet.cell_value(i, 18))
    elif s == "window and door":
        wallsetup = "DRWL"
    elif s == "door and window":
        wallsetup = "DLWR"
    elif s == "door and door":
        wallsetup = "DD"
    elif s == "window and window":
        wallsetup = "WW"
    else:
        wallsetup = "FELLLLLL"



def gothroughexcel():

    for i in range(sheet.nrows):
        if len(sheet.cell_value(i, 0)) == 5 and sheet.cell_value(i, 0)[0:2] == "VA" and sheet.cell_value(i, 9) != "" : #sorterar ut väggar som som har commits1data
            columns = checkvalueofcolumns(i) #kontroll av columns
            commit1(i,columns) #ger värden till variablerna
            j=1
            while j <= int(columns):
                print(sheet.row_values(i))
                commit2(i,j)
                j +=1
                i +=1





gothroughexcel()
