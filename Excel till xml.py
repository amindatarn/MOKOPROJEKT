
import xlrd
loc=(r"C:\Users\AmindaTärn\Desktop\Python\Filmeddata(KOPIA).xlsx")
import xml.etree.ElementTree as ET



def read_line(): #läser raderna i vald excelfil
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    sheet.cell_value(0, 0)

    print(sheet.row_values(5))
    #i=0
    # while i >= 0:
    #    print(sheet.row_values(i))
    #   i+=1

def create_XML():
    xml_doc = = ET.Element("root")



read_line()
