import xml.etree.ElementTree as ET
import os
os.chdir(r"C:\Users\AmindaTärn\Desktop\Python\Statesomxml")
root = ET.parse('007VA001_203006 (24).xml').getroot()
for i in root:
    for j in i:
        #print(j.tag,j.attrib)
        for k in j:
            print (k)
            print(k.tag,k.attrib)

print("********")
for i in root:
    print (i)

print("********")
for j in root:
    for k in j:
        print (k)
print(root)