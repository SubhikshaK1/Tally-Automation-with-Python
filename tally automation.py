''' Installation to be done from command prompt 
#pip install pyautogui
pip install openpyxl
pip install pandas'''

import pyautogui
import time
import pandas as pd

#click start for windows 10 pyautogui.click(x=1,y=887,button="left") or pyautogui.click(x=656,y=1053,button="left") for windows 11 or
pyautogui.press("win")
time.sleep(4)
print("start clicked")

#Open Tally
pyautogui.write("tally.ERP 9")
time.sleep(4)
pyautogui.press("enter")
time.sleep(8)
print("Tally opened")
pyautogui.press("w") #work in educational mode
time.sleep(1)
pyautogui.press("enter")
time.sleep(1)
pyautogui.press("down")#Silver Edition mode, (Single user)
time.sleep(2)
pyautogui.press("enter")
pyautogui.keyDown('alt')#step1: creation of a company
pyautogui.keyDown('f3')
pyautogui.keyUp('f3')
pyautogui.keyUp('alt')
time.sleep(4)
pyautogui.press("c")
time.sleep(3)

data = pd.read_excel(r'H:\Python prg\subhiksha try\svsltd.xlsx') #path wehere excel file is placed
x=data["company name"].fillna('') # data represents the excel file  to be read and ["company name"] represents the column header 
y=x.values
print(y[0])
pyautogui.write(y[0]) 
pyautogui.press("enter")
pyautogui.press("enter")
pyautogui.press("enter")
pyautogui.write("india")
pyautogui.press("enter")
pyautogui.write("Tamil Nadu")
pyautogui.press("enter")
for pe in range(1,7):
    pyautogui.press("enter")
defaultdt="1-4-"
d=data["year"].fillna('')
z=d.values
diy=int(z[0]-1)
dsy=str(diy)
finald=defaultdt+dsy
print(finald)
pyautogui.write(finald)
pyautogui.press("enter")
pyautogui.write(finald)
for pe in range(1,13):
    pyautogui.press("enter")
#creation of company done

#altering features
pyautogui.press('f11')
time.sleep(3)
pyautogui.press("enter")
pyautogui.press("enter")
pyautogui.press("enter")
pyautogui.press("y")
pyautogui.press("enter")
for peaf in range(1,9):
    pyautogui.press("enter")
pyautogui.press("y")
pyautogui.press("enter")
for ent in range(1,11):
    pyautogui.press("enter")
pyautogui.press("esc")
pyautogui.press("esc")

#inventory features
pyautogui.press("down")
pyautogui.press("enter")
pyautogui.press("enter")
pyautogui.press("enter")
pyautogui.press("enter")
pyautogui.press("y")
pyautogui.press("enter")
pyautogui.press("y")
pyautogui.press("enter")
for ent in range(1,16):
    pyautogui.press("enter")
for esc in range(1,5):
    pyautogui.press("esc")#Now at gatewayof tally
print("features altred")
pyautogui.press("enter")#acc info
pyautogui.press("down")#groups  -> ledger
pyautogui.press("enter")#single ledger create
pyautogui.press("enter")

lia=data["Liabilities"].fillna('')
l2=list(lia)
liamt=data["LAmount"].fillna('')
liamtv=liamt.values

ast=data["Assets"].fillna('')
a1=list(ast)
astamt=data["AAmount"].fillna('')
astamtv=astamt.values
astamtvl=list(astamtv)
print("astamtvl:",astamtvl)
eqsh="Equity Shares"
deb="Debentures"
pnl="Profit & Loss"
r="Reserve"
sc="Sundry Creditors"
p="Provision"
bu="Buildings"
la="Land"
fu="Furniture"
inv="Investments"
sun="Sundry Debtors"
ca="Cash"
bnk="Bank"
prp="prepaid Insurance"
out="Outstanding"
def search(listlib,listamt):
    for i in range(len(listlib)):
        if eqsh in listlib[i] :
            pyautogui.write(listlib[i])
            pyautogui.press("enter")
            pyautogui.press("enter")
            print(listlib[i])
            pyautogui.write("Capital account")
            print("Capital account")
            for ent in range(1,10):
                pyautogui.press("enter")
            pyautogui.write(str(listamt[i]))
            pyautogui.press("enter")
            pyautogui.press("enter")
            time.sleep(2)
        elif deb in listlib[i]:
            pyautogui.press("enter")
            pyautogui.write(listlib[i])
            pyautogui.press("enter")
            pyautogui.press("enter")
            print(listlib[i])
            pyautogui.write("Loans(Liability)")
            print("Loans(Liability)")
            for ent in range(1,10):
                pyautogui.press("enter")
            pyautogui.write(str(listamt[i]))
            pyautogui.press("enter")
            pyautogui.press("enter")
            pyautogui.press("enter")
            time.sleep(2)

        elif pnl in listlib[i]:
            pyautogui.press("esc")
            pyautogui.press("esc")
            pyautogui.press("a")
            pyautogui.press("p")
            pyautogui.press("enter")
            for ent in range(1,5):
                pyautogui.press("enter")
            pyautogui.write(str(listamt[i]))
            pyautogui.press("enter")
            time.sleep(2)
            pyautogui.press("enter")
            pyautogui.press("enter")
            pyautogui.press("esc")
            pyautogui.press("esc")
            pyautogui.press("c")
            time.sleep(2)
        elif r in listlib[i]:
            pyautogui.write(listlib[i])
            pyautogui.press("enter")
            pyautogui.press("enter")
            print(listlib[i])
            pyautogui.write("Reserves & surplus")
            print("Reserves & surplus")
            pyautogui.press("enter")
            pyautogui.press("enter")
            pyautogui.press("enter")
            pyautogui.write(str(listamt[i]))
            pyautogui.press("enter")
            pyautogui.press("enter")
            pyautogui.press("enter")
            time.sleep(2)

        elif r in listlib[i]:
            pyautogui.press("enter")
            pyautogui.write(listlib[i])
            pyautogui.press("enter")
            pyautogui.press("enter")
            print(listlib[i])
            pyautogui.write("Reserves & surplus")
            print("Reserves & surplus")
            for ent in range(1,10):
                pyautogui.press("enter")
            pyautogui.write(str(listamt[i]))
            pyautogui.press("enter")
            pyautogui.press("enter")
            time.sleep(3)

        elif sc in listlib[i]:
            pyautogui.press("enter")
            pyautogui.write(listlib[i])
            pyautogui.press("enter")
            pyautogui.press("enter")
            print(listlib[i])
            pyautogui.write(sc)
            print(sc)
            for ent in range(1,12):
                pyautogui.press("enter")
            time.sleep(4)
            pyautogui.write(str(listamt[i]))
            pyautogui.press("enter")
            pyautogui.press("enter")
            pyautogui.press("enter")
            pyautogui.press("enter")
            pyautogui.press("enter")
            time.sleep(3)
            
        elif p in listlib[i]:
            pyautogui.press("enter")
            pyautogui.write(listlib[i])
            pyautogui.press("enter")
            pyautogui.press("enter")
            print(listlib[i])
            pyautogui.write(p)
            print(p)
            for ent in range(1,4):
                pyautogui.press("enter")
            pyautogui.write(str(listamt[i]))
            pyautogui.press("enter")
            pyautogui.press("enter")
            pyautogui.press("enter")
            time.sleep(3)
        elif out in listlib[i]:
            pyautogui.press("enter")
            pyautogui.write(listlib[i])
            pyautogui.press("enter")
            pyautogui.press("enter")
            print(listlib[i])
            pyautogui.write("Current liability")
            print("Current liability")
            for ent in range(1,12):
                pyautogui.press("enter")
            pyautogui.write(str(listamt[i]))
            pyautogui.press("enter")
            pyautogui.press("enter")
            pyautogui.press("enter")
            time.sleep(3)
def ast1(listast,listastamt):
            for j in range(len(listast)):
                if bu in listast[j]:
                    pyautogui.write(listast[j])
                    pyautogui.press("enter")
                    pyautogui.press("enter")
                    print(listast[j])
                    pyautogui.write("fixed asset")
                    print("fixed asset")
                    for ent in range(1,10):
                        pyautogui.press("enter")
                    pyautogui.write(str(listastamt[j]))
                    pyautogui.press("enter")
                    pyautogui.press("enter")
                    time.sleep(2)
                    print(listast[j])
                    print("fixed asset")
                    print(str(listastamt[j]))
                elif la in listast[j]:
                    pyautogui.press("enter")
                    pyautogui.press("enter")
                    pyautogui.write(listast[j])
                    pyautogui.press("enter")
                    pyautogui.press("enter")
                    pyautogui.write("fixed asset")
                    for ent in range(1,10):
                        pyautogui.press("enter")
                    pyautogui.write(str(listastamt[j]))
                    pyautogui.press("enter")
                    pyautogui.press("enter")
                    time.sleep(2)
                    print(listast[j])
                    print("fixed asset")
                    print(str(listastamt[j]))
                elif fu in listast[j]:
                    pyautogui.press("enter")
                    pyautogui.press("enter")
                    pyautogui.write(listast[j])
                    pyautogui.press("enter")
                    pyautogui.press("enter")
                    pyautogui.write("fixed asset")
                    for ent in range(1,10):
                        pyautogui.press("enter")
                    pyautogui.write(str(listastamt[j]))
                    pyautogui.press("enter")
                    pyautogui.press("enter")
                    time.sleep(2)
                    print(listast[j])
                    print("fixed asset")
                    print(str(listastamt[j]))
                elif inv in listast[j]:
                    pyautogui.press("enter")
                    pyautogui.press("enter")
                    pyautogui.write(listast[j])
                    pyautogui.press("enter")
                    pyautogui.press("enter")
                    pyautogui.write(inv)
                    for ent in range(1,10):
                        pyautogui.press("enter")
                    pyautogui.write(str(listastamt[j]))
                    pyautogui.press("enter")
                    pyautogui.press("enter")
                    time.sleep(2)
                    print(listast[j])
                    print(inv)
                    print(str(listastamt[j]))
                elif sun in listast[j]:
                    pyautogui.press("enter")
                    pyautogui.press("enter")
                    pyautogui.write(listast[j])
                    pyautogui.press("enter")
                    pyautogui.press("enter")
                    pyautogui.write(sun)
                    for ent in range(1,12):
                        pyautogui.press("enter")
                    pyautogui.write(str(listastamt[j]))
                    pyautogui.press("enter")
                    pyautogui.press("enter")
                    pyautogui.press("enter")
                    pyautogui.press("enter")
                    time.sleep(2)
                    print(listast[j])
                    print(sun)
                    print(str(listastamt[j]))
                elif bnk in listast[j]:
                    pyautogui.press("enter")
                    pyautogui.press("enter")
                    pyautogui.write(listast[j])
                    pyautogui.press("enter")
                    pyautogui.press("enter")
                    pyautogui.write("Bank Account")
                    for ent in range(1,9):
                        pyautogui.press("enter")
                    pyautogui.write(str(listastamt[j]))
                    pyautogui.press("enter")
                    pyautogui.press("enter")
                    time.sleep(2)
                    print(listast[j])
                    print("bank ac")
                    print(str(listastamt[j]))
                elif prp in listast[j]:
                    pyautogui.press("enter")
                    pyautogui.press("enter")
                    pyautogui.write(listast[j])
                    pyautogui.press("enter")
                    pyautogui.press("enter")
                    pyautogui.write("Current asset")
                    for ent in range(1,11):
                        pyautogui.press("enter")
                    pyautogui.write(str(listastamt[j]))
                    pyautogui.press("enter")
                    pyautogui.press("enter")
                    time.sleep(2)
                    print(listast[j])
                    print("current asset")
                    print(str(listastamt[j]))
                    pyautogui.press("enter")
                    pyautogui.press("enter")        
search(l2,liamtv)
ast1(a1,astamtvl)
pyautogui.press("esc")
pyautogui.press("esc")
pyautogui.press("esc")
pyautogui.press("esc")

#Opening balance sheet
pyautogui.press("b")
pyautogui.keyDown('alt')
pyautogui.keyDown('f1')
pyautogui.keyUp('f1')
pyautogui.keyUp('alt')
time.sleep(4)
