
from pywinauto import application
from pywinauto.application import Application
import time
import pyautogui
import xlwings as xw
from xlwings import Range, constants
import pydirectinput
from datetime import date



#Extragere date din excelul de bovine

wb = xw.Book("D:\\Sacrif. bovina 21.05.2024.xls").sheets['Foaie1']


today = date.today()
formatted_date = today.strftime('%m.%d.%Y')
if xw.Book("D:\\Sacrif. bovina 21.05.2024.xls").sheets['Date'].range("I9").value != formatted_date:
    #xw.Book("D:\\Sacrif. bovina 21.05.2024.xls").sheets['Date'].range("F7").value = "NU"
    #xw.Book("D:\\Sacrif. bovina 21.05.2024.xls").sheets['Date'].range("F9").value = "NU"
    xw.Book("D:\\Sacrif. bovina 21.05.2024.xls").sheets['Date'].range("I9").value = formatted_date
    xw.Book("D:\\Sacrif. bovina 21.05.2024.xls").sheets['Date'].range("G3").value = ""



lastCell = wb.range('E' + str(wb.cells.last_cell.row)).end('up').row
nrReceptie = 0
if xw.Book("D:\\Sacrif. bovina 21.05.2024.xls").sheets['Date'].range("G3").value is None:
    nrReceptie = 9
else:
    nrReceptie = int(xw.Book("D:\\Sacrif. bovina 21.05.2024.xls").sheets['Date'].range("G3").value)
    nrReceptie = nrReceptie + 8

#print(xw.Book("D:\\Sacrif. bovina 21.05.2024.xls").sheets['Date'].range("G3").value)


nrCrotal = wb.range("E" + str(nrReceptie) + ":E" + str(lastCell)).value
varsta = wb.range("G" + str(nrReceptie) + ":G" + str(lastCell)).value
sex = wb.range("H" + str(nrReceptie) + ":H" + str(lastCell)).value
rasa = wb.range("I" + str(nrReceptie) + ":I" + str(lastCell)).value
propietar = wb.range("J" + str(nrReceptie) + ":J" + str(lastCell)).value
localitate = wb.range("K" + str(nrReceptie) + ":K" + str(lastCell)).value
codExploatatie = wb.range("L" + str(nrReceptie) + ":L" + str(lastCell)).value
nrPasaport = wb.range("M" + str(nrReceptie) + ":M" + str(lastCell)).value
masina = wb.range("N" + str(nrReceptie) + ":N" + str(lastCell)).value
varsta = [int(varsta) for varsta in varsta]
doctor = int(xw.Book("D:\\Sacrif. bovina 21.05.2024.xls").sheets['Date'].range("E2").value)



#Introducere date in p01

app = Application(backend="uia").start("C:\\vifout\\Putty\\putty.exe")
pid = application.process_from_module(module = "C:\\vifout\\Putty\\putty.exe")
#app.PuTTYConfiguration.print_control_identifiers()
btnOpen = app.PuTTYConfiguration.child_window(title="Open", auto_id="1009", control_type="Button")
srvOption = app.PuTTYConfiguration.child_window(title="srvvif57", control_type="ListItem")
srvOption.select()
btnOpen.click()
time.sleep(1)
app = Application(backend="uia").connect(process=pid)
#pp.VIF5_7.print_control_identifiers()
pyautogui.typewrite("p01")
pyautogui.press("enter")
time.sleep(1)
pyautogui.press("enter", presses=4)
pyautogui.press('r')
pyautogui.press('e')
pyautogui.press('enter', presses=2)

###################################


propietarAnterior = wb.range("J" + str(nrReceptie)).value
#print(propietarAnterior)

pyautogui.press("f3")
pyautogui.press("enter")
pyautogui.typewrite("0")
pyautogui.press("enter")
pyautogui.typewrite(masina[0])
pyautogui.press("enter")
pyautogui.typewrite(str(doctor))
pyautogui.press("enter")
pyautogui.press("f2")
pyautogui.typewrite("10301")
pyautogui.press("enter")
pyautogui.typewrite(nrCrotal[0])
pyautogui.press("enter")
pyautogui.press("enter")
pyautogui.typewrite(nrCrotal[0][:2])
pyautogui.press("enter")
pyautogui.typewrite(nrPasaport[0])
pyautogui.press("enter")
pyautogui.typewrite(rasa[0])
pyautogui.press("enter")
pyautogui.typewrite(sex[0])
pyautogui.press("enter")
pyautogui.typewrite(str(varsta[0]))
pyautogui.press("enter")
pyautogui.press("enter")
pyautogui.typewrite(propietar[0])
pyautogui.press("enter")
pyautogui.typewrite(propietar[0])
pyautogui.press("enter")
pyautogui.typewrite(codExploatatie[0])
pyautogui.press("enter")
pyautogui.typewrite(localitate[0])
pyautogui.press("enter")
pyautogui.press("f2")
pyautogui.press("f2")

for i in range(1,len(propietar)):
    if propietarAnterior != propietar[i]:
        pyautogui.press("f4")
        pyautogui.press("d")
        time.sleep(1)
        pyautogui.press("f3")
        time.sleep(1)
        pyautogui.press("enter")
        pyautogui.typewrite("0")
        pyautogui.press("enter")
        time.sleep(1)
        pyautogui.typewrite(masina[i])
        pyautogui.press("enter")
        pyautogui.typewrite(str(doctor))
        pyautogui.press("enter")
        pyautogui.press("f2")
        time.sleep(1)
        pyautogui.typewrite("10301")
        pyautogui.press("enter")
        pyautogui.typewrite(nrCrotal[i])
        pyautogui.press("enter")
        time.sleep(1)
        pyautogui.press("enter")
        pyautogui.typewrite(nrCrotal[i][:2])
        pyautogui.press("enter")
        pyautogui.typewrite(nrPasaport[i])
        pyautogui.press("enter")
        pyautogui.typewrite(rasa[i])
        pyautogui.press("enter")
        pyautogui.typewrite(sex[i])
        pyautogui.press("enter")
        pyautogui.typewrite(str(varsta[i]))
        pyautogui.press("enter")
        pyautogui.press("enter")
        pyautogui.typewrite(propietar[i])
        pyautogui.press("enter")
        pyautogui.typewrite(propietar[i])
        pyautogui.press("enter")
        pyautogui.typewrite(codExploatatie[i])
        pyautogui.press("enter")
        pyautogui.typewrite(localitate[i])
        pyautogui.press("enter")
        pyautogui.press("f2")
        pyautogui.press("f2")
        time.sleep(1)
        propietarAnterior = propietar[i]
    else:
        pyautogui.press("enter")
        pyautogui.typewrite(nrCrotal[i])
        time.sleep(1)
        pyautogui.press("enter")
        pyautogui.press("enter")
        pyautogui.typewrite(nrCrotal[i][:2])
        pyautogui.press("enter")
        pyautogui.typewrite(nrPasaport[i])
        pyautogui.press("enter")
        pyautogui.typewrite(rasa[i])
        pyautogui.press("enter")
        pyautogui.typewrite(sex[i])
        pyautogui.press("enter")
        pyautogui.typewrite(str(varsta[i]))
        pyautogui.press("enter")
        pyautogui.press("f2")
        pyautogui.press("f2")
        time.sleep(1)
        propietarAnterior = propietar[i]
pyautogui.press("f4")
pyautogui.press("d")
time.sleep(1)



xw.Book("D:\\Sacrif. bovina 21.05.2024.xls").sheets['Date'].range("G3").value = lastCell - 7




#Validare date in p02

pydirectinput.PAUSE = 0.05

crotaleSortate = sorted(nrCrotal)

app = Application(backend="uia").start("C:\\vifout\\Putty\\putty.exe")
pid = application.process_from_module(module = "C:\\vifout\\Putty\\putty.exe")
#app.PuTTYConfiguration.print_control_identifiers()
btnOpen = app.PuTTYConfiguration.child_window(title="Open", auto_id="1009", control_type="Button")
srvOption = app.PuTTYConfiguration.child_window(title="srvvif57", control_type="ListItem")
srvOption.select()
btnOpen.click()
time.sleep(1)
app = Application(backend="uia").connect(process=pid)
#pp.VIF5_7.print_control_identifiers()
pyautogui.typewrite("p02")
pyautogui.press("enter")
time.sleep(1)
pyautogui.press("enter", presses=4)
pyautogui.press('b')
pyautogui.press('e')
time.sleep(1)
"""
if xw.Book("D:\\Sacrif. bovina 21.05.2024.xls").sheets['Date'].range("F7").value == "NU":
    pyautogui.hotkey('ctrl', 'o')
    pyautogui.typewrite("10301")
    pyautogui.press("enter")
    pyautogui.press("enter")
    pyautogui.press("enter")
    pyautogui.typewrite("10301")
    pyautogui.press("enter")
    pyautogui.press("f2")
    time.sleep(1)
    xw.Book("D:\\Sacrif. bovina 21.05.2024.xls").sheets['Date'].range("F7").value = "DA"
if xw.Book("D:\\Sacrif. bovina 21.05.2024.xls").sheets['Date'].range("F9").value == "NU":
    pyautogui.hotkey('ctrl', 'o')
    pyautogui.typewrite("10301")
    pyautogui.press("enter")
    pyautogui.press("enter")
    pyautogui.press("enter")
    pyautogui.typewrite("10301_CAPURI")
    pyautogui.press("enter")
    pyautogui.press("f2")
    time.sleep(1)
    xw.Book("D:\\Sacrif. bovina 21.05.2024.xls").sheets['Date'].range("F9").value = "DA"
"""
for i in range(len(propietar)):
    pyautogui.hotkey('ctrl', 'o')
    pyautogui.typewrite("10301")
    pyautogui.press("enter")
    pyautogui.press("f2")
    pyautogui.press("enter")
    pyautogui.press("f5")
    time.sleep(1)
    for j in range(len(nrCrotal)):
        if nrCrotal[i] != crotaleSortate[j]:
            pydirectinput.press("down")
        else:
            pydirectinput.press("enter")
            crotaleSortate.pop(j)
            pydirectinput.press("f2")
            time.sleep(1)
            break

