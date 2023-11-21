import sap
import os
from threading import Thread


def run(tranz_name, file_name, path_save):

    session = sap.connect()

    session.findById("wnd[0]/tbar[0]/okcd").text = tranz_name
    session.findById("wnd[0]").sendVKey(0)

    session.findById("wnd[0]/usr/radP_FREAD").select()
    session.findById("wnd[0]/usr/ctxtP_FNAME").text = file_name
    if tranz_name == 'ZLO_PURSCHEDULE':
        session.findById("wnd[0]/usr/chkX_FILTER").selected = False

    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    # сообщение о коммерческой тайне - ok
    session.findById("wnd[0]").sendVKey(0)
    #pyautogui.keyDown('ENTER') - альтернатива функции выше, необходим пакет pyautogui

    # формат
    if tranz_name == 'ZLO_PURSCHEDULE' or tranz_name == 'ZLO_PUR_REPORT_NEW':
        session.findById("wnd[0]/tbar[1]/btn[33]").press()
        session.findById("wnd[1]/usr/cntlGRID/shellcont/shell").selectColumn("VARIANT")
        session.findById("wnd[1]/usr/cntlGRID/shellcont/shell").selectedRows = ""
        session.findById("wnd[1]/usr/cntlGRID/shellcont/shell").contextMenu()
        session.findById("wnd[1]/usr/cntlGRID/shellcont/shell").selectContextMenuItem("&FILTER")
        session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "/AVTO_FORMAT"
        session.findById("wnd[2]").sendVKey(0)
        session.findById("wnd[1]/usr/cntlGRID/shellcont/shell").selectedRows = "0"
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

    # сохранение
    path = path_save + '\\' + file_name + ".csv"

    if os.path.exists(path):
        os.remove(path)
    t = Thread(target=sap.save, args=(path,))
    t.start()


    if tranz_name == 'ZLO_PURSCHEDULE':
        session.findById("wnd[0]/tbar[1]/btn[19]").press()  # Shift-F7
    elif tranz_name == 'ZLO_RESERVE_MANAGER':
        session.findById("wnd[0]/usr/cntlCONT1/shellcont/shell").selectColumn("WRPRWUG")
        session.findById("wnd[0]/usr/cntlCONT1/shellcont/shell").pressToolbarButton("&SORT_DSC")
        session.findById("wnd[0]/usr/cntlCONT1/shellcont/shell").pressToolbarButton("ZCSV")
    else:
        session.findById("wnd[0]/tbar[1]/btn[7]").press()  # F7


    t.join()

    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    session.findById("wnd[0]/tbar[0]/btn[15]").press()


    os.system('Taskkill /IM excel.exe /t /f')

