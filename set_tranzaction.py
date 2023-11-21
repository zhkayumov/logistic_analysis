import sap
import datetime
import subprocess

def copy_to_clip(file_name):
    cmd = "clip < " + file_name
    w = subprocess.check_call(cmd, shell=True)
    return w


def run(tranz_name):


    session = sap.connect()

    print('Запускаю фон в транзакции: '+ str(tranz_name))
    session.findById("wnd[0]/tbar[0]/okcd").text = tranz_name
    session.findById("wnd[0]/tbar[0]/btn[0]").press()

    session.findById("wnd[0]/usr/ctxtP_FNAME").text = tranz_name + '_' + datetime.date.today().strftime("%d.%m.%Y")

    copy_to_clip("links.txt")
    session.findById("wnd[0]/usr/btn%_S_WM_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    if tranz_name == 'ZLO_PUR_REPORT_NEW':

        session.findById("wnd[0]/usr/btn%_SO_BSART_%_APP_%-VALU_PUSH").press()
        list_dok_zakupki = ['AZUB', 'AZNB', 'AZCC', 'CDUB', 'CDCC', 'ZUB', 'ZNB', 'ZCC']

        i = 0
        for elem in list_dok_zakupki:
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL-SLOW_I[1," + str(i) + "]").text = elem
            i = i + 1

        session.findById("wnd[1]/tbar[0]/btn[8]").press()

        session.findById("wnd[0]/usr/chkP_DEL").selected = True

        session.findById("wnd[0]/usr/btn%_SO_FRGKE_%_APP_%-VALU_PUSH").press()
        id_zakaza = ['Z', 'R']

        i = 0
        for elem in id_zakaza:
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL-SLOW_I[1," + str(i) + "]").text = elem
            i = i + 1

        session.findById("wnd[1]/tbar[0]/btn[8]").press()



        seven_days_ago = datetime.datetime.now() - datetime.timedelta(days=7)
        one_day_ago = datetime.datetime.now() - datetime.timedelta(days=1)

        # Дата заказа
        session.findById("wnd[0]/usr/ctxtSO_EBDAT-LOW").text = ""
        session.findById("wnd[0]/usr/ctxtSO_EBDAT-HIGH").text = ""
        # Дата поставки (план)
        session.findById("wnd[0]/usr/ctxtSO_EINDT-LOW").text = seven_days_ago.strftime("%d.%m.%Y")
        session.findById("wnd[0]/usr/ctxtSO_EINDT-HIGH").text = one_day_ago.strftime("%d.%m.%Y")
    # else:
    #     session.findById("wnd[0]/usr/radP_FGETA").select()
    #     #session.findById("wnd[0]/usr/ctxtP_VARI").text = "AVTO_FORMAT"

    if tranz_name == 'ZLO_RESERVE_MANAGER':

        three_day_ago = datetime.datetime.now() + datetime.timedelta(days=3)
        session.findById("wnd[0]/usr/ctxtS_DATUM-HIGH").text = three_day_ago.strftime("%d.%m.%Y")

        session.findById("wnd[0]/usr/btn%_S_ZEISB_%_APP_%-VALU_PUSH").press()
        list_Z = ['Z004', 'Z006', 'Z007', 'Z008']

        i = 0
        for elem in list_Z:
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL-SLOW_I[1," + str(i) + "]").text = elem
            i = i + 1

        session.findById("wnd[1]/tbar[0]/btn[8]").press()


    # далее ставим задания
    session.findById("wnd[0]").sendVKey(9)

    session.findById("wnd[1]/tbar[0]/btn[13]").press()

    # немедленно
    session.findById("wnd[1]/usr/btnSOFORT_PUSH").press()
    # конкретные дата и время
    # session.findById("wnd[1]/usr/btnDATE_PUSH").press()
    # session.findById("wnd[1]/usr/ctxtBTCH1010-SDLSTRTDT").text = datetime.date.today().strftime("%d.%m.%Y")
    #
    # # "hh:mm:ss" + 5 мин
    # nexttime = datetime.datetime.now() + datetime.timedelta(minutes=5)
    # session.findById("wnd[1]/usr/ctxtBTCH1010-SDLSTRTTM").text = nexttime.ctime().split()[3]
    #
    # session.findById("wnd[1]/usr/ctxtBTCH1010-SDLSTRTTM").setFocus
    # session.findById("wnd[1]/usr/ctxtBTCH1010-SDLSTRTTM").caretPosition = 8
    session.findById("wnd[1]/tbar[0]/btn[11]").press()

    session.findById("wnd[0]/tbar[0]/btn[15]").press()






