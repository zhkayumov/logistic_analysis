import sap
import get_from_tranzaction
import os
import datetime
import pathlib
import analysis

save_folder = str(pathlib.Path(pathlib.Path.cwd(), "SAP"))
tranz_list = ['ZLO_PURSCHEDULE', 'ZLO_PUR_REPORT_NEW', 'ZLO_RESERVE_MANAGER']

add_links.to_file()

add_links.get_stock_scrapping()

sap.run()
for elem in tranz_list:
    print('Выгружаю.' + elem)
    get_from_tranzaction.run(elem, elem + '_' + datetime.date.today().strftime("%d.%m.%Y"), save_folder)
sap.close()

analysis.get_comments()

print('Завершено.')

os._exit(0)
