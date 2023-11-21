import datetime
import pathlib
import pandas as pd
import pyodbc
from sqlalchemy.engine import URL
from sqlalchemy import create_engine
import openpyxl

ru_alphabet = {'а', 'б', 'в', 'г', 'д', 'е', 'ё', 'ж', 'з', 'и', 'й', 'к', 'л', 'м', 'н', 'о',
               'п', 'р', 'с', 'т', 'у', 'ф', 'х', 'ц', 'ч', 'ш', 'щ', 'ъ', 'ы', 'ь', 'э', 'ю', 'я'}

eng_alphabet = {'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p',
                'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z'}

symbols = {'.', ',', ':', ';'}

SOURCE_PATH = pathlib.Path(pathlib.Path.cwd(), "Shablon.xlsx")
save_folder = str(pathlib.Path(pathlib.Path.cwd(), "SAP"))


# запись сцепок в links.txt
def to_file():
    print("Забрал сцепки из шаблона.")

    links = pd.read_excel(io=SOURCE_PATH, engine='openpyxl', sheet_name='шаблон операций', usecols='D:E', header=0)
    links['сц'] = links['SAP'].astype('str') + links['PLU'].astype('str')
    links = links.drop(columns=['SAP', 'PLU'], axis=1)
    links.sort_values(by='сц')
    links.to_csv('links.txt', index=False)


def get_stock_scrapping():
    df = pd.read_excel(io=SOURCE_PATH, engine='openpyxl', sheet_name='шаблон операций', usecols='D:E', header=0)

    werks = df['SAP'].tolist()

    werks_new = ['0000']
    for werk in werks:
        if len(str(werk)) == 4 or bool(ru_alphabet.intersection(set(str(werk).lower()))) is False or \
                bool(symbols.intersection(set(str(werk)))) is False:
            werks_new.append(str(werk))
    werks = tuple(set(werks_new))

    goods = df['PLU'].tolist()
    goods_new = [0]
    for plu in goods:
        if len(str(plu)) < 10 and isinstance(plu, int):
            goods_new.append(int(plu))
    goods = tuple(set(goods_new))


    connection_string = "DRIVER={SQL Server};SERVER=********;DATABASE=Data;Trusted_Connection=yes"
    connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": connection_string})
    engine = create_engine(connection_url)

    print('Соединение с сервером ******** установлено!')


    query = """
    SOME SQL Scripts
    """.format(werks, goods)


    print('Выгружаю списания...')
    stock_scrapping = pd.read_sql_query(query, engine)
    file = str(save_folder + '\\' + 'STOCK_SCRAPPING' + '_' + datetime.date.today().strftime("%d.%m.%Y") + '.csv')
    stock_scrapping.to_csv(file, sep=';', encoding='windows-1251', index=False, decimal=',')

    engine.dispose()
