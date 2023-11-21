import datetime
import pathlib
import pandas as pd
import numpy as np
import openpyxl
import pyxlsb

save_folder = str(pathlib.Path(pathlib.Path.cwd(), "SAP"))


def sbor(file):
    file = str(save_folder + '\\' + file + '_' + datetime.date.today().strftime("%d.%m.%Y") + '.csv')
    return pd.read_csv(file, sep=';', encoding='cp1251', decimal=',', low_memory=False)


def get_comments():
    print('Чистим вкладку исходник в шаблоне.')
    wb = openpyxl.load_workbook('Shablon.xlsx')
    sheet = wb['исходник']
    sheet.delete_cols(0, 100)
    wb.save('Shablon.xlsx')

    print('Выгружаем таблицы.')
    pursh = sbor('ZLO_PURSCHEDULE')
    pur_report = sbor('ZLO_PUR_REPORT_NEW')
    dv = sbor('ZLO_RESERVE_MANAGER')
    dv = dv.drop_duplicates(['Завод', 'Материал'])
    stock_scrapping = sbor('STOCK_SCRAPPING')
    stock_scrapping['Списания, %'] = np.where(stock_scrapping['REVENUE_RUB_AMT'] == 0, 1, stock_scrapping['stock_scrapping_rub_amt'] / stock_scrapping['REVENUE_RUB_AMT'])
    stock_scrapping['Списания, %'] = np.where(stock_scrapping['Списания, %'] > 1, 1, stock_scrapping['Списания, %'])
    stock_scrapping['Списания, %'] = np.where(pd.isna(stock_scrapping['Списания, %']), 0, stock_scrapping['Списания, %'] * 100)
    pivot_pur_report = pd.pivot_table(pur_report, index=['Завод', '№ товара'],
                                      values=['Количество заказано', 'Количество поставлено'], aggfunc=np.sum)
    pivot_pur_report['SL'] = pivot_pur_report['Количество поставлено'] / pivot_pur_report['Количество заказано']
    pivot_pur_report = pivot_pur_report.reset_index()
    pivot_pur_report['SL'] = pivot_pur_report['SL'] * 100
    shablon = pd.read_excel(io='Shablon.xlsx', engine='openpyxl', sheet_name='шаблон операций', usecols='D:G', header=0)
    shablon = shablon.drop_duplicates(['SAP', 'PLU', 'Наименование товара', 'Проблема'])
    shablon['SAP'] = shablon['SAP'].astype('str')
    shablon['PLU'] = shablon['PLU'].astype('str')
    shablon['Проблема'] = shablon['Проблема'].str.title()
    shablon['SAP'] = shablon['SAP'].str.upper()
    normativ = pd.read_excel(io='Shablon.xlsx', engine='openpyxl', sheet_name='настройки', usecols='B:C', header=1,
                             nrows=1)
    normativ_2 = pd.read_excel(io='Shablon.xlsx', engine='openpyxl', sheet_name='настройки', usecols='B:C', header=14,
                               nrows=30)
    normativ_2 = normativ_2.drop_duplicates(['Категория УИ2 или УИ4', 'Норма списаний(кат)'])
    k_dv = pd.read_excel(io='Shablon.xlsx', engine='openpyxl', sheet_name='настройки', usecols='B:D', header=5, nrows=7)
    exception = pd.read_excel(io='Shablon.xlsx', engine='openpyxl', sheet_name='настройки', usecols='F:H', header=1)
    exception = exception.rename(columns=dict(
        {'Номер PLU/категории УИ2/категория УИ4': 'Материал', 'Проблема': 'Проблема_искл',
         'Комментарий': 'Исключение PLU'}))

    evening_werks = pd.read_excel(
        io='W:\TDM\Автозаказ\ЦЕНТРАЛИЗАЦИЯ\Данные для аналитики\Индикация вечерней поставки.xlsx', engine='openpyxl',
        sheet_name='итог', usecols='B:F', header=0)
    off_cz = pd.read_excel(io='W:\TDM\Автозаказ\ЦЕНТРАЛИЗАЦИЯ\Данные для аналитики\Переносы вправо.xlsx',
                           engine='openpyxl', sheet_name='Лист1', usecols='A:D', header=0)

    off_cz['Выравнивание'] = str('Отключение СЗ и ПЗ было произведено в рамках выравнивания нагрузки на РЦ')
    off_cz = off_cz.rename(columns=dict({'PLU': 'Материал'}))


    print('Собираем исходник для анализа.')
    pursh['Завод'] = pursh['Завод'].astype('str')
    pursh['Материал'] = pursh['Материал'].astype('str')
    base = pd.merge(shablon, pursh, how='left', left_on=['SAP', 'PLU'], right_on=['Завод', 'Материал'])
    base['Дата начала продажи'] = pd.to_datetime(base['Дата начала продажи'], dayfirst= True)
    base['Дата начала продажи'] = base['Дата начала продажи'].dt.date
    base['Дата заказа'] = pd.to_datetime(base['Дата заказа'], dayfirst= True)
    base['Дата заказа'] = base['Дата заказа'].dt.date

    base = base.drop(columns=['Завод', 'Материал', 'КрТекстМатериала'], axis=1)
    stock_scrapping['werks'] = stock_scrapping['werks'].astype('str')
    stock_scrapping['art_id'] = stock_scrapping['art_id'].astype('str')
    base = pd.merge(base, stock_scrapping, how='left', left_on=['SAP', 'PLU'], right_on=['werks', 'art_id'])
    base = base.drop(columns=['werks', 'art_id', 'name_werks', 'plu_name'], axis=1)
    pivot_pur_report['Завод'] = pivot_pur_report['Завод'].astype('str')
    pivot_pur_report['№ товара'] = pivot_pur_report['№ товара'].astype('str')
    base = pd.merge(base, pivot_pur_report, how='left', left_on=['SAP', 'PLU'], right_on=['Завод', '№ товара'])
    base = base.drop(columns=['Завод', '№ товара'], axis=1)
    dv['Завод'] = dv['Завод'].astype('str')
    dv['Материал'] = dv['Материал'].astype('str')
    base = pd.merge(base, dv[['Завод', 'Материал', 'Действ. с', 'Действ. по', 'МинЗаданОбъемЗап', 'Подробный текст']],
                    how='left', left_on=['SAP', 'PLU'], right_on=['Завод', 'Материал'])
    base = base.drop(columns=['Завод', 'Материал'], axis=1)
    evening_werks['Завод'] = evening_werks['Завод'].astype('str')
    base = pd.merge(base, evening_werks[['Завод', 'Индикатор вечерней поставки']], how='left', left_on=['SAP'],
                    right_on=['Завод'])
    base = base.drop(columns=['Завод'], axis=1)
    off_cz['WERKS'] = off_cz['WERKS'].astype('str')
    off_cz['Материал'] = off_cz['Материал'].astype('str')
    base = pd.merge(base, off_cz[['WERKS', 'Материал', 'Выравнивание']], how='left', left_on=['SAP', 'PLU'],
                    right_on=['WERKS', 'Материал'])
    base = base.drop(columns=['WERKS', 'Материал'], axis=1)
    base['Норматив списаний(общ)'] = normativ.at[0, 'Норма списаний']
    base = pd.merge(base, normativ_2[['Категория УИ2 или УИ4', 'Норма списаний(кат)']], how='left',
                    left_on=['Название УИ2'],
                    right_on=['Категория УИ2 или УИ4'])
    base = base.drop(columns=['Категория УИ2 или УИ4'], axis=1)
    normativ_2 = normativ_2.rename(columns=dict({'Норма списаний(кат)': 'Норма списаний(УИ4)'}))
    base = pd.merge(base, normativ_2[['Категория УИ2 или УИ4', 'Норма списаний(УИ4)']], how='left',
                    left_on=['Название УИ4'],
                    right_on=['Категория УИ2 или УИ4'])
    base = base.drop(columns=['Категория УИ2 или УИ4'], axis=1)
    base['Норматив списаний'] = np.where(pd.isna(base['Норма списаний(кат)']),
                                         np.where(pd.isna(base['Норма списаний(УИ4)']), base['Норматив списаний(общ)'],
                                                  base['Норма списаний(УИ4)']), base['Норма списаний(кат)'])
    base['Норматив списаний'] = base['Норматив списаний'] * 100
    base['Норматив SL'] = normativ.at[0, 'Норма SL']
    base['Норматив SL'] = base['Норматив SL'] * 100
    exception['Материал'] = exception['Материал'].astype('str')
    base = pd.merge(base, exception[['Материал', 'Проблема_искл', 'Исключение PLU']], how='left',
                    left_on=['PLU', 'Проблема'],
                    right_on=['Материал', 'Проблема_искл'])
    base = base.drop(columns=['Материал', 'Проблема_искл'], axis=1)
    exception = exception.rename(columns=dict({'Материал': 'УИ2', 'Исключение PLU': 'Исключение УИ2'}))
    base = pd.merge(base, exception[['УИ2', 'Проблема_искл', 'Исключение УИ2']], how='left',
                    left_on=['Название УИ2', 'Проблема'],
                    right_on=['УИ2', 'Проблема_искл'])
    base = base.drop(columns=['УИ2', 'Проблема_искл'], axis=1)
    exception = exception.rename(columns=dict({'УИ2': 'УИ4', 'Исключение УИ2': 'Исключение УИ4'}))
    base = pd.merge(base, exception[['УИ4', 'Проблема_искл', 'Исключение УИ4']], how='left',
                    left_on=['Название УИ4', 'Проблема'],
                    right_on=['УИ4', 'Проблема_искл'])
    base = base.drop(columns=['УИ4', 'Проблема_искл'], axis=1)
    base['Итоговые исключения'] = np.where(pd.isna(base['Исключение PLU']), np.where(pd.isna(base['Исключение УИ2']),
                                                                                     np.where(pd.isna(
                                                                                         base['Исключение УИ4']), 0,
                                                                                         base[
                                                                                             'Исключение УИ4'].astype(
                                                                                             'str')),
                                                                                     base['Исключение УИ2'].astype(
                                                                                         'str')),
                                           base['Исключение PLU'].astype('str'))
    base['Коэффициент для ДВ'] = np.where(base['Срок годности'] > 16, k_dv.at[4, 'значение'],
                                          np.where(base['Срок годности'] > 11, k_dv.at[3, 'значение'],
                                                   np.where(base['Срок годности'] > 6, k_dv.at[2, 'значение'],
                                                            np.where(base['Срок годности'] > 3, k_dv.at[1, 'значение'],
                                                                     k_dv.at[0, 'значение']))))
    base['Минимальная ДВ'] = k_dv.at[6, 'до']

    # TODO Cделать проверки на ПРОМО по виду РкМР.
    # TODO По итогу должны получить файл ИСХОДНИК и файл ИТОГ, в первом будут все расчеты, а во втором уже готовый файл;

    print('Проводим анализ.')
    base['Наименование УпрКластера'] = np.where(pd.isna(base['Наименование УпрКластера']), 'Некорректно заполнен шаблон', base['Наименование УпрКластера'].astype('str'))
    base['Наименование филиала ЦФО'] = np.where(pd.isna(base['Наименование филиала ЦФО']), 'Некорректно заполнен шаблон', base['Наименование филиала ЦФО'].astype('str'))
    base['Проверка поля проблема'] = [0 if x in ['Дефицит', 'Перетарка'] else 'Некорректно заполнено поле проблема'
                                      for x in base['Проблема']]
    base['Наличие в АМ'] = [
        0 if x in ['КРС', 'ЖЛТ', 'ЗЛН'] else 'Нет в матрице магазина или некорректный номер магазина/позиции' for x in
        base['Ид.']]
    base['Проверка блокировки к АЗ'] = np.where(base['Ид.'] == 'КРС',
                                                "Блок к АЗ: " + np.where(pd.isna(base['Приоритет ошибок']), '',
                                                                         base['Приоритет ошибок'].astype(
                                                                             'str')) + ' ' + np.where(
                                                    pd.isna(base['Расшифровка причины блокировки товара']), '',
                                                    base['Расшифровка причины блокировки товара'].astype('str')), 0)

    base['Проверка наличия ДВ'] = np.where((base['Проблема'] == 'Перетарка') & (base['МинЗаданОбъемЗап'] > 0),
                                           "Прогружена ДВ: " + base['Подробный текст'].astype('str'), 0)
    base['Проверка списаний'] = np.where(base['Списания, %'] > base['Норматив списаний'],
                                         "Потери за последние 7 дней " + np.around(
                                             base['Списания, %'].astype(float), 2).astype('str') + '%', 0)
    base['Проверка SL'] = np.where(base['SL'] < base['Норматив SL'],
                                   "Недопоставка товара, SL последних 7 дней: " + np.around(
                                       base['SL'].astype(float), 2).astype('str') + '%', 0)
    base['Проверка движения товара'] = np.where((base['СвобИспользЗапас'] > 0) & (base['Базисное значение'] == 0),
                                                "Есть остаток - продаж нет, товар без движения; необходимо провести ЛИ",
                                                0)
    base['Дней для продажи кванта'] = np.where(base['Базисное значение'] != 0,
                                               base['Квант'] / base['Базисное значение'], base['Срок годности'] + 1)
    base['Проверка оборачиваемости кванта'] = np.where((base['Квант'] > 1) & (base['Срок годности'] > 1) &
        (base['Дней для продажи кванта'] > base['Срок годности']) & (base['СвобИспользЗапас'] <= base['Квант']),
        "Остаток " + base['СвобИспользЗапас'].astype('str') + " не более одного кванта (" + base['Квант'].astype(
            'str') + ") - обратитесь к КМ, для снижения кванта/вывода из АМ", 0)
    base['Проверка ТЗ'] = np.where(
        (base['СвобИспользЗапас'] - base['Базисное значение'] * base['НОЗ(расчет)'] - base['Уровень нуля']) > 0, 1, 0)
    base['Проверка ТЗ в норме'] = np.where(((base['Проблема'] == 'Дефицит') & (base['Проверка ТЗ'] == 1) & (base['Срок годности'] > 1)) | (
            (base['Проблема'] == 'Перетарка') & (base['Проверка ТЗ'] == 0) & (base['Проверка списаний'] == 0)& (base['Срок годности'] > 1)),
                                           np.where(base['НОЗ(расчет)'] == 1,
                                                    'ТЗ в норме, остаток ' + base['СвобИспользЗапас'].astype(
                                                        'str') + ' кг/шт при средних продажах ' + base[
                                                        'Базисное значение'].astype(
                                                        'str') + ' в день и ежедневных поставках',
                                                    'ТЗ в норме, остаток ' + base['СвобИспользЗапас'].astype(
                                                        'str') + ' кг/шт при средних продажах ' + base[
                                                        'Базисное значение'].astype(
                                                        'str') + ' в день и и перерывом между поставками в ' + base[
                                                        'НОЗ(расчет)'].astype('str') + ' дн.'), 0)
    base['Объем ДВ'] = np.where(
        (np.around(base['Базисное значение'].astype(float), 2) * np.where((base['Индикатор вечерней поставки'] == 1) & (base['Проверка ТЗ в норме'] == 0), 0.5,
                                                                          base['Коэффициент для ДВ'])) < base[
            'Минимальная ДВ'],
        base['Минимальная ДВ'],
        np.around(base['Базисное значение'].astype(float), 2) * np.where(base['Индикатор вечерней поставки'] == 1, 0.5,
                                                                         base['Коэффициент для ДВ']))
    base['Объем ДВ'] = np.around(base['Объем ДВ'].astype(float), 0)
    base['Комментарий для вечерней поставки'] = np.where(base['Индикатор вечерней поставки'] == 1,
                                                         'Для ТТ с вечерней поставкой прогружена ДВ на 14 дней',0)
    base['Проверка отключенного СЗ'] = np.where((pd.notna(base['Действ. с_x'])) & (pd.isna(base['Страх. запас'])),
                                                'Пополняется только под продажи, позиция направлена в НПС', 0)
    """
    base['ПлановСрокПоставки'] = np.where(pd.isna(base['ПлановСрокПоставки']), 0, base['ПлановСрокПоставки'])
    base['Проверка даты ПРОМО'] = np.where((base['Дата начала продажи'] - datetime.datetime.now().date()).dt.days <= (base['ПлановСрокПоставки'] + 1), 1, 0)
    base['Проверка даты ПРОМО'] = np.where(pd.isna(base['Дата начала продажи']), 0, base['Проверка даты ПРОМО'])
    base['Проверка ПРОМО'] = np.where(base['Проверка даты ПРОМО'] ==0, 0, 'Прогружено ПРОМО  - обратитесь к КМ через своего ДК: ' + base['Название'].astype(str))
    base['Дата поставки с учетом нового заказа'] = np.where(pd.notna(base['Дата заказа']), np.where(base['Дата заказа'] == datetime.datetime.now().date(), base['Дата заказа'] + base['НОЗ(расчет)'] + base['ПлановСрокПоставки'], base['Дата заказа'] + base['ПлановСрокПоставки'], 0),0)
    base['Проверка ПРОМО'] = np.where((base['Проблема'] == 'Перетарка') & ((base['Название'].str.contains('ОСГ')).any()), 0, base['Проверка ПРОМО'])

    ru_alphabet = {'а', 'б', 'в', 'г', 'д', 'е', 'ё', 'ж', 'з', 'и', 'й', 'к', 'л', 'м', 'н', 'о',
                   'п', 'р', 'с', 'т', 'у', 'ф', 'х', 'ц', 'ч', 'ш', 'щ', 'ъ', 'ы', 'ь', 'э', 'ю', 'я'}
    symbols = {'.', ',', ':', ';'}
    werks = base['SAP'].tolist()
    werks_new = ['0000']
    for werk in werks:
        if len(str(werk)) == 4 or bool(ru_alphabet.intersection(set(str(werk).lower()))) is False or \
                bool(symbols.intersection(set(str(werk)))) is False:
            werks_new.append(str(werk))
    werks = set(werks_new)
    """

    print('Добавляем исходник в шаблон.')

    with pd.ExcelWriter('Shablon.xlsx', mode='a', engine="openpyxl", if_sheet_exists='replace') as writer:
        base.to_excel(writer, sheet_name='исходник', index=True)

    wb = openpyxl.load_workbook('Shablon.xlsx')
    sheet = wb['исходник']
    sheet['A1'] = 'cц'
    sheet['A2'] = '=B:B&C:C'

    wb.save('Shablon.xlsx')
