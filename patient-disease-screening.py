import pandas as pd
import re
import xlsxwriter
import telebot
import getpass
import platform
import time
import threading
from datetime import datetime, timedelta
from sqlalchemy import create_engine
import requests
import numpy as np
import os
from dotenv import load_dotenv

# Паттерны
load_dotenv()
engine = create_engine(os.getenv("DB_URL"))
bot = telebot.TeleBot(os.getenv("TELEGRAM_TOKEN"))
chat_id = os.getenv("CHAT_ID")
re_1 = r'[^0-9,.;/]' # Регулярное выражение для отсева букв и знаков
pd.set_option('display.max_columns', None)
pd.set_option('mode.chained_assignment', None)
pd.options.mode.chained_assignment = None
lock = threading.Lock() # Локер процессов
# Создание коннекторов данных
start_time = time.time() # Текущее время
datename = datetime.now().strftime('%d.%m %H.%M.%S') # Время создания файла
chid_list = ["Конс каб 72", "Конс каб 77", "Операционная № 1", "Операционная № 2", "Операционная № 3", "Опер 1.1", "Опер 1.2", "Опер 2.1", "Опер 2.2", "Опер 3.2"] # В каких кабинетах могут мыть операции

def SendTelegram(status, er):
    UserName = getpass.getuser()                                                                                                                                                                         # Имя пользователя (обычно оно User - не информативно)
    CompName = platform.node()                                                                                                                                                                           # Имя компьютера                                                                                                                                                                              # ID моей телеги
    if status == "try": # Если связь с телегой установлена
        bot.send_message(chat_id, datename+" пользователь "+UserName+" ("+CompName+") успешно воспользовался скриптом для подсчёта клиентов с болезнями")                                                # Отправка сообщения
    elif status == "except1": # Если нет подключения к SQL серверу
        bot.send_message(chat_id, "ERROR: " + datename+" пользователь "+UserName+" ("+CompName+") неудачно запустил скрипт для подсчёта клиентов с болезнями: " + er)                                                  # Отправка сообщения
def Input_lag():
    global start_date, end_date
    lock.acquire()
    while True:
        try:
            start_date = input("Введите дату начала в формате dd.mm.yy:\n")
            start_date = pd.to_datetime(datetime.strptime(start_date, '%d.%m.%y'))
            print("Начальная дата: " + str(start_date)); break
        except: print("Введённый параметр не соответствует допустимому формату даты - попробуйте написать дату по другому")
    while True:
        try:
            end_date = input("Теперь введите конечную дату (обратите внимание, что по умолчанию дате присвоится время 00:00:00, т.е. если вы хотите посчитать, например, до 10.10.24 ВКЛЮЧИТЕЛЬНО, то следует написать 11.10.24):\n")
            if end_date == "": end_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0); print(end_date)
            else: end_date = pd.to_datetime(datetime.strptime(end_date, '%d.%m.%y'))
            if end_date < start_date: print("Конечная дата меньше начальной, так быть не должно, попробуйте снова"); continue
            print("Конечная дата: " + str(end_date)); break
        except: print("Введённый параметр не соответствует допустимому формату даты - попробуйте написать дату по другому")
    print("Все опциональные параментры введены пользователем. Ожидаем дозагрузки данных и начала расчёта...")
    lock.release()
def GetSQL():
	global doctor_list, start_date, end_date, end_date_plan, start_time
	# SQL запросы
	lightquery_SCHEDULE = "SELECT schedid, dcode, dcode1, workdate, bhour, bmin, fhour, fmin, depnum, pcode, chid, status, clvisit, COMMENT from SCHEDULE WHERE pcode IS NOT NULL" # Расписание
	lightquery_BI_DOCTORS = "SELECT dcode, dname, depnum, lockdate from BI_DOCTORS" # Доктора WHERE lockdate IS NULL - без уволенных
	lightquery_SHEDMARKS = "SELECT mrkid, mrktext FROM SHEDMARKS" # Маркеры расписания (о/а, м/а)
	lightquery_CLIENTS = "SELECT pcode, fullname, COMMENT FROM CLIENTS" # Клиенты
	lightquery_DEPARTMENTS = "SELECT depnum, depname FROM DEPARTMENTS" # Список отделений
	lightquery_CHAIRS = "SELECT chid, chname FROM CHAIRS" # Рабочие места (кресла)
	# Получение данных
	SCHEDULE = pd.read_sql(lightquery_SCHEDULE, engine) # Расписание
	BI_DOCTORS = pd.read_sql(lightquery_BI_DOCTORS, engine) # dict_BI_DOCTORS_name, dict_BI_DOCTORS_dep
	SHEDMARKS = pd.read_sql(lightquery_SHEDMARKS, engine) # SHEDMARKS
	CLIENTS = pd.read_sql(lightquery_CLIENTS, engine) # dict_CLIENTS
	DEPARTMENTS = pd.read_sql(lightquery_DEPARTMENTS, engine) # dict_DEPARTMENTS
	CHAIRS = pd.read_sql(lightquery_CHAIRS, engine)  # dict_CHAIRS
	print(f"Все базы загружены в буфер за :{(time.time() - start_time):.2f}"); start_time = time.time()
	# Формирование списка докторов
	doctor_list = BI_DOCTORS[(BI_DOCTORS['depnum'] == 120363)] # Создаём список действующих врачей хирургов
	doctor_list = list(filter(None, doctor_list['dname'].tolist()))
	lock.acquire()
	# Обрезка расписания по датам
	SCHEDULE = SCHEDULE[((SCHEDULE['workdate'] >= start_date) & (SCHEDULE['workdate'] <= end_date))]
	print(SCHEDULE['workdate'].min(), SCHEDULE['workdate'].max(), SCHEDULE['workdate'].count())
	SCHEDULE_clients_list = SCHEDULE['pcode'].tolist();	print("словарь собран из " + str(len(SCHEDULE_clients_list)) + " клиентов")
	# Оформление словарей
	dict_DEPARTMENTS = dict(DEPARTMENTS[['depnum', 'depname']].values) # Словарь соответствий ID отделений и названий отделений
	dict_BI_DOCTORS_name = dict(BI_DOCTORS[['dcode', 'dname']].values) # Словарь соответствий ID сотрудника и полное имя
	dict_BI_DOCTORS_dep = dict(BI_DOCTORS[['dcode', 'depnum']].values) # Словарь соответствий ID сотрудника и отделение
	dict_SHEDMARKS = dict(SHEDMARKS[['mrkid', 'mrktext']].values) # Словарь соответствий ID марок расписания и наименований (о/а, м/а)
	dict_CLIENTS = dict(CLIENTS[['pcode', 'fullname']].values) # Словарь соответствий ID клиента и имена клиентов
	dict_CLIENTS_COMMENT = dict(CLIENTS[['pcode', 'COMMENT']].values) # Словарь соответствий ID клиента и имена клиентов
	dict_CHAIRS = dict(CHAIRS[['chid', 'chname']].values)  # Словарь соответствий ID рабочего места и его названия
	# Произведение замен по словарям и предрасчёт необходимых показателей
	SCHEDULE['dcode'] = SCHEDULE['dcode'].replace(-1, np.nan)
	SCHEDULE['Доктор'] = SCHEDULE[['dcode', 'dcode1']].bfill(axis=1).iloc[:, 0] # Объединение двух полей с кодом доктора
	SCHEDULE['Рабочее место'] = SCHEDULE['chid'].map(dict_CHAIRS)  # Присвоение рабочего места по ID
	SCHEDULE['Отделение доктора'] = SCHEDULE['Доктор'].map(dict_BI_DOCTORS_dep) # Отдельная колонка с присвоение доктору кода отделения
	SCHEDULE['Отделение'] = SCHEDULE['depnum'].fillna(SCHEDULE['Отделение доктора']) # Присвоить записи отделение доктора, если не указано в явном виде в расписании
	SCHEDULE['Отделение'] = SCHEDULE['Отделение'].map(dict_DEPARTMENTS) # Присвоить название отделения по коду
	SCHEDULE['Продолжительность процедуры, минуты'] = (SCHEDULE['fhour'] - SCHEDULE['bhour'])*60 + (SCHEDULE['fmin'] - SCHEDULE['bmin']) # Расчёт продолжительности окна приёма в расписании
	SCHEDULE['Имя доктора'] = SCHEDULE['Доктор'].map(dict_BI_DOCTORS_name) # Присвоение имени доктора по ID
	SCHEDULE['Имя пациента'] = SCHEDULE['pcode'].map(dict_CLIENTS) # Присвоение имён пациентам по их ID
	SCHEDULE['Примечание в карточке пациента'] = SCHEDULE['pcode'].map(dict_CLIENTS_COMMENT) # Присвоение примечаний в карточке пациентов по их ID
	SCHEDULE['Статус'] = SCHEDULE['status'].map(dict_SHEDMARKS) # Присвоение статуса по ID
	SCHEDULE = SCHEDULE.sort_values(by=['workdate'], ascending=True) # Сортировка по дате назначения
	# Присвеоение имён
	SCHEDULE = SCHEDULE[['schedid', 'pcode', 'workdate', 'Продолжительность процедуры, минуты', 'clvisit', 'Отделение', 'Статус', 'Рабочее место', 'Имя пациента', 'Имя доктора', 'Примечание в карточке пациента', 'COMMENT']]
	SCHEDULE.columns = ['Код записи в расписании', 'Код пациента', 'Дата назначения', 'Продолжительность процедуры, минуты', 'Посещение', 'Отделение', 'Тип операции', 'Рабочее место', 'Имя пациента', 'Имя доктора', 'Примечание в карточке пациента', 'Примечание в статусе визита']
	OPERATIONS_BASE(SCHEDULE)
def classify_disease(notes_card, notes_visit):
	notes = str(notes_card).lower() + " " + str(notes_visit).lower()
	# Ключевые слова для каждой группы заболеваний
	hepatitis_keywords_re_B = [re.compile(r'\b' + re.escape(k) + r'\b') for k in ['геп в', 'геп b', 'геп б']]
	hepatitis_keywords_re_C = [re.compile(r'\b' + re.escape(k) + r'\b') for k in ['гепатит', 'геп c', 'геп с']]
	syphilis_keywords_re = [re.compile(r'\b' + re.escape(k) + r'\b') for k in ['сифилис', 'сифилису', 'rw']]
	hiv_aids_keywords_re = [re.compile(r'\b' + re.escape(k) + r'\b') for k in ['вич', 'спид']]
	# Проверка на наличие заболеваний
	for keyword_re in hepatitis_keywords_re_B:
		if keyword_re.search(notes): return 'Гепатит B'
	for keyword_re in hepatitis_keywords_re_C:
		if keyword_re.search(notes): return 'Гепатит C'
	for keyword_re in syphilis_keywords_re:
		if keyword_re.search(notes): return 'Сифилис'
	for keyword_re in hiv_aids_keywords_re:
		if keyword_re.search(notes): return 'ВИЧ (СПИД)'
	return '' # Если заболевание не найдено
def check_treatment_status(notes_card, notes_visit):
	notes = str(notes_card).lower() + " " + str(notes_visit).lower()
	treatment_keywords = [re.compile(r'\b' + re.escape(k) + r'\b') for k in ["пролечен", "антитела", "вылечен", "пролеченный", "полож", "пол", "проле", "снята", "снят"]]
	for keyword_re in treatment_keywords:
		if keyword_re.search(notes): return 'лечение'
	return ''
def create_summary_pivot(df):
	df['Дата назначения'] = pd.to_datetime(df['Дата назначения'])
	df['Месяц_Год'] = df['Дата назначения'].dt.strftime('%m.%y')
	df['Сортировка_Месяц_Год'] = df['Дата назначения'].dt.strftime('%Y%m').astype(int)
	ordered_months = df.sort_values(by='Сортировка_Месяц_Год')['Месяц_Год'].unique().tolist()
	df['Месяц_Год'] = pd.Categorical(df['Месяц_Год'], categories=ordered_months, ordered=True)
	pivot_table = pd.pivot_table(df, values='Код пациента', index='Группа заболеваний', columns='Месяц_Год', aggfunc='count', fill_value=0)
	pivot_table.loc['ИТОГО'] = pivot_table.sum()
	return pivot_table
def OPERATIONS_BASE(base):
	global operatios_base, start_time
	operatios_base = base[(base['Посещение'] == 1) & (base['Дата назначения'] <= end_date) & (base['Дата назначения'] >= start_date) & (base['Рабочее место'].isin(list(filter(None, chid_list)))) & (base['Продолжительность процедуры, минуты'] > 40) & (base['Имя доктора'].isin(list(filter(None, doctor_list)))) & ((base['Тип операции'] == 'о/а операция') | (base['Тип операции'] == 'м/а операция')) & ((base['Отделение'] == 'Хирургия стационарная') | (base['Отделение'] == 'Хирургия амбулаторная') | (base['Отделение'] == 'Хирургия амбулаторная,Хирургия стационарная'))]
	# Добавление новой колонки "Группа заболеваний"
	operatios_base['Группа заболеваний'] = operatios_base.apply(lambda row: classify_disease(row['Примечание в карточке пациента'], row['Примечание в статусе визита']), axis=1)
	table_base = operatios_base[['Код записи в расписании', 'Дата назначения', 'Имя пациента', 'Тип операции', 'Группа заболеваний', 'Примечание в карточке пациента', 'Примечание в статусе визита']]
	only_dis = operatios_base[operatios_base['Группа заболеваний'] != ''].copy()
	only_dis['Статус'] = only_dis.apply(lambda row: check_treatment_status(row['Примечание в карточке пациента'], row['Примечание в статусе визита']), axis=1)
	only_dis = only_dis[['Код записи в расписании', 'Дата назначения', 'Имя пациента', 'Тип операции', 'Группа заболеваний', 'Статус', 'Примечание в карточке пациента', 'Примечание в статусе визита']]
	
	with pd.ExcelWriter("Клиенты с болезнями " + datename + '.xlsx', engine="xlsxwriter") as writer:
		table_base.to_excel(writer, sheet_name='Список всех операций', index=False, freeze_panes=(1, 0))
		only_dis.to_excel(writer, sheet_name='Пациенты с заболеваниями', index=False, freeze_panes=(1, 0))
		
		workbook = writer.book  # Доступ к объектам (форматам, диаграммам xlsxwriter)
		# Все операции
		sheet = writer.sheets["Список всех операций"]  # Выбор активного листа
		sheet.set_column(0, 0, 10); sheet.set_column(1, 1, 20);	sheet.set_column(2, 2, 35); sheet.set_column(3, 3, 15); sheet.set_column(4, 4, 20); sheet.set_column(5, 5, 50); sheet.set_column(6, 6, 50)
		sheet.autofilter(0, 0, len(table_base), len(table_base.columns) - 1)
		red_format = workbook.add_format({'bg_color': '#FFC7CE'})
		sheet.conditional_format(1, 0, len(table_base), len(table_base.columns) - 1, {'type': 'formula', 'criteria': '=$E2<>""', 'format': red_format})
		
		# Пациенты с заболеваниями
		sheet = writer.sheets["Пациенты с заболеваниями"]  # Выбор активного листа
		sheet.set_column(0, 0, 10); sheet.set_column(1, 1, 20);	sheet.set_column(2, 2, 35); sheet.set_column(3, 3, 15); sheet.set_column(4, 4, 20); sheet.set_column(5, 5, 10); sheet.set_column(6, 6, 50); sheet.set_column(7, 7, 50)
		sheet.autofilter(0, 0, len(only_dis), len(only_dis.columns) - 1)
		red_light_format = workbook.add_format({'bg_color': '#FAF1F0'})
		sheet.conditional_format(1, 0, len(only_dis), len(only_dis.columns) - 1, {'type': 'formula', 'criteria': '=$F2<>""', 'format': red_light_format})
		sheet.conditional_format(1, 0, len(only_dis), len(only_dis.columns) - 1, {'type': 'formula', 'criteria': '=AND($E2<>"", $F2="")', 'format': red_format})
		
		# Сводная таблица делается из полной выборки с группами заболеваний
		if not operatios_base.empty:
			summary_pivot_df = operatios_base[operatios_base['Группа заболеваний'] != ''].copy()
			summary_pivot = create_summary_pivot(summary_pivot_df)
			if not summary_pivot.empty:
				summary_pivot.to_excel(writer, sheet_name='Сводная по заболеваниям, мес', index=True, header=True)
				sheet_summary = writer.sheets["Сводная по заболеваниям, мес"]
				border_format = workbook.add_format({'border': 1})
				num_rows = len(summary_pivot) + 1  # +1 для строки заголовка
				num_cols = len(summary_pivot.columns) + 1  # +1 для индексного столбца
				sheet_summary.conditional_format(0, 0, num_rows - 1, num_cols - 1, {'type': 'no_blanks', 'format': border_format})
				sheet_summary.freeze_panes(1, 1)
				for row_num in range(num_rows): sheet_summary.set_row(row_num, 25) # Высота строки в пикселях (стандартная 15, 30 будет в 2 раза больше)
				for i, col in enumerate(summary_pivot.columns):
					max_len = max(len(str(col)), summary_pivot[col].astype(str).map(len).max() if not summary_pivot.empty else 0)
					sheet_summary.set_column(i + 1, i + 1, max_len + 2)  # +1 для колонок данных, т.к. 0-й столбец - индекс
				# Настройка ширины для первой колонки (индекс заболеваний)
				max_index_len = max(len(str(idx)) for idx in summary_pivot.index)
				sheet_summary.set_column(0, 0, max_index_len + 2)
try:
	Input_lag()
	GetSQL()
	SendTelegram("try", "")
except requests.exceptions.RequestException as err:
    SendTelegram("except1", err)
