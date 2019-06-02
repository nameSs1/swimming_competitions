import xlrd
import pyodbc

# -----------------Excel------------------------------------------------------------------------------------------
location_excel = 'D:\\for SQL Server\\'
excel_name = 'excel_person.xls'  # excel_person2.xls или excel_person.xls
excel_list = '[Лист1$]'  # [ПРОТОКОЛ$] или [Лист1$]
# -----------------SQL--------------------------------------------------------------------------------------------
driver_sql = '{SQL Server}'
server_sql = 'DESKTOP-NE8ID00\\SQLSERVER'
database_sql = 'just_like_others'

ranks = [None, '2юн', '1юн', '3', '2', '1', 'кмс', 'мс', 'мсмк', 'змс']
nambers_0_to_9 = '1234567890'


def get_time(raw_time):  # Преобразует время для второго экселя
    if raw_time == 'дисквал.' or raw_time == 'DSQ':
        return None
    raw_time = raw_time.replace(",", ".")
    raw_time = raw_time.replace(":", ".")
    if raw_time.count(".") == 2:
        raw_time = raw_time.replace(".", ":", 1)
    if (raw_time.index('.') + 2) == len(raw_time):
        raw_time += '0'
    if len(raw_time) == 5:
        raw_time = '00:00:' + raw_time
    else:
        raw_time = '00:0' + raw_time
    return raw_time


def parser_excel_first_type (excel_file):
    event = dict.fromkeys(['title_event', 'date_event', 'city_event', 'pool'])
    competition = dict.fromkeys(['gender', 'distance', 'style', 'birth_year_comp', 'day_comp'])
    competition['day_comp'] = 1
    results = []

    def rang_pars(rang_str):  # Прреобразует разряды к одному типу
        if rang_str == '':
            return None
        elif type(rang_str) is int:
            return str(rang_str)
        elif type(rang_str) is float:
            return str(int(rang_str))
        else:
            return rang_str

    def style_unification(style):  # приведение к одному виду
        if style == 'батт.':
            style = 'баттерфляй'
        elif style == 'в/ст':
            style = 'вольный стиль'
        elif style == 'к/пл':
            style = 'комплексное плавание'
        elif style == 'н/сп':
            style = 'на спине'
        return style

    def parsing_event(string, i):  # Парсинг данных event
        if i == 0:
            event['title_event'] = string[0]
        elif i == 1 or i == 2:
            event['title_event'] += string[0]
        elif i == 4:
            event['city_event'] = string[1][string[1].index('г.') + 2:]
            event['pool'] = int(string[6][string[6].index('бассейн') + 7:string[6].index('м')])
            substring_date = string[1].split(' ', 1)
            event_date = (str(substring_date[0][:2]) + str(substring_date[0][-8:])).split('.')
            event_date.reverse()
            event['date_event'] = '-'.join(event_date)

    def parsing_competition(string):  # Парсинг информации о соревнованиях (дисциплине)
        list_comp = string[1].split()
        competition['distance'] = int(list_comp[0])
        competition['style'] = style_unification(list_comp[1])
        if list_comp[2] == 'девочки':
            competition['gender'] = 'Ж'
        else:
            competition['gender'] = 'М'
        competition['birth_year_comp'] = int(list_comp[3])

    def parsing_swimmer(string):  # Узнаем информацию о плавце
        name = string[1].split()
        year = int(string[2])
        city_club = string[4].split(',', 1)
        city = city_club[0]
        if city == 'Могилёв':  # Костыль для буквы Ё
            city = 'Могилев'
        if len(city_club) == 2:
            club = city_club[1].lstrip()
        else:
            club = city  # если клуба нет, то вместо него город
        time = get_time(string[5])
        rank = rang_pars(string[3])
        new_rank = rang_pars(string[6])
        if ranks.index(rank) < ranks.index(new_rank):
            rank = new_rank
        keys = ['firstname', 'lastname', 'birth_year', 'rank', 'city', 'club', 'time', 'country']
        if club == 'Латвия':
            country = 'LAT'
        else:
            country = None
        values = (name[1], name[0], year, rank, city, club, time, country)
        result = {k: v for k, v in zip(keys, values) if v is not None}
        return result


    for i in range(excel_file.nrows):
        string = excel_file.row_values(i)
        empty_cells = string.count('')  # Считает сколько пустых ячеек
        if empty_cells in (1, 2, 8, 10, 12) or string[0] == '№':  # Пустые строки, ничего не делаем
            continue
        elif i < 5:  # Парсинг данных event
            parsing_event(string, i)
            continue
        elif empty_cells == 11 and string[1] == '':  # Узнаем день соревнований
            competition['day_comp'] = int(string[4][:2])
        elif empty_cells == 11:
            parsing_competition(string)
        else:
            swimmer = parsing_swimmer(string)
            swimmer.update(competition)
            swimmer.update(event)
            results.append(swimmer)
    return results


def parser_excel_second_type(excel_file):
    event = dict.fromkeys(['title_event', 'date_event', 'city_event', 'pool'])
    competition = dict.fromkeys(['gender', 'distance', 'style', 'birth_year_comp', 'day_comp'])
    results = []

    def parsing_event(string, i):  # Парсинг данных event
        if i == 0:
            event['title_event'] = string[1]
        elif i == 1 or i == 2:
            event['title_event'] += string[1]
        elif i == 3:
            event['city_event'] = string[1][string[1].index('г.') + 2:string[1].index(',')]
            event['pool'] = int(string[1][string[1].index('бассейн') + 7:string[1].index('м')])
            substring_date = string[1].split(',')
            event_date = (str(substring_date[2][:2]) + str(substring_date[2][-8:])).split('.')
            event_date.reverse()
            event['date_event'] = '-'.join(event_date)

    def parsing_competition(string):  # Парсим информацию о competition
        string = string.split()
        for j, v in enumerate(string, start=1):
            if j == 1:
                if v == 'Девушки' or v == 'Девочки':
                    competition['gender'] = 'Ж'
                else:
                    competition['gender'] = 'М'
                continue
            elif j == 2:
                if len(v) < 5:
                    competition['birth_year_comp'] = int(v)
                elif len(v) == 8:
                    competition['birth_year_comp'] = int(v[:4])
                else:
                    competition['birth_year_comp'] = int(v[5:9])
                continue
            if '0' in v:
                nine = '1234567890'
                distance = [namber for namber in v if namber in nine]
                competition['distance'] = int(''.join(distance))
                break
        if i == 926 and excel_name == 'excel_person2.xls':
            competition['gender'] = 'М'
        style = []
        if '0' not in string[-2]:
            style.append(string[-2])
        style.append(string[-1])
        competition['style'] = ' '.join(style)

    def parsing_swimmer(string):  # Парсим информацию о плавце
        name = string[1].split()
        year = int('20' + str(string[2]))
        city_club = string[3].split(',', 1)
        if 'Гомель' in city_club[0]:
            city = 'Гомель'
        else:
            city = city_club[0]
        if len(city_club) == 2:
            club = city_club[1]
        else:
            club = None
        country = string[4]
        if empty_cells == 3:
            time = None
        else:
            time = get_time(str(string[5]))
        if string[6] == '':
            points = None
        else:
            points = int(string[6])
        keys = ['firstname', 'lastname', 'birth_year', 'country', 'city', 'club', 'time', 'points']
        values = (name[1], name[0], year, country, city, club, time, points)
        swimmer = {k: v for k, v in zip(keys, values) if v is not None}
        return swimmer

    for i in range(excel_file.nrows):
        string = excel_file.row_values(i)
        empty_cells = string.count('')  # Считает сколько пустых ячеек
        if empty_cells in (1, 4, 5, 6, 7, 9):  # Пустые строки, ничего не делаем
            continue
        elif empty_cells == 8 and i < 4:  # Узнаем информацию о мероприятии(event)
            parsing_event(string, i)
        elif empty_cells == 8 and 'день' in string[1]:  # Узнаем день соревнований
            competition['day_comp'] = int(string[1][:2])
        elif empty_cells == 8:  # Узнаем информацию о competition
            parsing_competition(string[1])
        elif string[0] != '':
            swimmer = parsing_swimmer(string)
            swimmer.update(competition)
            swimmer.update(event)
            results.append(swimmer)
    return results


def reading_excel(location_excel, excel_name):
    rb = xlrd.open_workbook(location_excel + excel_name, formatting_info=True)
    sheet = rb.sheet_by_index(0)
    count_cols = (sheet.ncols)  # Определяем тип файла эксель
    if count_cols == 12:
        return parser_excel_first_type(sheet)
    elif count_cols == 9:
        return parser_excel_second_type(sheet)
    else:
        str_for_print = "Не подходящий тип экселя. Колличество столбцов: {}. Необходимо 9 или 12".format(count_cols)
        print(str_for_print)


def connect_sql_server(driver_sql, server_sql, database_sql):  # Подключается с рерверу, возвращает курсор
    connection_str_sql_server = "Driver={}; Server={}; Database={};".format(driver_sql, server_sql, database_sql)
    conn_sql_server = pyodbc.connect(connection_str_sql_server)
    return conn_sql_server.cursor()


def insert_into_tables(results):  # вставляем данные по таблицам
    def create_string_for_sql(dictionary, table_name):  # Принимает словарь и название таблицы. Формирует сторку для SQL
        string_one = "insert into [{}] (".format(table_name)  # вставляем название таблицы
        string_two = ")  values ("
        keys_and_values = tuple(dictionary.items())  # получаем названия полей и их значения
        columns = [(str(k[0]) + ",") for k in keys_and_values]  # получаем список полей в виде подстрок
        columns[-1] = columns[-1].rstrip(',')
        values = [("'" + str(k[1]) + "',") for k in keys_and_values]  # получаем список значений в виде подстрок
        values[-1] = values[-1].rstrip(',')  # доробатываем подстроку последнего значения убрав запетую
        columns = ' '.join(columns)  # обьеденяем подстроки в строки
        values = ' '.join(values)  # обьеденяем подстроки в строки
        string_finish = string_one + columns + string_two + values + ")"  # склеиваем в одну строку
        return string_finish

    def create_string_for_select_id(dictionary, table_name):  # формирует строку для получения id
        # select table_name + ID from table_name where dictionary
        str_where = [(str(k) + " = '" + str(dictionary[k]) + "'") for k in dictionary.keys()]  # список для строки where
        str_where = ' and '.join(str_where)
        str_select_id = "select [{0}ID] from [{0}] where {1}".format(table_name, str_where)
        return str_select_id

    def insert_select_id(values, table):  # Вставляем значения в таблицу, получаем значение ID
        dictionary = {v[0]: v[1] for v in values}
        insert = create_string_for_sql(dictionary, table)
        try:
            cursor.execute(insert)
        except:
            select_id = create_string_for_select_id(dictionary, table)
            cursor.execute(select_id)
            id = cursor.fetchone()[0]
        else:
            cursor.commit()
            cursor.execute("SELECT SCOPE_IDENTITY() AS ID")
            id = cursor.fetchone()[0]
            if table == 'Competition':
                correct_date_competition(id, result['day_comp'] - 1)  # корректеровка даты соревнований
        return id

    def correct_date_competition(id, add_days):  # исправляет дату в Competition
        update_str = "update Competition set Date = DATEADD(day, {}, Date) where CompetitionID = {}" \
                     "".format(add_days, id)
        try:
            cursor.execute(update_str)
        except:
            pass
        else:
            cursor.commit()

    for result in results:  # Идем по списку результатов
        pool_columns = (('City', result['city_event']), ('Name', result['title_event']), ('PoolSize', result['pool']))
        pool_id = insert_select_id(pool_columns, 'Pool')
        discipline_columns = (('Style', result['style']), ('Distance', result['distance']))
        discipline_id = insert_select_id(discipline_columns, 'Discipline')
        group_columns = (('Name', result['birth_year_comp']), ('Gender', result['gender']))
        group_id = insert_select_id(group_columns, 'Group')
        competition_columns = (('GroupID', group_id),
                               ('DisciplineID', discipline_id),
                               ('PoolID', pool_id),
                               ('Date', result['date_event']))
        competition_id = insert_select_id(competition_columns, 'Competition')
        swimming_club_columns = (('Name', result['club']), ('City', result['city']))
        swimming_club_id = insert_select_id(swimming_club_columns, 'SwimmingClub')
        swimmer__columns = (('SwimmingClubID', swimming_club_id),
                            ('FirstName', result['firstname']),
                            ('LastName', result['lastname']),
                            ('YearOfBirth', result['birth_year']),
                            ('Gender', result['gender']))
        swimmer_id = insert_select_id(swimmer__columns, 'Swimmer')
        if 'time' not in result:
            table = 'Disqualification'
            columns = (('CompetitionID', competition_id), ('SwimmerID', swimmer_id))
        else:
            table = 'Result'
            columns = (('CompetitionID', competition_id), ('SwimmerID', swimmer_id), ('Time', result['time']))
        result_id = insert_select_id(columns, table)


results = reading_excel(location_excel, excel_name)
cursor = connect_sql_server(driver_sql, server_sql, database_sql)
insert_into_tables(results)