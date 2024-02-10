import xlrd
import jpype
jpype.startJVM()
from asposecells.api import Workbook
import os
import pandas as pd
import openpyxl


"""
Конвертация в читаемый формат.
"""
def xls_converting(file_name):
    workbook = Workbook(file_name)
    workbook.save("work_file.xls")



class TableProcessing:
    def __init__(self, file_name):

        # Получаем данные из файла.
        try:
            wb = xlrd.open_workbook(file_name)

        except Exception:
            try:
                xls_converting(file_name)
                wb = xlrd.open_workbook('work_file.xls')
            except Exception:
                raise Exception('Файл не найден')
        self.wb = wb



    """
    Вычисление региона с самой высокой медианной з/п;
    Вычисление самой высокооплачиваемой профессии.
    (файл с одним листом).
    """
    def salary_calculation(self):

        # - 1: Вводим переменные (словари, данные из файла).
        d_region = {}
        d_prof = {}
        sh = self.wb.sheet_by_index(0)
        vals = [sh.row_values(rownum) for rownum in range(sh.nrows)]

        # - 1.1: Создаём словарь для текущего листа; вносим туда все виды профессий.
        for profs in vals[0][1:]:
            d_prof[profs] = 0

        # - 1.2: Вносим в словарь с медианой названия регионов; заполняем значения обоих словарей.
        for j in vals[1:]:
            numbers = j[1:]
            if len(numbers) > 0:
                try:
                    z = 0
                    for key, value in d_prof.items():
                            d_prof[key] += numbers[z]
                            z += 1
                    numbers.sort()
                    central_numb = float(numbers[len(numbers) // 2])
                    central_numb_2 = float(numbers[len(numbers) // 2 - 1])
                    if len(numbers) % 2 != 0:
                        d_region[j[0]] = central_numb
                    else:
                        d_region[j[0]] = (central_numb + central_numb_2) / 2
                except ValueError:
                    continue

        # - 2: Узнаём название региона с самой высокой медианой.
        p = 0
        region_output = ''
        for key, value in d_region.items():
            if value > p:
                region_output = key
                p = value
            elif value == p:
                region_output += ', ' + key

        # - 3: Узнаём название профессии с самой высокой оплатой по регионам.
        p = 0
        prof_output = ''
        for key, value in d_prof.items():
            if value > p:
                prof_output = key
                p = value
            elif value == p:
                prof_output += ', ' + key
        out = region_output + ' ' + prof_output

        # - 4: Получаем ответ.
        return out



    """
    Список самых каллорийных продуктов (сразу выводит результат по вызову функции)
    (файл с одним листом).
    """
    def nutritious_food(self):

        # - 1: Вводим переменные (словари, списки, данные из файла).
        d= {}
        sp = []
        sp_new = []
        sh = self.wb.sheet_by_index(0)
        vals = [sh.row_values(rownum) for rownum in range(sh.nrows)]

        # - 1.1: Заполняем словарь необходимыми значениями.
        for j in vals:
            if len(j) > 1:
                try:
                    if float(j[1]) not in d:
                        d[float(j[1])] = j[0]
                    else: d[float(j[1])] += '*' +j[0]
                except ValueError:
                    continue

        # - 2: Переносим значения в список и сортируем по коллориям.
        for key, value in d.items():
            sp.append(str(key) + '*' + value)
        sp.sort()
        sp.reverse()

        # - 3: Сортируем лексикографически блюда с одинаковой каллорийностью.
        for el in sp:
            sp_el = el.split('*')
            sp_el.sort()
            for name in sp_el[1:]:
                sp_new.append(name)

        # - 4: Печатаем список блюд.
        for food in sp_new:
            print(food)



    """
    Подсчёт энергетической ценности для имеющейся еды
    (файл с двумя листами: Справочник и Раскладка).
    """
    def food_energic(self):

        # - 1: Вводим переменные (словари, списки, данные из файла).
        d = {}
        sh_1 = self.wb.sheet_by_name('Справочник')
        sh_2 = self.wb.sheet_by_name('Раскладка')
        vals_energic = [sh_1.row_values(rownum) for rownum in range(sh_1.nrows)]
        vals_food = [sh_2.row_values(rownum) for rownum in range(sh_2.nrows)]

        # - 1.1: Заполняем словарь значениями (вес/ккал/Б/Ж/У).
        for j in vals_food[1:]:
            if len(j) > 1:
                try:
                    if j[0] not in d:
                        d[j[0]] = str(j[1])
                    else:
                        d[j[0]] = str(float(d[j[0]]) + float(j[1]))
                except ValueError:
                    continue
        for j in vals_energic[1:]:
            if len(j) > 4:
                if j[0] in d:
                    try:
                        for score in j[1:5]:
                            if score == '':
                                score = 0
                            d[j[0]] += '/'
                            d[j[0]] += str(float(score)/100)
                    except ValueError:
                        continue

        # - 2: Считаем суммарное количество Кккал/Б/Ж/У; округляем.
        sp_out = [0, 0, 0 ,0]
        for value in d.values():
            energic = value.split('/')[1:5]
            weight = value.split('/')[0]
            i = 0
            while i < 4:
                sp_out[i] += float(energic[i])*float(weight)
                i += 1
        str_out = ''
        for result in sp_out:
            str_out += str(int(result)) + ' '

        # - 3: Ответ.
        return str_out


    """
    Подсчёт энергетической ценности для имеющейся еды на каждый день похода.
    (файл с двумя листами: Справочник и Раскладка).
    """
    def food_energic_all_days(self):

        # - 1: Вводим переменные (словари, списки, данные из файла).
        day_lst = []
        d_out = {}
        sh_1 = self.wb.sheet_by_name('Справочник')
        sh_2 = self.wb.sheet_by_name('Раскладка')
        vals_energic = [sh_1.row_values(rownum) for rownum in range(sh_1.nrows)]
        vals_food = [sh_2.row_values(rownum) for rownum in range(sh_2.nrows)]

        # - 1.1: Создаём список дней.
        for j in vals_food[1:]:
            if j[0] not in day_lst:
                day_lst.append(int(j[0]))
        day_lst.sort()

        # - 2: Считаем суммарное количество Кккал/Б/Ж/У; округляем (для каждого дня).
        for h in day_lst:

            # - 2.1: Заполняем словарь значениями (вес/ккал/Б/Ж/У), (для текущего дня).
            d = {}
            for j in vals_food[1:]:
                if j[0] == h:
                    if len(j) > 2:
                        try:
                            if j[1] not in d:
                                d[j[1]] = str(j[2])
                            else:
                                d[j[1]] = str(float(d[j[1]]) + float(j[2]))
                        except ValueError:
                            continue
            for j in vals_energic[1:]:
                if len(j) > 4:
                    if j[0] in d:
                        try:
                            for score in j[1:5]:
                                if score == '':
                                    score = 0
                                d[j[0]] += '/'
                                d[j[0]] += str(float(score) / 100)
                        except ValueError:
                            continue

            # - 2.2: Считаем суммарное количество Кккал/Б/Ж/У; округляем (для текущего дня).
            sp_out = [0, 0, 0, 0]
            for value in d.values():
                energic = value.split('/')[1:5]
                weight = value.split('/')[0]
                i = 0
                while i < 4:
                    sp_out[i] += float(energic[i]) * float(weight)
                    i += 1
            str_out = ''
            for result in sp_out:
                str_out += str(int(result)) + ' '
            str_out = str_out.rstrip(' ')
            d_out[h] = str_out

        # - 3: Ответ.
        for k in day_lst:
            print(d_out[k])



"""
Функция для заполнения общей ведомости по имеющимся расчётным листкам.
"""
def salary_calculation_using_tables(dir_name):
    # - 1: Создаём пустой список и заполняем значениями "ФИО, Начислено".
    sp_out = []
    # - 1.1: Читаем файлы с ЗП сотрудников.
    for file in os.listdir(dir_name):
        filename = dir_name + '/' + file
        df = pd.read_excel(filename, engine='openpyxl')
        # - 1.2: Достаём нужные значения.
        for value in df.values:
            i = 0
            while i < len(value):
                if value[i] == 'ФИО':
                    name = value[i+1]
                elif value[i] == 'Начислено':
                    money = value[i+1]
                i += 1
        # - 1.3: Добавляем значения в список.
        sp_out.append((name, ' ', str(int(money))))
    # - 2: Сортируем спсиок по алфавиту.
    sp_out.sort()

    # - 3: Записываем результат в блокнот.
    with open('work_file.txt', 'w', encoding="UTF-8") as ouf:
        for els in sp_out:
            for el in els:
                ouf.write(el)
            ouf.write('\n')






# x = TableProcessing('trekking3.xlsx')
# print(x.salary_calculation())          # salaries.xlsx
# x.nutritious_food()                    # trekking1.xlsx
# print(x.food_energic())                # trekking2.xlsx
# x.food_energic_all_days()              # trekking3.xlsx


salary_calculation_using_tables('roga')

