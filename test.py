from urllib.request import urlopen
from bs4 import BeautifulSoup
import re

class WordCounterOnTheSite:
    def __init__(self, url):
        self.url = url
        try:
            self.html = urlopen(str(url)).read().decode('utf-8')
        except Exception:
            raise Exception ('Сайт ' + url + ' - не найден.')



    """
    Количество всех вхождений слова, с учётом 
    системной хуйни (с учётом регистра).
    """
    def total_score(self, word):
        return str(self.html.count(word)) + " вхождений."    # Ответ.



    """
    Количество вхождений слова, которые видно 
    в браузере (с учётом регистра).
    """
    def visible_on_the_site(self, word):
        s = self.html                   # Короткий вид для переменной.

        s = s[s.index('<body'):]        # Сносим заголовок.

        y = r'<script'                  # Сносим скрипты.
        z = re.findall(y, s)
        for script in z:
            s1 = s.index('<script')
            s2 = s.index('/script>') + 8
            s = s[:s1] + s[s2:]

        Python_re = r'(\<[^\<\>]*\>)'   # Сносим остальную системную хуйню.
        z = re.findall(Python_re, s)
        for el in z:
            if el in s:
                s = s.replace(el, '*')

        return str(s.count(word)) + " вхождений."       # Ответ.



    """
    Нахождение максимально часто встречающихся строк между 
    тегами <code> и </code> (вывод в алфавитном порядке).
    """
    def frequent_line_in_code_tag(self):
        s = self.html                   # Короткий вид для переменной.

        # Находим все нужные строки по условию.
        y = r'\<code\>\w+<\/code\>'
        z = re.findall(y, s)

        # Если на сайте нет таких строк:
        if len(z) == 0:
            return 'Строки с тегами "code" не найдены.'

        # Считаем количество повторений для каждой строки.
        d = {}
        for code in z:
            if code not in d:
                d[code] = 1
            else:
                d[code] += 1

        # Вытаскиваем максимально часто встречающиеся строки.
        i = 0
        for key, value in d.items():
            if value > i:
                i = value
                j = []
                j.append(key)
            if value == i:
                if key not in j:
                    j.append(key)

        # Убираем теги и сортируем список.
        new_j = []
        for designs in j:               #
            designs = designs.removeprefix('<code>')
            designs = designs.removesuffix('</code>')
            new_j.append(designs)
        new_j.sort()
        str_out = ''
        for word in new_j:
            str_out += word + ' '
        return str_out                  # Ответ.



    """
    Суммирование значений ячеек таблицы формата html.
    """
    def sum_of_cell_values(self):
        # Вытаскиваем ячейки с сайта.
        soup = BeautifulSoup(self.html, 'html.parser')
        z = soup.find_all('td')

        # Если на сайте нет ячеек таблицы:
        if len(z) == 0:
            return 'Ячейки таблицы не найдены.'

        # Суммируем значения ячеек.
        sum = 0
        unk_str = 0
        for cell in z:
            cell = str(cell.string)
            cell = cell.removeprefix(' ')
            cell = cell.removesuffix(' ')
            try:
                sum += int(cell)
            except ValueError:
                unk_str +=1
        return ('Сумма значений всех ячеек: ' + str(sum) + '\n'
                + 'Количество ячеек, не являющихся числами: ' + str(unk_str))     # Ответ.




x = WordCounterOnTheSite('https://stepik.org/media/attachments/lesson/209723/5.html')
print(x.total_score('Python'))
print(x.visible_on_the_site('Python'))
print(x.frequent_line_in_code_tag())
print(x.sum_of_cell_values())