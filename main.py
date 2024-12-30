from http.server import HTTPServer, SimpleHTTPRequestHandler
import argparse
import datetime
from pprint import pprint
from collections import defaultdict

import pandas as pd
from jinja2 import Environment, FileSystemLoader, select_autoescape


def get_actual_age():
    year_of_foundation = 1920
    date = (datetime.datetime.now()).year - year_of_foundation
    last_digit = date % 10
    last_two_digit = date % 100

    if last_two_digit in [11, 12, 13, 14]:
        return f'Уже {date} лет с вами'
    elif last_digit == 1:
        return f'Уже {date} год с вами'
    elif last_digit in [2, 3, 4]:
        return f'Уже {date} года с вами'
    else:
        return f'Уже {date} лет с вами'


def main(file_path):
    wine_from_excel = pd.read_excel(file_path, sheet_name='Лист1').fillna('')
    wines = wine_from_excel.to_dict(orient='records')
    wines_by_category = defaultdict(list)

    for wine in wines:
        category = wine['Категория']
        wines_by_category[category].append(wine)

    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('template.html')
    rendered_page = template.render(
        age_of_winery=get_actual_age(),
        wines=wines_by_category
    )

    with open('index.html', 'w', encoding='utf8') as file:
        file.write(rendered_page)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Обработка таблицы с напитками'
    )
    parser.add_argument(
        'file', metavar='F', type=str,
        nargs='?', default='wine.xlsx',
        help='Путь к excel файлу (по умолчанию wine.xlsx в текущей папке)'
    )

    args = parser.parse_args()
    main(args.file)

server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
server.serve_forever()
