import argparse
import collections
import datetime
import pandas

from http.server import HTTPServer
from http.server import SimpleHTTPRequestHandler
from jinja2 import Environment
from jinja2 import FileSystemLoader
from jinja2 import select_autoescape


def open_template(template):
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )
    template = env.get_template(template)
    return template


def get_wines(excel_file):
    excel_wines = pandas.read_excel(excel_file, sheet_name='Лист1',
                                    na_values='None', keep_default_na=False)
    wines = collections.defaultdict(list)
    for wine in excel_wines.to_dict(orient='records'):
        wines[wine['Категория']].append(wine)
    return wines


def render_page(template, year, wines):
    rendered_page = template.render(year=year, all_wines=sorted(wines.items()))
    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)


def start_server():
    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('-f', '--file_path', default='template.xlsx', help='Путь к файлу с перечнем напитков' )
    file_path = parser.parse_args().file_path
    year = datetime.datetime.now().year - 1920
    render_page(open_template('template.html'),
                year, get_wines(file_path))
    start_server()
