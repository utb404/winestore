import datetime
import collections
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape
import pandas


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


def get_year(start_year):
    year = datetime.datetime.now().year - start_year
    return year


def render_page(template, year, wines):
    rendered_page = template.render(year=year, wines=sorted(wines.items()))
    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)


def start_server():
    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    render_page(open_template('template.html'),
                get_year(1920),
                get_wines('wine3.xlsx'))
    start_server()
