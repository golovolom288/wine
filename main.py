from http.server import HTTPServer, SimpleHTTPRequestHandler

from jinja2 import Environment, FileSystemLoader, select_autoescape

from datetime import datetime

import pandas

from collections import defaultdict


def get_wine_age(start_age):
    age = datetime.now().year - start_age
    if 11 <= age % 100 <= 20 or age % 10 == 0:
        year_word = "лет"
    elif age % 10 == 1:
        year_word = "год"
    else:
        year_word = "года"
    return f"{age} {year_word} с вами."


def load_excel_data(excel_path):
    excel_data = pandas.read_excel(excel_path)
    excel_data.fillna("", inplace=True)
    excel_data = excel_data.to_dict(orient="records")
    wines = defaultdict(list)
    for row in excel_data:
        category = row["Категория"]
        wines[category].append(row)
    return wines


def render_template(data):
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )
    template = env.get_template('template.html')
    return template.render(data)


def main():
    data_to_render = {
        "age": get_wine_age(1920),
        "wines": load_excel_data("wine3.xlsx")
    }
    rendered_page = render_template(
        data_to_render
    )
    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)
    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == "__main__":
    main()
