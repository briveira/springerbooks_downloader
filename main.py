import xlrd
import requests
import os
from bs4 import BeautifulSoup


main_folder = "descargas"


def lista_libros():
    wb = xlrd.open_workbook('FreeEnglishTextbooks.xlsx')
    sheet = wb.sheet_by_index(0)

    # leemos lista
    ret = []
    for i in range(sheet.nrows):
        if i > 0:
            ret.append({
                'title': sheet.cell_value(i, 0).replace(' ', '_').replace('/', '-'),
                'package': sheet.cell_value(i, 11).replace(' ', '_'),
                'year': sheet.cell_value(i, 4),
                'url': sheet.cell_value(i, 18)
            })
    return ret


def descarga_paginas_de_cada_libro(books):

    if not os.path.exists(main_folder):
        os.mkdir(main_folder)

    # descargamos paginas si no lo hemos hecho ya

    s = requests.Session()
    for libro in books:
        carpeta = f"{main_folder}/{libro['package']}"
        html_file = f'{carpeta}/{libro["title"]}.html'
        if not os.path.exists(carpeta):
            os.mkdir(carpeta)
        if not os.path.exists(html_file):
            try:
                req = s.get(libro['url'])
                page = req.content
                with open(html_file, 'w+') as f:
                    f.write(str(page))
                print(f'descargado {html_file}')
            except Exception as e:
                print(f'{libro["title"]} - excepcion {str(e)}')


def procesado_archivos(books):

    s = requests.Session()
    for libro in books:
        carpeta = f"{main_folder}/{libro['package']}"
        html_file = f'{carpeta}/{libro["title"]}.html'
        download_file = f'{carpeta}/{libro["title"]}_{int(libro["year"])}'
        if os.path.exists(html_file):
            with open(html_file, 'r') as f:
                html = f.read()
                soup = BeautifulSoup(html, 'html.parser')
                for item in soup.find_all('div', class_='cta-button-container__item'):
                    for a in item.find_all('a'):
                        link_descarga = f'https://link.springer.com{a.attrs["href"]}'
                        extension = link_descarga.split(".")[-1]
                        file = f'{download_file}.{extension}'
                        if not os.path.exists(file):
                            req = s.get(link_descarga)
                            if req.status_code == 200:
                                with open(file, 'wb+') as f:
                                    f.write(req.content)
                                    print(f'descargado - {file}')
                            else:
                                print(f'no encontrado: {link_descarga}')


if __name__ == '__main__':
    books = lista_libros()
    descarga_paginas_de_cada_libro(books)
    procesado_archivos(books)