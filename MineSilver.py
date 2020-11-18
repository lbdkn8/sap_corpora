from bs4 import BeautifulSoup
import requests as req
import xml.etree.ElementTree as ET
import openpyxl as pyxl
import os
import string
import re
from transliterate import translit


fname = 'sap_data.xlsx'                                    # file name
columnsnames = ['id', 'author', 'title', 'year',
                'trend', 'genre', 'birth', 'wordcount']

# для нового автора меняются только две строки снизу
author_addr = 'bunin'                                     # смотрится в адресной строке страницы со всеми стихами автора
author = 'Бунин'                                          # просто фамилия
kauthor = translit(author, reversed=True)                 # здесь транслит можно и ёбнуть, *если* мешать будет

trend = 'undef'
genre = 'undef'
birth = '1870'

corepath = f'{os.getcwd()}\\SilverAgePoets_data'
metapath = f'{corepath}\\{fname}'
authorpath = f'{corepath}\\{author}'
filepath = f'{authorpath}\\files'


def parse4poems():
    wnumber = 0
    tlink = 'https://slova.org.ru/{}/'
    ftl = tlink.format(f'{author_addr}/index')
    resp = req.get(ftl)
    soup = BeautifulSoup(resp.text, 'lxml')
    poems = soup.find('div', id='stihi_list')

    links = [tag['href'] for tag in poems.find_all('a')]
    print(
        f'Начинаем обработку.\n'
        f'Всего ссылок: {len(links)}\n'
    )
    copies = dict()
    for link in links:
        l = tlink.format(link)
        poem_resp = req.get(l)
        poem_page = BeautifulSoup(poem_resp.text, 'lxml')

        text_tag = poem_page.find('pre')
        year_tag = text_tag.find('i')
        ptxt = text_tag.text
        year = year_tag.text if year_tag else 'undef'
        title = poem_page.find('h3').text
        copies[title] = 1 if title not in copies.keys() else copies[title] + 1
        unpnctor = re.compile('[%s]' % re.escape(string.punctuation))
        text_unpncted = unpnctor.sub('', ptxt)
        wordcount = len(text_unpncted.split())
        wnumber += wordcount

        try:
            ktitle = translit(title, reversed=True)
        except:
            ktitle = title

        idpoem = f'{year}_{kauthor}_{ktitle}'
        data = [idpoem, author, title, year,
                trend, genre, birth, wordcount]

        root = ET.Element('poem')
        tree = ET.ElementTree(root)
        ET.SubElement(root, 'meta',
                      author=author, year=year, trend=trend,
                      genre=genre, nwords=str(wordcount), source=l)
        b = ET.SubElement(root, 'body')
        ET.SubElement(b, 'title').text = title
        ET.SubElement(b, 'text').text = ptxt

        os.chdir(filepath)
        if f'{ktitle}.xml' not in os.listdir():
            tree.write(f'{filepath}\\{ktitle}.xml', 'utf-8')
        else:
            tree.write(f'{filepath}\\{ktitle}({copies[title]}).xml', 'utf-8')
        os.chdir(authorpath)
        print(
            f'~~~\n'
            f'Название: {title}\n'
            f'Количество слов: {wordcount}\n'
            f'Осталось ссылок: {len(links) - links.index(link) - 1}'
        )
        wbsheet.append(data)
    wb.save(metapath)
    print(f'\n"{author}" - общее количество слов: {wnumber}')


if 'SilverAgePoets_data' not in os.listdir():
    os.mkdir(corepath)
    os.chdir(corepath)
    wb = pyxl.Workbook()
    wbsheet = wb.active
    wbsheet.append(columnsnames)
    wb.save(fname)
    os.mkdir(authorpath)
    os.chdir(authorpath)
    os.mkdir(filepath)
    parse4poems()
else:
    os.chdir(corepath)
    if f'{author}' not in os.listdir():
        wb = pyxl.load_workbook(fname)
        sheets = wb.sheetnames
        wbsheet = wb[sheets[0]]
        os.mkdir(authorpath)
        os.chdir(authorpath)
        os.mkdir(filepath)
        parse4poems()
    else:
        print('Nice?\n'
              'Nice!')
