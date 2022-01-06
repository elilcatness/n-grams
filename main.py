import os
from argparse import ArgumentParser
from docx import Document

EXTENSIONS = ['txt', 'doc', 'docx']


class InvalidFileExtension(Exception):
    pass


def get_extension(filename: str):
    return filename.split('.')[-1]


def get_text_from_doc(filename: str):
    extension = get_extension(filename)
    if extension not in EXTENSIONS:
        raise InvalidFileExtension(f'Неверное расширение файла: {filename}')
    if extension == 'txt':
        with open(filename, encoding='utf-8') as f:
            text = f.read().strip()
    else:
        text = ''
        doc = Document(filename)
        for p in doc.paragraphs:
            text += '\n' + p.text
    return text.lstrip('\n')


def main(n_grams: str):
    given_n_grams = [int(x.strip()) for x in n_grams.split(',') if x.strip().isdigit() and int(x.strip())]
    filenames = [f_name for f_name in os.listdir() if get_extension(f_name) in EXTENSIONS]
    for filename in filenames:
        text = get_text_from_doc(filename)
        if not text:
            continue
        supported_n_grams = [n_gram for n_gram in given_n_grams if n_gram <= len(text)]
        print(filename, ', '.join(map(str, supported_n_grams)))

if __name__ == '__main__':
    parser = ArgumentParser()
    parser.add_argument('-n-grams', type=str, required=False)
    n_grams = parser.parse_args().n_grams
    if not n_grams:
        n_grams = input('Введите через запятую N-граммы: ')
    main(n_grams)