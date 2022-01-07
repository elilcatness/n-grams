import csv
import os
from argparse import ArgumentParser
from docx import Document

EXTENSIONS = ['txt', 'doc', 'docx']


class InvalidFileExtension(Exception):
    pass


def get_extension(filename: str):
    return filename.split('.')[-1]


def get_words_from_file(filename: str):
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
    return [x.strip(',.:;-–—_') for x in text.lower().lstrip('\n').split() if x.strip(',.:;-–—_')]


def main(raw_n_grams: str):
    n_grams = []
    for n_gram in raw_n_grams.split(','):
        try:
            n_grams.append(int(n_gram.strip().replace(' ', '')))
        except ValueError:
            continue
    if not n_grams:
        return print('Ни одна N-грамма не прошла валидацию')
    n_grams.sort()
    data = {n_gram: {} for n_gram in n_grams}
    filenames = [f_name for f_name in os.listdir() if get_extension(f_name) in EXTENSIONS]
    total_words_length = 0
    for filename in filenames:
        words = get_words_from_file(filename)
        if not words:
            continue
        supported_n_grams = [n_gram for n_gram in n_grams if n_gram <= len(words)]
        words_length = len(words)
        total_words_length += words_length
        for i in range(words_length):
            for n_gram in supported_n_grams:
                if i + n_gram < words_length:
                    phrase = ' '.join(words[i:i + n_gram])
                    data[n_gram][phrase] = data[n_gram].get(phrase, 0) + 1
    output = {key: list(data[key].items()) for key in data.keys()}
    with open('output.csv', 'w', newline='', encoding='utf-8') as csv_file:
        writer = csv.writer(csv_file, delimiter=';')
        row = []
        for n in n_grams:
            row.extend([f'{n}_phrase', f'{n}_count', f'{n}_percent'])
        writer.writerow(row)
        for i in range(len(max(output.values(), key=lambda x: len(x)))):
            row = []
            missed_count = 0
            for n in n_grams:
                if i < len(output[n]):
                    row.extend([output[n][i][0], output[n][i][1],
                               f'{round(output[n][i][1] / total_words_length * 100, ndigits=15)}%'])
                else:
                    missed_count += 1
            for _ in range(missed_count * 3):
                row.insert(0, '')
            writer.writerow(row)


if __name__ == '__main__':
    parser = ArgumentParser()
    parser.add_argument('-n-grams', type=str, required=False)
    n_grams = parser.parse_args().n_grams
    while not n_grams:
        n_grams = input('Введите через запятую N-граммы: ')
    main(n_grams)
