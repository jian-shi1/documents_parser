import os
import sys
import argparse
from constants import file_formats, archieve_formats, wordlist_filename
from docx import Document
from numpy.random import choice, randint
from string import ascii_uppercase


class DocumentGenerator:
    _max_words_name = 3
    _max_paragraphs = 100
    _max_paragraph_words = 500
    _max_columns = 100
    _max_rows = 1000
    _max_doc_num = 1
    _max_docx_num = 1
    _max_xls_num = 1
    _max_xlsx_num = 1
    _max_pdf_num = 1
    _max_depth = 2
    _max_zip = 1
    _max_7zip = 1
    _max_rar = 1

    banwords = set()
    word_bank = set()
    all_formats = set()
    workdir = ''

    def __init__(self, dir):
        self.workdir = dir
        self.all_formats = file_formats | archieve_formats
        # банк слов из большого файла с английскими словами
        with open(wordlist_filename) as wb:
            self.word_bank = wb.read().splitlines()

        # банворды для нейминга это уже существующие в директории названия, будет пополняться
        existing_filenames = []
        for _, _, files in os.walk(dir):
            existing_filenames.extend(files)
            break
        self.banwords = set([filename for filename in existing_filenames if os.path.splitext(
            filename)[1] in self.all_formats])

    # как сделать уникальное название?
    # в словарь здесь сохраняем имена файлов уже присутствующих в этой директории.
    # когда генерим новое название, начинаем заменять рандомную букву пока не получим уникальное.
    # добавляем во множество файлов.
    def _gen_wordlist(self, numwords):
        return choice(self.word_bank, numwords)

    def _gen_paragraph(self):
        words_num = randint(1, self._max_paragraph_words)
        return ' '.join(self._gen_wordlist(words_num))

    def _gen_title(self, idx, ext):
        num_words_in_title = randint(1, self._max_words_name + 1)
        s = '_'.join(self._gen_wordlist(num_words_in_title)) + ext
        while s in self.banwords:
            s = '_'.join(self._gen_wordlist(num_words_in_title)) + ext
        self.banwords.add(s)
        return s

    def _generate_doc(self, num, ext):
        for i in range(num):
            doc = Document()
            name = os.path.join( self.workdir, self._gen_title(i, ext))
            num_pars = randint(1, self._max_paragraphs+1)
            for _ in range(num_pars):
                doc.add_paragraph(self._gen_paragraph())
            doc.save(name)
            print(f"created {name}")

    def _generate_xls(self, num, ext):
        raise NotImplementedError  # TODO

    def _generate_pdf(self, num):
        raise NotImplementedError  # TODO

    def generate(self, *, doc_num=1, docx_num=1, xls_num=1, xlsx_num=1,
                 pdf_num=1, zip_num=1, rar_num=1, szip=1):
        existing_filenames = []
        for _, _, filenames in os.walk(self.workdir):
            existing_filenames.extend(filenames)
            break
        existing_filenames = set(existing_filenames)
        self._generate_doc(doc_num, ".doc")
        self._generate_doc(docx_num, ".docx")
        # self._generate_xls(xls_num, ".xls")
        # self._generate_xls(xlsx_num, ".xlsx")
        # self._generate_pdf(pdf_num)
        # TODO: generate archieves with recursion possibility


def main():
    arg_parser = argparse.ArgumentParser(
        description='Generate many files in dir')
    arg_parser.add_argument('dir', type=str, help='Input directory')
    args = arg_parser.parse_args()
    gen = DocumentGenerator(args.dir)
    gen.generate()
    # print(args.dir)

# doc, docx, xls, xlsx, pdf, zip, rar, 7z
# внутри архива может быть другие архивы, но допускаем вложенность не более 3, чтобы не уходить глубоко слишком


"""
как будем генерить файлы:
doc, docx: генерим случайное количество параграфов, для каждого - случайное количество слов
xls, xlsx: случайное количество колонок, случайные названия колонок, случайно заполняем колонки
"""

if __name__ == "__main__":
    main()
