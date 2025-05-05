import argparse
from constants import file_formats, archieve_formats, db_name
import docx
import os
import pandas as pd
from patoolib import extract_archive
from pypdf import PdfReader
from shutil import rmtree
import xlrd
import sqlite3


class DocumentParser:
    all_formats = set()

    def read_doc(filename, df):
        doc = docx.Document(filename)
        text = [p.text for p in doc.paragraphs]
        df["text"].append('\n'.join(text))

    def read_xls(filename, df):
        workbook = xlrd.open_workbook(filename)
        sheets_data = []
        for sheet in workbook.sheets():
            sheets_data.append('\n'.join(" ".join(sheet.row_values(rownum))
                               for rownum in range(sheet.nrows)))
        df['text'].append("\n".join(sheets_data))

    def read_xlsx(filename, df):
        content = pd.read_excel(filename)
        df['text'].append(" ".join(content.columns) + "\n" + "\n".join(' '.join(rowtup)
                          for rowtup in content.itertuples(index=False, name=None)))

    def read_pdf(filename, df):
        reader = PdfReader(filename)
        df['text'].append('\n'.join(page.extract_text()
                          for page in reader.pages))

    def __init__(self):
        self.all_formats = file_formats | archieve_formats
        if '' in self.all_formats:
            self.all_formats.remove('')

    process_file = {
        ".doc": read_doc,
        ".docx": read_doc,
        ".xls": read_xls,
        ".xlsx": read_xlsx,
        ".pdf": read_pdf
    }

    def parse(self, workdir, df_dict, is_archieve=False, archive_name=''):
        existing_filenames = []
        for cur_dir, dirs, files in os.walk(workdir):
            for file in files:
                full_processed_file = os.path.join(
                    os.path.abspath(cur_dir), file)
                filename, ext = os.path.splitext(full_processed_file)
                if ext in archieve_formats:
                    # unzip
                    folder = os.path.join(os.path.dirname(
                        full_processed_file), filename)
                    extract_archive(full_processed_file,
                                    outdir=folder, verbosity=0)
                    # process a folder recursively
                    self.parse(folder, df_dict, True, full_processed_file)
                    # delete folder
                    rmtree(folder)
                else:
                    df_dict["name"].append(filename)
                    df_dict["extension"].append(ext)
                    df_dict["is_archived"].append(int(is_archieve))
                    df_dict["archive_name"].append(archive_name)
                    df_dict["filepath"].append(str(full_processed_file))
                    self.process_file[ext](full_processed_file, df_dict)


def main():
    arg_parser = argparse.ArgumentParser(
        description='Parse files in directory')
    arg_parser.add_argument('dir', type=str, help='Input directory')
    arg_parser.add_argument('csv', type=str, help='Name of csv final document')
    args = arg_parser.parse_args()
    document_parser = DocumentParser()
    df = {
        "name": [],
        "text": [],
        "extension": [],
        "filepath": [],
        "is_archived": [],
        "archive_name": [],
    }
    document_parser.parse(args.dir, df, False)
    df = pd.DataFrame(df)
    df.to_csv(args.csv, index=False)
    print(f"Files are processed and uploaded to {args.csv}")

    # наша FTI-таблица виртуальная, так что to_sql сразу не сработает
    # переносить будем через временную таблицу
    conn = sqlite3.connect(db_name)
    df.to_sql("tmp", conn, if_exists='replace', index=False)
    conn.execute("""
        INSERT INTO documents
        SELECT name, text, extension, filepath, is_archived, archive_name 
        FROM tmp;
    """)
    conn.execute("DROP TABLE tmp")
    conn.commit()
    conn.close()
    print("Data is inserted to SQL")


if __name__ == "__main__":
    main()
