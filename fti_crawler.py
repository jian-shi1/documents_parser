import argparse
import docx
import patoolib
import pandas as pd
import os
import xlrd

from constants import file_formats, archieve_formats


class DocumentParser:
    all_formats = set()

    # def normal_document_entry(filename, )

    def read_doc(filename, df):

        doc = docx.Document(filename)
        text = [p.text for p in doc.paragraphs]
        df["text"].append('\n'.join(text))
        # raise NotImplementedError

    def read_xls(filename, df):
        workbook = xlrd.open_workbook(filename)
        sheets_data = []
        for sheet in workbook.sheets():
            sheets_data.append('\n'.join(" ".join(sheet.row_values(rownum))
                               for rownum in range(sheet.nrows)))

        df['text'].append("\n".join(sheets_data))

    def read_xlsx(filename, df):
        print(f"Processing {filename}...")
        content = pd.read_excel(filename)
        df['text'].append(" ".join(content.columns) + "\n" + "\n".join(' '.join(rowtup)
                          for rowtup in content.itertuples(index=False, name=None)))

    def read_pdf(filename, df):

        raise NotImplementedError

    def read_archieve(df):
        ...
        # extract archieve
        # raise NotImplementedError

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

    def parse(self, workdir, df_dict, is_archieve=False, archieve_name=''):
        # os walk:
        existing_filenames = []
        for cur_dir, dirs, files in os.walk(workdir):
            # print(cur_dir, dirs, files)
            for file in files:
                full_processed_file = os.path.join(
                    os.path.abspath(cur_dir), file)
                filename, ext = os.path.splitext(full_processed_file)
                # print(filename, ext)
                if ext in archieve_formats:
                    print(f"Archieve {filename}, skip for now.")
                    # unzip
                    # process a folder recursively
                    # delete folder
                    ...
                else:
                    if ext in {'.docx', '.doc', ".xls", ".xlsx"}:  # TODO: убрать это
                        df_dict["name"].append(filename)
                        df_dict["extension"].append(ext)
                        df_dict["is_archieved"].append(int(is_archieve))
                        df_dict["archieve_name"].append(archieve_name)
                        self.process_file[ext](full_processed_file, df_dict)

        # meet an archieve - work it as dir
        # recursively go to every dir
        # to_csv
        # TODO: insert into sql
        # raise NotImplementedError


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
        "is_archieved": [],
        "archieve_name": [],
    }
    document_parser.parse(args.dir, df, False)
    # print(df)
    # for k, v in df.items():
    #     print(k, len(v))
    df = pd.DataFrame(df)
    # df.astype({"is_archieved": 'int32'}, copy=False)
    # print(df.head())
    df.to_csv(args.csv, index=False)
    print(f"Files are processed and uploaded to {args.csv}")


if __name__ == "__main__":
    main()
