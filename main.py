import os

import colorama
from oletools.olevba import VBA_Parser

from vba_utils import view_macros, analyse_macros


def get_document():
    try:
        file_path = os.path.join(os.getcwd(), "document/", input("Enter document name: ").strip())

        if not os.path.exists(file_path):
            raise FileNotFoundError("File not found! Check that the file is in the 'document' directory.")
        if not os.path.isfile(file_path):
            raise FileNotFoundError("The provided path is a directory, not a file. Please provide a file path!")
        return file_path
    except FileNotFoundError as e:
        print(e)
        return None
    except Exception as e:
        print("An error occurred: ", e)
        return None


def main():
    colorama.init(autoreset=True)
    document_path = get_document()
    if document_path is not None:
        file = VBA_Parser(document_path)

        while True:
            try:
                view_macros.detect_macros(file)
                analyse_macros.analyse_macros(file)
                break
            except Exception as e:
                print("An error occurred: ", e)


if __name__ == "__main__":
    main()
