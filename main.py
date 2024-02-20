import os

import colorama
from oletools.olevba import VBA_Parser

from vba_utils import view_macros, analyse_macros, deobfuscate_macros


def display_info():
    banner = """
    ------------------------------------------------
      __  __                                     
     |  \/  |                                    
     | \  / | __ _  ___ _ __ ___  ___            
     | |\/| |/ _` |/ __| '__/ _ \/ __|           
     | |  | | (_| | (__| | | (_) \__ \           
     |_| /\_|\__,_|\___|_|| \___/|___/           
        /  \   _ __   __ _| |_   _ ___  ___ _ __ 
       / /\ \ | '_ \ / _` | | | | / __|/ _ \ '__|
      / ____ \| | | | (_| | | |_| \__ \  __/ |   
     /_/    \_\_| |_|\__,_|_|\__, |___/\___|_|   
                              __/ |              
                             |___/               
    ------------------------------------------------
    """
    print(banner)

    info = """
    This is a simple tool to detect, analyse and 
    deobfuscate VBA macros in Microsoft Office files.
    It is simply a wrapper around the oletools 
    library. This tool should not be used for any
    production purposes! By continuing, you agree
    that you have read and understood the disclaimer.
    Press ENTER to continue...

    
    ------------------------------------------------
    """
    print(info)
    input()
    print("\033[2J\033[H", end="", flush=True)


def initialize():
    colorama.init(autoreset=True)
    display_info()
    document_path = get_document()
    print("\033[2J\033[H", end="", flush=True)
    file = VBA_Parser(document_path)
    return file


def get_document():
    while True:
        try:
            file_path = os.path.join(os.getcwd(), "document/", input("Enter document name: ").strip())

            if not os.path.exists(file_path):
                raise FileNotFoundError("File not found! Check that the file is in the 'document' directory.")
            if not os.path.isfile(file_path):
                raise FileNotFoundError("The provided path is a directory, not a file. Please provide a file path!")
            return file_path
        except FileNotFoundError as e:
            print(e)
        except Exception as e:
            print("An error occurred: ", e)


def handle_option(file, option):
    print("\033[2J\033[H", end="", flush=True)
    if option == "1":
        view_macros.detect_macros(file)
    elif option == "2":
        analyse_macros.analyse_macros(file)
    elif option == "3":
        deobfuscate_macros.deobfuscate(file)
    elif option == "4":
        file.close()
        return True
    else:
        print("Invalid option! Please try again.")
    return False


def selection_menu():
    print("Select an option:")
    print("1. View macros")
    print("2. Analyse macros")
    print("3. Deobfuscate macros")
    print("4. Exit")
    return input("Enter option: ").strip()


def main():
    file = initialize()

    while True:
        option = selection_menu()
        to_exit = handle_option(file, option)
        if to_exit:
            break


if __name__ == "__main__":
    main()
