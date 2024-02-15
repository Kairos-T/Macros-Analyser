from colorama import Fore


def detect_macros(file):
    if file.detect_vba_macros():
        print()
        print("The following macros were found:")
        for (filename, stream_path, vba_filename, vba_code) in file.extract_macros():
            print("-" * 80)
            print("OLE stream:", Fore.BLUE + stream_path)
            print("VBA filename:", Fore.BLUE + vba_filename)
            print("-" * 80)
            print(Fore.GREEN + vba_code)
    else:
        print("No VBA macros found! Exiting...")
        exit(0)
