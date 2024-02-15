
def detect_macros(file):
    if file.detect_vba_macros():
        print("The following macros were found:")
        for (filename, stream_path, vba_filename, vba_code) in file.extract_macros():
            print("-" * 80)
            print("OLE stream:", stream_path)
            print("VBA filename:", vba_filename)
            print("-" * 80)
            print(vba_code)
    else:
        print("No VBA macros found! Exiting...")
        exit(0)
