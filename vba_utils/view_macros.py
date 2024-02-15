
def detect_macros(file):
    if file.detect_vba_macros():
        print("This document contains VBA macros.")
    else:
        print("No VBA macros found! Exiting...")
        exit(0)
