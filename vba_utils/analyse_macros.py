
def analyse_macros(file):
    results = file.analyze_macros(True)
    if results is None:
        print("No VBA macros found! Exiting...")
        exit(0)
    else:
        print("The following keywords were found:")
        print("-" * 80)
        for kw_type, keyword, description in results:
            print(f"Keyword type: {kw_type}, Keyword: {keyword}, Description: {description}")
        print("-" * 80)
