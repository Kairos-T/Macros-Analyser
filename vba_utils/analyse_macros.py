from colorama import Fore, Style
from tabulate import tabulate


def analyse_macros(file):
    results = file.analyze_macros(True)
    if results:
        print("The following keywords were found:")
        print("-" * 80)

        headers = [f"{Fore.BLUE}Keyword type", f"{Fore.BLUE}Keyword", f"{Fore.BLUE}Description"]

        data = [
            [f"{Fore.RED}{kw_type}", f"{Fore.RED}{keyword}", f"{Style.RESET_ALL}{description}"]
            for kw_type, keyword, description in results
        ]
        table = tabulate(data, headers=headers, tablefmt="plain")
        print(table)
        print("-" * 80)

        print("Summary")
        print("-" * 80)
        summary_data = [
            [f"{Fore.RED}AutoExec keywords{Style.RESET_ALL}", file.nb_autoexec],
            [f"{Fore.RED}Suspicious keywords{Style.RESET_ALL}", file.nb_suspicious],
            [f"{Fore.RED}IOCs{Style.RESET_ALL}", file.nb_iocs],
            [f"{Fore.RED}Hex obfuscated strings{Style.RESET_ALL}", file.nb_hexstrings],
            [f"{Fore.RED}Base64 obfuscated strings{Style.RESET_ALL}", file.nb_base64strings],
            [f"{Fore.RED}Dridex obfuscated strings{Style.RESET_ALL}", file.nb_dridexstrings],
            [f"{Fore.RED}VBA obfuscated strings{Style.RESET_ALL}", file.nb_vbastrings]
        ]

        summary_table = tabulate(summary_data, headers=[f"{Fore.BLUE}Category{Style.RESET_ALL}",
                                                        f"{Fore.BLUE}Count{Style.RESET_ALL}"], tablefmt="fancy_grid")
        print(summary_table)

    else:
        print(Fore.GREEN + "No IOCs found!\n")
