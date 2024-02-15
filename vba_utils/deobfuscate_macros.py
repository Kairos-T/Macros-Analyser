from colorama import Fore


def deobfuscate(file):
    deobfuscated = file.reveal()
    if deobfuscated:
        print("Deobfuscated code:")
        print("-" * 80)
        print(Fore.GREEN + deobfuscated)
        print("-" * 80)
    else:
        print("No obfuscated strings found!")
