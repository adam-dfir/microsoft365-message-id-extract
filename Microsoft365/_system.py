class System:
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    def output(type,message):
        if type == "E":
            print(System.FAIL+"[-] "+System.ENDC+message)
        elif type == "W":
            print(System.WARNING+"[!] "+System.ENDC+message)
        elif type == "S":
            print(System.OKGREEN+"[+] "+System.ENDC+message)
        else:
            print(System.OKCYAN+"[*] "+System.ENDC+message)
