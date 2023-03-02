from interface import settings_root

MAIN_SETTINGS = {
    'exel': None,
    'word': None,
    'result folder': None,
    'sheet with information': None,
    'sheet with charts': None,
    'result file name': None,
}

def main():
    settings_root(MAIN_SETTINGS)

if __name__ == '__main__':
    main()