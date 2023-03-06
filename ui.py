from excel import Excel, yes_no
from os import system, name, get_terminal_size
from colorama import Fore, Style


def clear_screen() -> None:
    if name == 'nt':
        system('cls')
    else:
        system('clear')


class UI:
    def __init__(self) -> None:  # 73 cols
        self._ex = Excel()
        self.cols, self.row = get_terminal_size()
        self.running = True
        self._step = 1

    def terminal_size(self) -> None:
        self.cols, self.row = get_terminal_size()

    def welcome(self) -> None:
        to_print = '''
                      /^--^\     /^--^\     /^--^\\
                      \____/     \____/     \____/
                     /      \   /      \   /      \\
                    |        | |        | |        |
                     \__  __/   \__  __/   \__  __/      huynhdainhan
|^|^|^|^|^|^|^|^|^|^|^|^\ \^|^|^|^/ /^|^|^|^|^\ \^|^|^|^|^|^|^|^|^|^|^|^|
| | | | | | | | | | | | |\ \| | |/ /| | | | | | \ \ | | | | | | | | | | |
| | | | | | | | | | | | / / | | |\ \| | | | | |/ /| | | | | | | | | | | |
| | | | | | | | | | | | \/| | | | \/| | | | | |\/ | | | | | | | | | | | |
#########################################################################
| | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | |
| | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | | |

'''
        print(Fore.YELLOW, to_print, Style.RESET_ALL)

    def options(self, step) -> None:
        if step == 1:
            while True:
                try:
                    inp = input("Step 1: Input data folder: ").strip()
                    if not inp.endswith('/'):
                        inp = inp + '/'
                    self._ex.input_file_path = inp
                    self._step = 2
                    break
                except ValueError as error:
                    print(Fore.RED, error, Style.RESET_ALL)

        if step == 2:
            while True:
                try:
                    print(Fore.GREEN, "Input folder:",
                          self._ex.input_file_path, Style.RESET_ALL)
                    inp = input("Step 2: Output data folder: ").strip()
                    if not inp.endswith('/'):
                        inp = inp + '/'
                    self._ex.output_file_path = inp
                    self._step = 3
                    break
                except ValueError as error:
                    print(Fore.RED, error, Style.RESET_ALL)

        if step == 3:
            while True:
                try:
                    print(Fore.GREEN, "Input folder:",
                          self._ex.input_file_path, Style.RESET_ALL)
                    print(Fore.GREEN, "Output folder:",
                          self._ex.output_file_path, Style.RESET_ALL)
                    inp = input("Step 3: Name your output file (.xlsx): ").strip()
                    if not inp.endswith('.xlsx'):
                        inp = inp + '.xlsx'
                    self._ex.output_file_name = inp
                    self._step = 4
                    break
                except ValueError as error:
                    print(Fore.RED, error, Style.RESET_ALL)

        if step == 4:
            try:
                print(Fore.LIGHTGREEN_EX,
                      "ⓘ Loading data...", Style.RESET_ALL)
                self._ex.prepare_data()
                print(Fore.LIGHTGREEN_EX,
                      "ⓘ List file will be merged", Style.RESET_ALL)
                for idx, file in enumerate(self._ex.excel_file_list):
                    print(Fore.YELLOW, idx+1, '-\t', file, Style.RESET_ALL)

                opt = yes_no("Continue? (Y/n): ")

                if opt:
                    self._step = 5
                else:
                    self.running = False
            except ValueError as error:
                print(error)
                opt = yes_no("Try again? (Y/n): ")

                if opt:
                    self._step = 1
                else:
                    self.running = False
            except:
                opt = yes_no("Something went wrong. Try again? (Y/n): ")

                if opt:
                    self._step = 1
                else:
                    self.running = False

        if step == 5:
            try:
                print(Fore.LIGHTGREEN_EX,
                      "ⓘ Processing data...", Style.RESET_ALL)
                self._ex.process()
                print(Fore.LIGHTGREEN_EX,
                      "ⓘ Result file saved to:", self._ex.output_file_path+self._ex.output_file_name, Style.RESET_ALL)

                opt = yes_no("Wanna merge more files? (Y/n): ")

                if opt:
                    self._step = 1
                    self._ex.reset()
                else:
                    self.running = False
            except:
                opt = yes_no("Something went wrong. Try again? (Y/n): ")

                if opt:
                    self._step = 1
                else:
                    self.running = False

    def run(self) -> None:
        while (self.running):
            clear_screen()

            if not self.cols > 73 and not self.row > 27:
                print(Fore.BLUE, "Columns:", self.cols,
                      "x Rows:", self.row, Style.RESET_ALL)
                print(
                    Fore.YELLOW, "⚠ Your window size is too small...\n Columns must > 73 and Rows must > 27", Style.RESET_ALL)
                self.terminal_size()
                continue

            self.welcome()
            self.options(self._step)
