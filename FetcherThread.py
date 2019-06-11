import os
from re import search
from time import time_ns

from PyQt5.QtCore import QThread, pyqtSignal
from pandas import ExcelFile, ExcelWriter, DataFrame


# noinspection PyArgumentList
class FetcherThread(QThread):
    # Declare Signals
    begin = pyqtSignal(int)  # Send progress maximum
    update = pyqtSignal(str, int)  # Send label and progress
    get_save = pyqtSignal(str)  # Request save path

    def __init__(self, path):
        super().__init__()
        self.path = path
        self.save = None
        self.canceled = False
        self.errs = 0

    # noinspection PyTypeChecker
    @staticmethod
    def get_subdirs(d: str) -> [str]:
        return list(filter(os.path.isdir, [f'{d}/{f}' for f in os.listdir(d)]))

    @staticmethod
    def get_files(d: str) -> [str]:
        ext = ('.xls', '.xlsx')  # Excel extensions
        files = list(filter(lambda f: f.endswith(ext), [
            f'{d}/{f}' for f in os.listdir(d) if not f.startswith('~$')]))

        # TODO: if not files, check if there are sub directories in the student's folder

        files.sort(key=os.path.getmtime, reverse=True)
        return files

    @staticmethod
    def get_name(df: DataFrame) -> str:
        for i in range(10):  # Check rows 1 through 10
            if df.shape[0] >= i + 1 and df.shape[1] >= 0:
                cell = df.iat[i, 0]  # cell at column A
                if isinstance(cell, str) and cell.startswith('Student'):
                    return cell[13:].strip()

    @staticmethod
    def get_id(df: DataFrame) -> str:
        for i in range(10):  # Check rows 1 through 10
            if df.shape[0] >= i + 1 and df.shape[1] >= 5:
                cell = df.iat[i, 5]  # cell at column F
                if isinstance(cell, float) and 9999999 < cell <= 999999999:
                    return str(cell)
                elif isinstance(cell, str):
                    cell = cell.strip()
                    if len(cell) == 8 and cell.isdigit():
                        return '0' + cell
                    elif len(cell) == 9 and cell.isdigit():
                        return cell

    @staticmethod
    def regex_id(s: str) -> str:
        match = search('[0-9]{8,9}', s)
        if match is not None:
            reg = match.group()
            if len(reg) == 8:
                return '0' + reg
            else:
                return reg

    # noinspection PyBroadException
    def get_info(self, d: str) -> (str, str):
        files = FetcherThread.get_files(d)
        wsu_id, name = '', ''

        for file in files:
            try:
                with ExcelFile(file) as book:
                    for sheet_name in book.sheet_names:
                        sheet = book.parse(sheet_name=sheet_name, header=None)
                        try:
                            if not wsu_id:
                                wsu_id = FetcherThread.get_id(sheet)
                        except Exception as e:  # Allow other sheets in book to be tried
                            print(f'Caught exception on {file} \n {e}')
                            self.errs += 1
                        try:
                            if not name:
                                name = FetcherThread.get_name(sheet)
                        except Exception as e:  # Allow other sheets in book to be tried
                            print(f'Caught exception on {file} \n {e}')
                            self.errs += 1
                        if wsu_id and name:
                            return wsu_id, name
            except Exception as e:
                print(f'Caught exception on {file} \n {e}')
                self.errs += 1
        else:
            if not wsu_id:
                # See if we can regex an ID from the folder name
                wsu_id = FetcherThread.regex_id(d)

        return wsu_id, name

    def run(self):
        start = time_ns()
        folders = self.get_subdirs(self.path)
        ids, names = [], []

        self.begin.emit(len(folders))
        for count, folder in enumerate(folders):
            self.update.emit(folder, count)
            wsu_id, name = self.get_info(folder)
            ids.append(wsu_id)
            names.append(name)
            if self.canceled:
                return

        elapsed = (time_ns() - start) / (10 ** 9)

        # Request save path from GUI thread
        self.update.emit('Saving to Excel', len(folders))
        self.get_save.emit(self.path.split('/')[-1])

        # TODO: If ID and Name are the same AND NOT NONE over two folders, remove the duplicates
        # TODO: do it here while the user is selecting a save path

        # Wait for GUI thread to send save path
        while self.save is None:
            self.yieldCurrentThread()

        if self.save:
            try:
                with ExcelWriter(self.save, engine='xlsxwriter') as writer:
                    df = DataFrame({'ID': ids, 'Name': names, 'Folder': folders})
                    df.to_excel(writer, self.path.split('/')[-1][:31], index=False)
                    # FIXME: Overwriting a file when the file is already open throws error
                    writer.save()
            except Exception as e:
                self.errs += 1
                print(e)

        print(f'Done! Processed {len(folders)} folders in {elapsed} seconds with {self.errs} errors.')
