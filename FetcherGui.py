from PyQt5.Qt import QApplication, QMainWindow, QProgressDialog
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QFileDialog, QLabel

from FetcherThread import FetcherThread


# noinspection PyUnresolvedReferences, PyArgumentList
class FetcherGui(QMainWindow):
    def __init__(self, app: QApplication):
        super().__init__()

        # Get top directory
        self.file_viewer = QFileDialog()
        path = self.file_viewer.getExistingDirectory()

        # Setup worker thread
        self.worker = FetcherThread(path)
        self.worker.begin.connect(self.start)
        self.worker.update.connect(self.update_progress)
        self.worker.get_save.connect(self.save)
        self.worker.finished.connect(self.done)

        # Setup progress widget
        self.progress = QProgressDialog()
        self.progress.setWindowTitle('ID Fetcher')
        self.progress.setFixedWidth(450)
        self.progress.canceled.connect(self.cancel)
        self.progress.setAutoClose(False)
        self.progress.setAutoReset(False)

        label = QLabel()
        label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.progress.setLabel(label)

        # Start worker thread
        if path:
            self.worker.start()
            app.exec()
        else:
            self.progress.close()
            self.close()

    def save(self, name: str):
        self.worker.save, *_ = self.file_viewer.getSaveFileName(
            directory=name, filter='Excel Files (*.xlsx)')

    def cancel(self):
        self.worker.canceled = True

    def start(self, maximum: int):
        self.progress.setMaximum(maximum)
        self.progress.show()

    def update_progress(self, label: str, value: int):
        if '/' in label:
            drive, *_, folder = label.split('/')
            self.progress.setLabelText(f'{drive}/.../{folder}')
        else:
            self.progress.setLabelText(label)
        self.progress.setValue(value)

    def done(self):
        self.progress.close()
        self.close()
