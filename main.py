#!/usr/bin/env python3

import sys

from PyQt5.Qt import QApplication

from FetcherGui import FetcherGui

if __name__ == '__main__':
    app = QApplication(sys.argv)
    gui = FetcherGui(app)
