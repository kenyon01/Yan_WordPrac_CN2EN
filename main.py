import sys
from PyQt5.QtWidgets import QApplication, QFileDialog
from app import SpellingApp

def main():
    app = QApplication(sys.argv)

    spelling_app = SpellingApp()
    spelling_app.show()

    filename, _ = QFileDialog.getOpenFileName(spelling_app, 'Open Word File', '.', 'Excel files (*.xlsx *.xls)')
    if filename:
        spelling_app.load_words(filename)

    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
