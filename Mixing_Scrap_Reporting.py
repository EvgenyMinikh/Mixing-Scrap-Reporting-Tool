from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.uic import *
from configparser import ConfigParser
import json
import sys

MAIN_UI_WINDOW_PATH = './main_window.ui'
CONFIG_FILENAME = './config.cfg'
SHIFT_SUPERVISORS_FILE_PATH = './shift_supervisors.csv'
INCONSISTENCY_REASONS_FILE_PATH = './list_items.json'

config = ConfigParser()
config.read_file(open(CONFIG_FILENAME, encoding="utf8"))
BUFFER_WORKBOOK_PATH = config.get('Paths', 'BUFFER_WORKBOOK_PATH')
FIRST_LINE_NUMBER = int(config.get('List Settings', 'FIRST_LINE_NUMBER'))
LAST_LINE_NUMBER = int(config.get('List Settings', 'LAST_LINE_NUMBER'))
SHIFTS = config.get('List Settings', 'SHIFTS')


def get_mixing_line_numbers_list(first, last):
    result = [str(i) for i in range(first, last + 1)]
    result.insert(0, "")
    return result


def get_shift_letters(st):
    result = [s for s in st]
    result.insert(0, "")
    return result


def get_shift_supervisors_list(file_path):
    result = []
    with open(file_path, mode='r', encoding='UTF8') as f:
        for line in f:
            result.append(line.strip())

    result.insert(0, "")
    return sorted(result)


def read_json_from_file(path):
    with open(path) as json_file:
        json_data = json.load(json_file)

    return json_data


def get_inconsistency_types_list(json_data):
    value_list = sorted(json_data.keys())
    value_list.insert(0, "")
    return value_list


class MainWindow(QMainWindow):
    def __init__(self, *args, **kwargs):
        super(MainWindow, self).__init__(*args, **kwargs)
        loadUi(MAIN_UI_WINDOW_PATH, self)
        self.set_current_date()

        # fill comboboxes with simple data
        self.comboBox_line_number.addItems(get_mixing_line_numbers_list(FIRST_LINE_NUMBER, LAST_LINE_NUMBER))
        self.comboBox_shift_number.addItems(get_shift_letters(SHIFTS))
        self.comboBox_writeoff_shift_number.addItems(get_shift_letters(SHIFTS))

        self.comboBox_blank_author.addItems(get_shift_supervisors_list(SHIFT_SUPERVISORS_FILE_PATH))

        json_data = read_json_from_file(INCONSISTENCY_REASONS_FILE_PATH)
        self.comboBox_inconsistency_type.addItems(get_inconsistency_types_list(json_data))
        self.comboBox_inconsistency_type.activated.connect(lambda: self.get_inconsistency_reason_list(json_data))

        self.setFixedSize(self.size())
        self.show()

    # fill comboboxes with current date
    def set_current_date(self):
        current_date = QDate.currentDate()
        self.dateEdit_producing_date.setDate(current_date)
        self.dateEdit_writeoff_date.setDate(current_date)

    def clean_all_forms(self):
        self.set_current_date()
        self.comboBox_line_number.setCurrentIndex(0)
        self.comboBox_shift_number.setCurrentIndex(0)
        self.comboBox_writeoff_shift_number.setCurrentIndex(0)
        self.comboBox_blank_author.setCurrentIndex(0)

    def get_inconsistency_reason_list(self, json_data):
        value_list = json_data[self.comboBox_inconsistency_type.currentText()]
        value_list.insert(0, "")
        self.comboBox_inconsistency_reason.clear()
        self.comboBox_inconsistency_reason.addItems(value_list)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_window = MainWindow()
    app.exec_()
