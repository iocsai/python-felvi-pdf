import csv
import logging
import os
import re
import shutil
import sys
from collections import OrderedDict
from datetime import datetime
from logging.handlers import TimedRotatingFileHandler

import tabula
import xlsxwriter

_VERSION = "0.1.0"
FORMATTER = logging.Formatter("%(asctime)s %(levelname)s [%(name)s] [%(message)s]")
IN_FOLDER = "in"
NICKNAMES = ("kiscsillag", "littlestar")
LOG_FOLDER = "logs"
LOG_FILE = os.path.join(LOG_FOLDER, f"output-{datetime.today().strftime('%Y-%m-%d')}.log")
OM_ID_PATTERN = re.compile(r"\d{11}")
OUT_FOLDER = "out"
SEPARATOR = " -> "
TEMP_FOLDER = "temp"
XLSX_NAME = ""


class PDFConverter:

    def __init__(self, path, school_name, class_type, pages="all", password=""):
        if not os.path.exists(TEMP_FOLDER):
            os.mkdir(TEMP_FOLDER)
        self.csv_file = os.path.join(TEMP_FOLDER, ".".join([school_name, class_type, "csv"]))
        tabula.convert_into(path, self.csv_file, pages=pages, password=password)

    @staticmethod
    def cleanup():
        shutil.rmtree(TEMP_FOLDER)


class Processing:
    _SCHOOLS = ("Fazekas", "Kossuth", "Csokonai", "Ady", "Medgyessy", "Mechwart", "Vegyipari", "TAG")

    key_to_pos = dict()
    student_dict = dict()

    def __init__(self, source):
        s = source.split(".")
        if s[0].capitalize() in self._SCHOOLS:
            temp = PDFConverter(os.path.join(IN_FOLDER, in_file), s[0], s[1])
            self.school_name = s[0]
            self.class_type = s[1]
            self.csv_file = temp.csv_file
            self.process()

    def process(self):
        getattr(self, "_".join(["process", self.school_name]))()
        pass

    def process_csokonai(self):
        raw_data = dict()
        with open(self.csv_file, 'r') as csv_file:
            logger.info(f"Processing {self.school_name}.{self.class_type}")
            csv_reader = csv.reader(csv_file, delimiter=',')
            line_count = 0
            for row in csv_reader:
                if line_count != 0 and OM_ID_PATTERN.match(row[0]) or row[0] in NICKNAMES:
                    logger.info(f"{row[0]} -> {row[1]}")
                    raw_data[row[0]] = float(row[1].replace(',', '.'))
                line_count += 1
        self.student_dict = OrderedDict(
            {k: v for k, v in sorted(raw_data.items(), key=lambda item: item[1], reverse=True)})
        self.key_to_pos = {k: pos for pos, k in enumerate(self.student_dict)}
        logger.info(f'Processed {line_count} lines.')

    def process_medgyessy(self):
        raw_data = dict()
        with open(self.csv_file, 'r') as csv_file:
            logger.info(f"Processing {self.school_name}.{self.class_type}")
            csv_reader = csv.reader(csv_file, delimiter=',')
            line_count = 0
            for row in csv_reader:
                if line_count != 0 and OM_ID_PATTERN.match(row[0]) or row[0] in NICKNAMES:
                    logger.info(f"{row[0]} -> {row[2]}")
                    raw_data[row[0]] = float(row[2].replace(',', '.')) if row[2].__contains__(',') else 0
                line_count += 1
        self.student_dict = OrderedDict(
            {k: v for k, v in sorted(raw_data.items(), key=lambda item: item[1], reverse=True)})
        self.key_to_pos = {k: pos for pos, k in enumerate(self.student_dict)}
        logger.info(f'Processed {line_count} lines.')


class XlsxExport:
    EXT = "xlsx"
    START_ROW = 3

    def __init__(self, filename, school_list):
        os.path.join(OUT_FOLDER, ".".join([filename, self.EXT]))
        workbook = xlsxwriter.Workbook(os.path.join(OUT_FOLDER, ".".join([filename, self.EXT])))
        for school in school_list:
            worksheet = workbook.add_worksheet(".".join([school.school_name, school.class_type]))
            worksheet.set_column('B:B', 12)
            worksheet.write('A1', school.school_name)
            worksheet.write('A2', school.class_type)
            row = self.START_ROW
            for k in school.student_dict:
                worksheet.write(row, 0, school.key_to_pos[k] + 1)
                worksheet.write(row, 1, k)
                worksheet.write(row, 2, school.student_dict[k])
                row += 1
        workbook.close()


def get_logger(logger_name):
    llogger = logging.getLogger(logger_name)
    llogger.setLevel(logging.DEBUG)
    llogger.addHandler(get_console_handler())
    llogger.addHandler(get_file_handler())
    llogger.propagate = False
    return llogger


def get_console_handler():
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(FORMATTER)
    return console_handler


def get_file_handler():
    file_handler = TimedRotatingFileHandler(LOG_FILE, when='midnight')
    file_handler.setFormatter(FORMATTER)
    return file_handler


def startup():
    if not os.path.exists(IN_FOLDER) or not os.listdir(IN_FOLDER):
        logger.error("Input folder empty or does no exists!")
        sys.exit(1)
    if not os.path.exists(OUT_FOLDER):
        os.mkdir(OUT_FOLDER)
    if not os.path.exists(LOG_FOLDER):
        os.mkdir(LOG_FOLDER)


if __name__ == '__main__':
    logger = get_logger("Felvi-PDF-processor")
    startup()
    schools = list()
    for in_file in os.listdir(IN_FOLDER):
        logger.info(f"Source found: {in_file}")
        schools.append(Processing(in_file))
    XlsxExport("test", schools)
    PDFConverter.cleanup()
    pass
