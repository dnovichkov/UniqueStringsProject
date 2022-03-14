import argparse

import loguru
import xlrd
import xlsxwriter
from fuzzywuzzy import fuzz


def contains_number(string: str) -> bool:
    return any(map(str.isdigit, string))


def transform(src_filename: str = 'Пример.xlsx', res_filename: str = 'Результат.xlsx'):
    xl_workbook = xlrd.open_workbook(src_filename)
    sheet_names = xl_workbook.sheet_names()
    if len(sheet_names) != 1:
        loguru.logger.warning(f'Больше одного листа в файле {src_filename}')
        xl_workbook.release_resources()
        return

    neccessary_sheet_name = sheet_names[0]
    xl_sheet = xl_workbook.sheet_by_name(neccessary_sheet_name)
    rows_count = xl_sheet.nrows
    loguru.logger.debug(f'rows_count = {rows_count}')

    data_range = range(1, rows_count)

    vals_1 = []
    vals_2 = []

    for row_idx in data_range:
        val_1 = str(xl_sheet.cell(row_idx, 1).value)
        vals_1.append(val_1)

        val_2 = str(xl_sheet.cell(row_idx, 2).value)
        vals_2.append(val_2)

    xl_workbook.release_resources()

    try:
        wr_workbook = xlsxwriter.Workbook(res_filename)
    except xlsxwriter.exceptions.FileCreateError as ex:
        loguru.logger.error(f'Не удалось создать файл {res_filename}: {ex} - возможно, он уже открыт')
        return
    worksheet = wr_workbook.add_worksheet()
    worksheet.set_column('A:A', 14)
    cell_format = wr_workbook.add_format()

    for row_count, val in enumerate(vals_1):
        new_val = val
        for ext_val in vals_2:
            delta = fuzz.WRatio(val, ext_val)

            is_good_delta = delta > 97 and not contains_number(val)
            if val and (val in ext_val or is_good_delta):
                loguru.logger.debug(f'delta between {val} and {ext_val} = {delta}')
                vals_1[vals_1.index(val)] = ext_val
                new_val = ext_val
                loguru.logger.debug(f'Меняем {val} на {ext_val}')
                break
        loguru.logger.debug(f'Записываем {new_val}')
        worksheet.write(row_count, 0, new_val, cell_format)
    try:
        wr_workbook.close()
    except xlsxwriter.exceptions.FileCreateError as ex:
        loguru.logger.error(f'Не удалось создать файл {res_filename}: {ex} - возможно, он уже открыт')
        return


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Пытается преобразовать значения из одного столбца в другой, создает файл с результатом')
    parser.add_argument("-src", dest="src", default='Пример.xlsx', help="Файл, откуда читаем данные")
    parser.add_argument("-dest", dest="dest", default='Результат.xlsx', help="Файл, куда записываем результат")
    args = parser.parse_args()
    src_filename = args.src
    dest_filename = args.dest
    transform(src_filename, dest_filename)

