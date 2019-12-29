#!/usr/bin/env python
# cording:utf-8
import csv
import os

import sys

import xlrd


class HappyTransform:
    def __init__(self, path):
        self.__reader(path)

    def __reader(self, input_path):
        if os.path.isdir(input_path):
            self.__read_dir(input_path, os.path.join(input_path, 'Contents'))

        if os.path.isfile(input_path):
            self.__read_file(input_path, os.path.join(os.path.split(input_path)[0], 'Contents'))

    def __read_dir(self, input_path, output_path):
        files = os.listdir(input_path)
        for one in files:
            one_in_path = os.path.join(input_path, one)
            one_out_path = os.path.join(output_path, one)
            if os.path.isdir(one_in_path):
                self.__read_dir(one_in_path, one_out_path)

            if os.path.isfile(one_in_path):
                self.__read_file(one_in_path, output_path)

    def __read_file(self, input_path, output_path):
        suffix_check = os.path.splitext(input_path)[1].lower()
        if suffix_check.endswith('.xls') or suffix_check.endswith('.xlsx'):
            self.__file_content(input_path, output_path)
        else:
            print('Please check out the file"', input_path, '", weather it is endWith ".xls/.xlsx".')

    def __file_content(self, single_file, target_file):
        sheet_content = {}
        work_data = xlrd.open_workbook(single_file)
        sheet_names = work_data.sheet_names()

        for sheet in sheet_names:
            content = []
            sheet_data = work_data.sheet_by_name(sheet)
            row_num = sheet_data.nrows

            for i_row in range(row_num):
                row_line = sheet_data.row_values(i_row)
                content.append(row_line)
            sheet_content[sheet] = content

        self._main_processing_(single_file, target_file, sheet_content)

    def _main_processing_(self, path, out_file_path, records):
        split_path = os.path.split(path)
        if not os.path.exists(out_file_path):
            os.makedirs(out_file_path)

        for key in records:
            output_path = os.path.join(out_file_path, '[Transformed]--' + os.path.splitext(split_path[1])[0] + '(' + key + ').csv')
            self.__writer(output_path, records[key])

    def __writer(self, csv_path, content_list):
        if os.path.exists(csv_path):
            os.remove(csv_path)

        csv_writer = csv.writer(open(csv_path, 'w', newline=''))
        csv_writer.writerows(content_list)


if __name__ == '__main__':
    HappyTransform(sys.argv[1])
