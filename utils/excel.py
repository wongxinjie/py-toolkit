#!/usr/bin/env python
# -*- coding: utf-8 -*-
""""
    mine.py
    ~~~~~~~~~~~~~~~~~~~~


    :author: wongxinjie
    :date created: 2019-06-25 10:03
"""
from io import BytesIO

import xlrd
import xlsxwriter


def write_excel(headers, content_matrix, save_path=None):
    """
    写excel工具函数
    :param headers: list of str, excel表头
    :param content_matrix: list of list of str, excel内容
    :param save_path: 保存路径，None 则返回二进制内容
    :return:
    """
    if save_path is None:
        output = BytesIO()
    else:
        output = save_path

    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet()

    row = 0
    header_format = workbook.add_format({"bold": True})
    for idx, text in enumerate(headers):
        worksheet.write(row, idx, text, header_format)

    row += 1
    for line in content_matrix:
        for idx, text in enumerate(line):
            worksheet.write(row, idx, text)
        row += 1

    workbook.close()
    if save_path is None:
        return output.getvalue()


def read_excel_to_dict(stream=None, excel_path=None):
    """
    读excel文件流或者文件，并把内容转换成字典
    :param stream:
    :param excel_path:
    :return: dict
        headers: list of str
        content: list of list of str
    """
    if stream is not None and isinstance(stream, bytes):
        workbook = xlrd.open_workbook(file_contents=stream)
    else:
        workbook = xlrd.open_workbook(excel_path)

    sheet = workbook.sheet_by_index(0)
    total_row_count = sheet.nrows

    payload = dict()
    headers = [r.value for r in sheet.row(0)]
    payload['headers'] = headers

    content_matrix = []
    for idx in range(1, total_row_count):
        line = [r.value for r in sheet.row(idx)]
        content_matrix.append(line)

    payload['content'] = content_matrix
    return payload
