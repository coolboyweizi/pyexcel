#!/usr/bin/env python
# coding:utf-8
"""
xrld模块的简单遍历数据

def filter: 过滤条件。索引号
"""
import xlrd,sys



class simpleXlrd:
    # 迭代量
    iter = 0

    # 字段
    indexes = []

    # 列字段与索引关系
    colIndex = {}

    # 过滤的函数集合
    filter_dict = {}

    # 修饰的函数集合
    fills_dict = {}

    def __init__(self, filename, sheet, start=0):
        if isinstance(sheet, int):
            self.data = xlrd.open_workbook(filename).sheet_by_index(sheet)
        else:
            self.data = xlrd.open_workbook(filename).sheet_by_name(sheet)
        sys.setrecursionlimit(self.data.nrows)
        # 获取field的键值对

        if start >= 0:
            for index, field in enumerate(self.data.row_values(start)):
                self.colIndex.update({field: index})

    def __fields(self, field):
        """
        字段的转换。
        1、如果字段是str格式，则去匹配索引
        2、如果字段是int格式，则去验证索引
        :param field:
        :return: int
        """
        if isinstance(field, int) and field < len(self.colIndex):
            index = field
        else:
            index = self.colIndex.get(field)
            if index is None:
                raise IndexError("索引不存在：%s" % field)
        return index

    def filter(self, field, function):
        """
        数据过滤. 内部函数只能返回bool类型
        :param field:
        :param index:
        :param function:
        :return:
        """
        index = self.__fields(field)

        funcs = self.filter_dict.get(index)

        if funcs is None:  # 没有数据
            funcs = [function]
        else:
            funcs.append(function)
        self.filter_dict[index] = funcs
        return self

    def filterWithFields(self, function, **fields):
        pass

    def fills(self, field, function: object):
        """
        填充数据，内部函数集合
        :param field:
        :param function:
        :return:
        """
        index = self.__fields(field)
        funcs = self.filter_dict.get(index)

        if not funcs:  # 没有数据
            funcs = [function]
        else:
            funcs.append(function)
        self.fills_dict[index] = funcs
        return self

    def fields(self, fields: str):
        """
        过滤字段
        :param fields:
        :return:
        """
        for field in fields.split(","):
            self.indexes.append(
                self.__fields(field)
            )
        return self

    def __next__(self):
        """
        迭代器
        :return:
        """
        line = self.data.row(self.iter)
        self.iter += 1
        if self.iter >= self.data.nrows:
            raise StopIteration

        # 过滤处理
        for index in self.filter_dict:
            funcs = self.filter_dict.get(index)
            value = line[index]
            for func in funcs:
                if not func(line):
                    return self.__next__()
        # 字段修饰处理
        for index in self.fills_dict:
            funcs = self.fills_dict.get(index)
            for func in funcs:
                line[index] = func(line[index])

        # 如果字段筛选
        if len(self.indexes) > 0:
            line = list(filter(lambda item: line.index(item) in self.indexes, line))

        return line

    def __iter__(self):
        return self
