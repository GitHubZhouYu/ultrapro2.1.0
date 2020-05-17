"""设计好一个类,读取数据流，并设计多种不同方法，
应对多种数据提取要求，实现可以扩展，低耦合，高
复用
"""
import xlrd as xd
from datetime import timedelta, datetime
from foo.processing.constant import *


class ElectData(object):
    """
    该类提供电化学数据模板
    """
    def __init__(self, filepath, starttime):
        self.__filepath = filepath
        self.book_data = xd.open_workbook(filepath,
                                          on_demand=True,
                                          ragged_rows=True)
        self.sheet_names = self.book_data.sheet_names()
        self.starttime = starttime
        self._cap = None

    def __len__(self):
        return self.book_data.nsheets

    def __del__(self):
        self.book_data.release_resources()


    @property
    def cap(self):
        """"""
        return self._cap

    @cap.setter
    def cap(self, value):
        self._cap = value

    def get_cap(self, timegap=None, cycle=1):
        """
        :param timegap: 如果给定，则以该参数作为间隔时间
        :param cycle: 如果给定，则取出对应循环号数据
        """

        for sheetname in self.sheet_names:
            if sheetname.startswith("Detail_"):
                sheet = self.book_data.sheet_by_name(sheetname)
                for j in range(1, sheet.nrows):
                    cap = sheet.cell_value(j, COLCAP)
                    cyc = sheet.cell_value(j, COLCYC)
                    time_num = sheet.cell_value(j, COLTIME)
                    time_date = xd.xldate.xldate_as_datetime(time_num, 0)
                    if not timegap:
                        if isinstance(cycle, int) and cyc == cycle:
                            yield cap
                    else:
                        delta = timedelta(seconds=timegap)
                        if self.starttime + delta == time_date:
                            yield cap
                            self.starttime += delta


        pass


def run():
    starttime = datetime(2020,4,22,17,18,40)
    data = ElectData(r'G:\项目\ATLpackSOC\LFP体系\ATL 2020-4-23\1.xls', starttime)
    print(len(data))
    for item in data.get_cap(timegap=100, cycle=2):
        print(item)
    del data

run()




