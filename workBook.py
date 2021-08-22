import string
import xlsxwriter
import sys
from dateHelper import DateHelper
from database import Database
from datetime import datetime
from collections import Counter


class WorkBook:
    letters = list(string.ascii_uppercase)
    numbers = [i for i in range(10)]

    def __init__(self, path="test.xlsx"):
        self.workbook = xlsxwriter.Workbook(path)
        self.worksheet = self.workbook.add_worksheet()
        self.create_default_cells()
        self.database = Database()
        self.tasks_month = list(map(lambda x: DateHelper.get_month(str(x[2])), self.database.get_all_tasks()))
        self.tasks_year = list(map(lambda x: DateHelper.get_year(str(x[2])), self.database.get_all_tasks()))
        self.tasks_days = list(map(lambda x: DateHelper.get_day(str(x[2])), self.database.get_all_tasks()))
        self.tasks = list(self.database.get_all_tasks())
        self.database.close()
        self.tasks_date = []
        self.create_default_cells()
        # self.create_date_cells(self.tasks_month, self.tasks_year, self.tasks_days)
        self.tasks = self.change_tasks_format()
        self.add_tasks_to_worksheet()
        try:
            self.workbook.close()
        except xlsxwriter.exceptions.FileCreateError:
            print("we can't close xlsx file : Permission denied")

    def create_default_cells(self):
        hours = []
        cell_format = self.workbook.add_format({'bold': True, 'font_color': 'black', 'bg_color': '#8ECAE6',
                                                'align': 'center'})
        cell_format.set_shrink()
        for index in enumerate(range(1, 25), start=4):
            hours.append(str(index[0]).zfill(2) if index[0] < 24 else str(index[0] - 24).zfill(2))

        for index, h in enumerate(hours):
            self.worksheet.write(index + 2, 0, int(h), cell_format)

    @staticmethod
    def get_month_color(month):
        switcher = {
            1: "#fec5bb",
            2: "#fae1dd",
            3: "#d8e2dc",
            4: "#ece4db",
            5: "#ffe5d9",
            6: "#fec89a",
            7: "#e5e5e5",
            8: "#bdb2ff",
            9: "#ccd5ae",
            10: "#e5e5e5",
            11: "#ffffff",
            12: "#e9edc9",
        }
        return switcher.get(month, "Invalid month")

    def create_date_cells(self, tasks_month, tasks_year, tasks_days):
        cell_format = self.workbook.add_format({'bg_color': '#FFD7BA', 'bold': 'center', 'align': 'center'})
        days_cell_format = self.workbook.add_format({'bold': True, 'font_color': 'black', 'align': 'center'})

        # self.worksheet.merge_range(0, 1, 0, 4, 'this cells are merged', cell_format)
        self.tasks_date = []
        for tm, ty in zip(tasks_month, tasks_year):
            self.tasks_date.append(str(tm) + "/" + str(ty))

        counter = dict(Counter(self.tasks_date))
        col = 1
        for i in counter:
            date_cell_format = self.workbook.add_format({'bold': True, 'font_color': 'black', 'bg_color': '#f8edeb',
                                                         'align': 'center'})
            date_cell_format.set_border(1)
            date_cell_format.set_bg_color(WorkBook.get_month_color(int(str(i).split('/')[0])))
            if int(counter.get(i)) > 1:
                self.worksheet.merge_range(0, col, 0, int(counter.get(i)) + col - 1, str(i), date_cell_format)
            else:
                self.worksheet.write(0, col, str(i), date_cell_format)
            col += int(counter.get(i))

        for index, d in enumerate(tasks_days):
            self.worksheet.write(1, index + 1, int(d), cell_format)

        # for row in self.tasks:
        #     print(datetime.strptime(row[2], '%Y-%m-%d %H:%M:%S.%f').day)

    def add_tasks_to_worksheet(self):
        col = 0
        for t in self.tasks:
            date_cell_format = self.workbook.add_format({'bold': True, 'font_color': 'black', 'align': 'center'})
            date_cell_format.set_border(1)
            date_cell_format.set_bg_color(WorkBook.get_month_color(int(str(t[0]).split('/')[1])))
            if len(t[1]) > 1:
                if len(list(dict.fromkeys(t[1]))) > 1:
                    self.worksheet.merge_range(0, col + 1, 0, col + len(list(dict.fromkeys(t[1]))), str(t[0]),
                                               date_cell_format)
                else:
                    self.worksheet.write(0, col + 1, str(t[0]), date_cell_format)
                days_cell_format = self.workbook.add_format({'bold': True, 'font_color': 'black', 'align': 'center'})
                task_cell_format = self.workbook.add_format({'font_color': 'white', 'bg_color': '#52b788',
                                                             'align': 'center'})
                column = col
                for index, d in enumerate(list(dict.fromkeys(t[1])), start=column):
                    # adding days
                    self.worksheet.write(1, index + 1, d, days_cell_format)
                    print(index)

                for index, d in enumerate(t[1]):
                    # add tasks for each day using the hour and date info
                    h = t[2][index]

                    if column < col + len(t[1]):
                        if (index > 0 and t[1][index - 1] != d) or (
                                index == 0 and len(t[1]) > 1 and t[1][index + 1] != d):
                            column += 1

                    if h in range(4, 24):
                        self.worksheet.write(h - 2, column+1, str(t[3][index]), task_cell_format)
                    else:
                        self.worksheet.write(h + 22, column+1, str(t[3][index]), task_cell_format)

            else:
                self.worksheet.write(0, col + len(t[1]), str(t[0]), date_cell_format)
                self.worksheet.write(1, col + len(t[1]), t[1][0], days_cell_format)
            col += len(list(dict.fromkeys(t[1])))

    def change_tasks_format(self):
        data = []
        for t in self.tasks:
            task_info = [str(DateHelper.get_year(t[2])) + "/" + str(DateHelper.get_month(t[2])),
                         [DateHelper.get_day(t[2])],
                         [DateHelper.get_hour(t[2])],
                         [t[1]]
                         ]
            if len(data) >= 1:
                last_task_info = data[len(data) - 1]
                if task_info[0] == last_task_info[0]:
                    index = -1
                    if task_info[1] in last_task_info[1]:
                        index = last_task_info.index(task_info[1])
                    else:
                        last_task_info[1].append(task_info[1].pop(0))
                    last_task_info[3].append(task_info[3].pop(0))
                    last_task_info[2].append(task_info[2].pop(0))
                    continue
            data.append(task_info)
        return data


sys.path.append(".")
