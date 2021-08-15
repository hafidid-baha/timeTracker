from datetime import datetime
import sys


class DateHelper:

    @staticmethod
    def get_hour(date, string_format='%Y-%m-%d %H:%M:%S.%f'):
        return datetime.strptime(str(date), string_format).hour

    @staticmethod
    def get_day(date, string_format='%Y-%m-%d %H:%M:%S.%f'):
        return datetime.strptime(str(date), string_format).day

    @staticmethod
    def get_month(date, string_format='%Y-%m-%d %H:%M:%S.%f'):
        return datetime.strptime(str(date), string_format).month

    @staticmethod
    def get_year(date, string_format='%Y-%m-%d %H:%M:%S.%f'):
        return datetime.strptime(str(date), string_format).year


sys.path.append(".")
