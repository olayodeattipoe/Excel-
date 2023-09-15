from openpyxl import load_workbook


class PefCalc:

    def __init__(self, month, file_name,):
        self.month = month
        self.file_name = file_name

    def calc(self):

        start = int(input("start: "))
        stop = int(input("stop: "))

        new_value = []
        value = []
        wb = load_workbook(self.file_name)
        sheet1 = wb['Table 1']

        for _num in range(start, stop + 1):
            _func = 'F' + str(_num)
            value.append(sheet1[_func].value)
        for item in value:
            if item is not None:
                new_value.append(item)
        print(f'The Withdrawals Amount for the month {self.month} is Ghc {sum(new_value)}')

        for _num in range(start, stop + 1):
            _func = "G" + str(_num)
            value.append(sheet1[_func].value)
        for item in value:
            if item is not None:
                new_value.append(item)
        print(f'The Lodgement Amount for the month {self.month} is Ghc {sum(new_value)}')

