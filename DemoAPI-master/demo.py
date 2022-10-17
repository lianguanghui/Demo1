import xlrd
import xlwt


class ReadExcel():
    def __init__(self, filename, sheetname="Sheet1"):
        self.data = xlrd.open_workbook(filename)
        self.table = self.data.sheet_by_name(sheetname)

        # 获取总行数、总列数
        self.nrows = self.table.nrows
        self.ncols = self.table.ncols
        self.keys = self.table.row_values(0)

    def read_data(self):
        list_data = []
        # keys = self.table.row_values(0)

        if self.nrows > 1:
            for col in range(1, self.nrows):
                api_dict = dict(zip(self.keys, self.table.row_values(col)))
                # print(api_dict)
                list_data.append(api_dict)
        print(list_data)
        now_apilist = []
        for data in list_data:
            url = data['url']
            method = data['method']
            params =data['params']
            headers = data['headers']
            now_apilist.append(url)
            now_apilist.append(method)
            now_apilist.append(params)
            now_apilist.append(headers)
        print(now_apilist)
        return list_data

    def write_data(self):
        datas = self.read_data()
        workbook = xlwt.Workbook(encoding='utf-8')
        sheet1 = workbook.add_sheet("Sheet1", cell_overwrite_ok=True)
        # print(self.keys)
        i = 0
        if i < self.ncols:
            for key in self.keys:
                sheet1.write(0, i, label=key)
                i += 1
        workbook.save("test.xls")


if __name__ == '__main__':
    filename = "database/DemoAPITestCase.xls"
    r = ReadExcel(filename)
    r.read_data()
    # r.write_data()
