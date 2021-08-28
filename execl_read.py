import xlrd


class Reader:
    file_path_ = ""
    file_ = xlrd.Book()

    def __init__(self):
        self, file_path_ = ""

    def __init__(self, file_path):
        self.file_path_ = file_path
        self.file_ = xlrd.open_workbook_xls(file_path)
        print("open execl:%s" % file_path)

    def read_excel_sheet(self,sheet_name):
        # 通过excel表格名称(rank)获取工作表
        return self.file_.sheet_by_name(sheet_name)

    def read_excel_sheet(self):
        # 通过excel表格名称(rank)获取工作表
        return self.file_.sheets()[0]

