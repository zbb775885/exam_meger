import xlwt


class Writer:
    file_path_ = ""
    file_ = xlwt.Workbook()
    alignment = xlwt.Alignment()  # Create Alignment
    alignment.horz = xlwt.Alignment.HORZ_CENTER  # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED, HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
    alignment.vert = xlwt.Alignment.VERT_CENTER
    align_style_ = xlwt.XFStyle()
    align_style_.alignment = alignment

    def __init__(self):
        self.file_path_ = ""

    def __init__(self, file_path):
        self.file_path_ = file_path
        print("open execl:%s" % file_path)

    def __del__(self):
        if "" != self.file_path_:
            self.file_.save(self.file_path_);

    def add_excel_sheet(self, sheet_name):
        return self.file_.add_sheet(sheet_name)

    def write_excel_sheet(self, sheet, row_start, row_end, column_start, column_end, data):
        data = str(data)
        sheet.col(column_start).width = (len(data) + 1) * 500

        return sheet.write_merge(row_start, row_end, column_start, column_end, data, self.align_style_);
