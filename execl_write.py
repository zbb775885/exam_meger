import xlwt


class Writer:
    file_path_ = ""
    file_ = xlwt.Workbook()
    alignment = xlwt.Alignment()  # Create Alignment
    alignment.horz = xlwt.Alignment.HORZ_CENTER  # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED, HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
    alignment.vert = xlwt.Alignment.VERT_CENTER

    # 设置边框
    borders = xlwt.Borders()
    # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
    # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
    borders.left = 1
    borders.right = 1
    borders.top = 1
    borders.bottom = 1

    font = xlwt.Font()
    # 字体类型
    font.name = 'name Arial'
    # 字体颜色
    font.colour_index = 4
    # 字体加粗
    font.bold = True

    align_style_ = xlwt.XFStyle()
    align_style_.alignment = alignment
    align_style_.borders = borders

    font_align_style_ = xlwt.XFStyle()
    font_align_style_.alignment = alignment
    font_align_style_.font = font
    font_align_style_.borders = borders

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
        sheet.col(column_start).width = (len(data) + 1) * 450

        return sheet.write_merge(row_start, row_end, column_start, column_end, data, self.align_style_);

    def write_excel_sheet_bold(self, sheet, row_start, row_end, column_start, column_end, data):
        data = str(data)
        sheet.col(column_start).width = (len(data) + 1) * 450

        return sheet.write_merge(row_start, row_end, column_start, column_end, data, self.font_align_style_);