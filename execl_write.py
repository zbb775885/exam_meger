import xlwt

def write_excel(file_path):
    # 创建excel文件
    file = xlwt.Workbook()
    print("open execl:%s " % file_path)

    # 增加一个sheet表
    sheet = file.add_sheet("汇总成绩")
    sheet.col(0).width = 200 * 30

    # 保存成excel文件
    file.save(file_path)
