import json

import execl_read
import execl_write
import gen_format
import json_reader


class ScoreAttr:
    value_ = 0
    subject_ = ""
    is_main_ = True

    # 定义分数
    def __init__(self):
        self.value_ = 0
        self.subject_ = ""
        self.is_main_ = True

    # 定义分数
    def __init__(self, value, subject, is_main):
        self.value_ = value
        self.subject_ = subject
        self.is_main_ = is_main

    # 设置分数
    def set_score(self, score):
        self.value_ = score

    # 获取分数
    def get_score(self):
        return self.value_


class StudentAttr:
    score_attr_map_ = {}
    score_total_main_ = 0
    score_total_ = 0
    class_rank_ = 0
    name_ = ""
    gender_ = ""
    class_ = ""
    number_ = int(0)


    # 定义属性类
    def __init__(self):
        self.score_attr_map_ = {}

    # 定义分数属性
    def set_score_attr(self, score_attr):
        self.score_attr_map_ = score_attr

    # 定义分数属性
    def get_score_attr(self):
        return self.score_attr_map_

    # 增加学科分数
    def set_score(self, subject, score):
        if score < 0:
            score = 0
        self.score_attr_map_[subject].set_score(score)
        # score_total = 0
        # score_total_main = 0
        # for subject, score_attr in self.score_attr_map_.items():
        #     score_total += score_attr.get_score()
        #     if True == score_attr.is_main_:
        #         score_total_main += score_attr.get_score()
        # self.score_total_ = score_total
        # self.score_total_main_ = score_total_main

    # 读取学科分数
    def get_score(self, subject):
        return self.score_attr_map_[subject].get_score()

    # 读取总分数
    def get_total_score(self):
        return self.score_total_

    # 读取主科目总分数
    def get_total_main_score(self):
        return self.score_total_main_



# 定义子类
class Student:
    student_attr_ = StudentAttr()

    # 定义学生对象
    def __init__(self, score_conf_json):
        # 读取分数属性
        json_score_conf = json_reader.read_file(score_conf_json)
        score_attr = {}
        for subject_attr in json_score_conf["score_attr"]:
            score_attr[subject_attr["subject"]] = ScoreAttr(0, subject_attr["subject"], subject_attr["is_main"])

        self.student_attr_.set_score_attr(score_attr)

    # 增加学科分数
    def add_score(self, subject, score):
        self.student_attr_.set_score(subject, score)

    # 读取学科分数
    def get_score(self, subject):
        return self.student_attr_.get_score(subject)

    # 读取总分数
    def get_total_score(self):
        return self.student_attr_.get_total_score()

    # 读取主科目总分数
    def get_total_main_score(self):
        return self.student_attr_.get_total_main_score()


if __name__ == '__main__':
    excel_files = []
    json_excel_files = json_reader.read_file("format_conf.json")
    excel_files = json_excel_files["excel_files"]
    print(excel_files)

    student_map = {}
    # 遍历配置的excel文件
    for excel_file in excel_files:
        # 读取excel文件的一个sheet
        excel_reader = execl_read.Reader(excel_file)
        sheet = excel_reader.read_excel_sheet()
        # 每一列的主题
        column_topic_map = {}
        for row in range(sheet.nrows):  # 循环读取表格内容（每次读取一行数据）
            cells = sheet.row_values(row)  # 每行数据赋值给cells
            # 第一行则提取key值
            if 0 == row:
                column = 0
                for topic in cells:
                    column_topic_map[topic] = column
                    column += 1
            else:
                student = Student("score_conf.json")
                # 读取每个学生的分数信息
                for subject in student.student_attr_.get_score_attr().keys():
                    if subject in column_topic_map:
                        student.student_attr_.set_score(subject, cells[column_topic_map[subject]])
                student_attr = student.student_attr_

                # 班级	号次		班名	 XM	年名	语数英		YSY年
                # 班级	号次	姓名	性别	语文	数学	英语	物理	化学	生物	政治	历史	地理	总分	班名	XM	年名	语数英	YSY班	YSYXM	YSY年
                student_attr.name_ = cells[column_topic_map["姓名"]]
                student_attr.gender_ = cells[column_topic_map["性别"]]
                student_attr.class_ = cells[column_topic_map["班级"]]
                student_attr.number_ = cells[column_topic_map["号次"]]
                student_attr.score_total_ = cells[column_topic_map["总分"]]
                student_attr.score_total_main_ = cells[column_topic_map["语数英"]]
                student_attr.class_rank_ = cells[column_topic_map["年名"]]
                print(student.student_attr_.number_, " 语文  ", student.get_score("语文"));
