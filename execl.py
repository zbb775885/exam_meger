import json

import execl_read
import execl_write
import gen_format
import json_reader
import copy


class ScoreAttr:
    # 定义分数
    def __init__(self):
        self.value_ = 0
        self.subject_ = ""
        self.is_main_ = True
        self.rank_=0
        self.assign_value_ = 0
        self.level=""
        self.level_map = {}
        self.rank_name_ =""
        self.total_students_ = 0

    # 定义分数
    def __init__(self, value, subject, is_main,rank_name, total_students):
        self.value_ = value
        self.subject_ = subject
        self.is_main_ = is_main
        self.rank_ = 0
        self.assign_value_ = 0
        self.level = ""
        self.level_map = {}
        self.rank_name_ = rank_name
        self.total_students_ = total_students

    # 设置分数
    def set_score(self, score):
        self.value_ = score
        for level,range in self.level_map:
            if score >= range[0] and score <= range[1]:
                self.level = level

    # 获取分数
    def get_score(self):
        return self.value_

    # 设置赋分
    def set_assign_score(self, score):
        self.assign_value_ = score

    # 获取赋分
    def get_assign_score(self):
        return self.assign_value_


class ExamAttr:
    # 定义属性类
    def __init__(self):
        self.score_attr_map_ = {}
        self.score_total_main_ = 0
        self.score_total_main_rank_ = 0
        self.score_total_not_main_ = 0
        self.score_total_not_main_rank_ = 0
        self.score_total_ = 0
        self.score_total_class_rank_ = 0

    # 定义分数属性
    def set_score_attr(self, score_attr):
        self.score_attr_map_ = copy.deepcopy(score_attr)

    # 定义分数属性
    def get_score_attr(self):
        return self.score_attr_map_

    # 增加学科分数
    def set_score(self, subject, score):
       # if score < 0:
       #     score = 0
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
    # 定义学生对象
    def __init__(self, score_conf_json):
        self.score_attr_map_save_ = {}
        self.exam_attr_map_ = {}
        self.name_ = ""
        self.gender_ = ""
        self.class_ = int(0)
        self.number_ = int(0)
        # 读取分数属性
        json_score_conf = json_reader.read_file(score_conf_json)
        for subject_attr in json_score_conf["score_attr"]:
            self.score_attr_map_save_[subject_attr["subject"]] = ScoreAttr(0, subject_attr["subject"],
                                                                           subject_attr["is_main"],
                                                                           subject_attr["rank_name"],
                                                                           subject_attr["total_students"])

    # 增加单次考试信息
    def add_exam_attr(self, exam_name, exam_attr):
        self.exam_attr_map_[exam_name] = copy.deepcopy(exam_attr)

    # 获取单次考试信息
    def get_exam_attr(self, exam_name):
        return self.exam_attr_map_[exam_name]


def read_student_exam_infos(student_map, excel_files):
    for excel_file in excel_files:
        # 读取excel文件的一个sheet
        excel_reader = execl_read.Reader(excel_file)
        sheet = excel_reader.read_excel_sheet()
        # 每一列的主题
        column_topic_map = {}
        count = 0
        score_rank={}
        for row in range(sheet.nrows):  # 循环读取表格内容（每次读取一行数据）
            cells = sheet.row_values(row)  # 每行数据赋值给cells
            # 第一行则提取key值
            if 0 == row:
                column = 0
                for topic in cells:
                    column_topic_map[topic] = column
                    column += 1
            else:
                if not cells[column_topic_map["姓名"]] in student_map.keys():
                    student_map[cells[column_topic_map["姓名"]]] = copy.deepcopy(Student("./score_conf.json"))
                else:
                    student = ""
                student = student_map[cells[column_topic_map["姓名"]]]

                # 读取每个学生的分数信息
                exam_attr = ExamAttr()
                exam_attr.set_score_attr(student.score_attr_map_save_)
                for subject in student.score_attr_map_save_.keys():
                    if subject in column_topic_map:
                        if "" != cells[column_topic_map[subject]]:
                            exam_attr.set_score(subject, cells[column_topic_map[subject]])
                        else:
                            exam_attr.set_score(subject, -10)
                        #将所有的分数放入对应学科的排名表
                        # if not subject in score_rank.keys():
                        #     score_rank[subject]=[]
                        # score = cells[column_topic_map[subject]]
                        # if "" == cells[column_topic_map[subject]] or score < 0:
                        #     score = -10
                        # score_rank[subject].append(score)

                        if exam_attr.get_score_attr()[subject].is_main_ == False:
                            if "" != cells[column_topic_map[subject + "赋分"]]:
                                exam_attr.get_score_attr()[subject].set_assign_score(cells[column_topic_map[subject + "赋分"]])
                            else:
                                exam_attr.get_score_attr()[subject].set_assign_score(-1)
                            #print(int(column_topic_map[subject + "赋分"]))
                        #print(int(column_topic_map[exam_attr.get_score_attr()[subject].rank_name_]))
                        if "" != cells[column_topic_map[exam_attr.get_score_attr()[subject].rank_name_]]:
                            exam_attr.get_score_attr()[subject].rank_ = int(cells[column_topic_map[exam_attr.get_score_attr()[subject].rank_name_]])
                        else:
                            exam_attr.get_score_attr()[subject].rank_=-1


                exam_attr.score_total_ = cells[column_topic_map["总分"]]
                exam_attr.score_total_main_ = cells[column_topic_map["语数英"]]
                exam_attr.score_total_main_rank_ = cells[column_topic_map["语数英名"]]
                exam_attr.score_total_class_rank_ = cells[column_topic_map["年名"]]
                exam_attr.score_total_not_main_ = cells[column_topic_map["7选3"]]
                exam_attr.score_total_not_main_rank_ = cells[column_topic_map["7选3名"]]
                exam_name = excel_file.split(".")[0]
                last_exam_name = exam_name
                student.add_exam_attr(exam_name, exam_attr)
                # 班级	号次		班名	 XM	年名	语数英		YSY年
                # 班级	号次	姓名	性别	语文	数学	英语	物理	化学	生物	政治	历史	地理	总分	班名	XM	年名	语数英	YSY班	YSYXM	YSY年
                if "姓名" in column_topic_map.keys():
                    student.name_ = cells[column_topic_map["姓名"]]
                else:
                    student.name_ = "未知"

                if "性别" in column_topic_map.keys() :
                    student.gender_ = cells[column_topic_map["性别"]]
                else:
                    student.gender_ = "未知"

                if "班级" in column_topic_map.keys():
                    student.class_ = cells[column_topic_map["班级"]]
                else:
                    student.class_ = "未知"

                if "号次" in column_topic_map.keys():
                    student.number_ = cells[column_topic_map["号次"]]
                else:
                    student.number_ = "未知"

                # print(student_map[student.name_].number_, " 111语文  ", exam_name, " ",
                #       student_map[student.name_].get_exam_attr(exam_name).get_score("语文"),
                #       len(student_map[student.name_].get_exam_attr(exam_name).score_attr_map_));
                # print(student.number_, " 语文  ", exam_name, " ", student.get_exam_attr(exam_name).get_score("语文"));
        #对本次考试的各个学科做排名并将名字写入学生成绩中
        # for rank in score_rank.values():
        #     #print(rank)
        #     rank.sort()
        #     rank.reverse()
        #
        # exam_name = excel_file.split(".")[0]
        # for student in student_map.values():
        #     for subject in score_rank.keys():
        #         if exam_name in student.exam_attr_map_.keys():
        #             score = student.exam_attr_map_[exam_name].score_attr_map_[subject].get_score()
        #             student.exam_attr_map_[exam_name].score_attr_map_[subject].rank_ = score_rank[subject].index(score) + 1
        #         #else:
        #             #student.exam_attr_map_[exam_name].score_attr_map_[subject].rank_ = len(student_map)
        #           #  #score_rank[subject].index(score)

def write_student_exam_infos_to_excel(student_map, title):
    save_file = title + ".xls"
    writer = execl_write.Writer(save_file)
    sheet = writer.add_excel_sheet(title)

    # print(student_map.number_, " 语文  ", "高一上期中", " ", student.get_exam_attr("高一上期中").get_score("语文"));
    row = 0
    student_map_sorted = sorted(student_map.items(), key=lambda student_item: student_item[1].class_)
    for student_pair in student_map_sorted:
        student = student_pair[1]
        #print(student.number_, " 语文  ", "高一上期中", " ", student.get_exam_attr("高一上期中").get_score("语文"));
        column = 0;
        main_cnt = 0
        no_main_cnt = 0
        for subject in student.score_attr_map_save_.values():
            if subject.is_main_ == True:
                main_cnt+=1
            else:
                no_main_cnt+=1
        column_len = main_cnt * 2 + no_main_cnt * 3+ 9
        writer.write_excel_sheet(sheet, row, row, 0, column_len - 1, title)

        row += 1
        writer.write_excel_sheet(sheet, row, row + 1, column, column, "班级")
        writer.write_excel_sheet(sheet, row, row + 1, column + 1, column + 1, "姓名");
        column += 2
        for key in student.score_attr_map_save_.keys():
            if student.score_attr_map_save_[key].is_main_ == True:
                writer.write_excel_sheet(sheet, row, row, column, column + 1, key + "/" + str(student.score_attr_map_save_[key].total_students_));
                writer.write_excel_sheet(sheet, row + 1, row + 1, column, column, "成绩");
                writer.write_excel_sheet(sheet, row + 1, row + 1, column + 1, column + 1, "名次");
                column += 2
            else:
                writer.write_excel_sheet(sheet, row, row, column, column + 2, key + "/" + str(student.score_attr_map_save_[key].total_students_));
                writer.write_excel_sheet(sheet, row + 1, row + 1, column, column, "成绩");
                writer.write_excel_sheet(sheet, row + 1, row + 1, column + 1, column + 1, "名次");
                writer.write_excel_sheet(sheet, row + 1, row + 1, column + 2, column + 2, "赋分");
                column += 3

        writer.write_excel_sheet(sheet, row, row, column, column + 1, "总分");
        writer.write_excel_sheet(sheet, row + 1, row + 1, column, column, "成绩");
        writer.write_excel_sheet(sheet, row + 1, row + 1, column + 1, column + 1, "名次");
        column += 2

        writer.write_excel_sheet(sheet, row, row, column, column + 1, "语数英");
        writer.write_excel_sheet(sheet, row + 1, row + 1, column, column, "成绩");
        writer.write_excel_sheet(sheet, row + 1, row + 1, column + 1, column + 1, "名次");
        column += 2

        writer.write_excel_sheet(sheet, row, row, column, column + 1, "7选3");
        writer.write_excel_sheet(sheet, row + 1, row + 1, column, column, "成绩");
        writer.write_excel_sheet(sheet, row + 1, row + 1, column + 1, column + 1, "名次");
        column += 2

        writer.write_excel_sheet(sheet, row, row + 1, column, column, "考试")

        row += 2

        for exam_name in student.exam_attr_map_.keys():
            # 开始写成绩
            exam_attr = student.exam_attr_map_[exam_name]  # 高一上期中
            column = 0
            writer.write_excel_sheet(sheet, row, row, column, column, int(student.class_))
            column += 1
            writer.write_excel_sheet(sheet, row, row, column, column, student.name_)
            column += 1
            for key in student.score_attr_map_save_.keys():

                if student.score_attr_map_save_[key].is_main_ == True:
                    if exam_attr.score_attr_map_[key].get_score() >= 0:
                        writer.write_excel_sheet(sheet, row, row, column, column, exam_attr.score_attr_map_[key].get_score());
                        if exam_attr.score_attr_map_[key].rank_ > 0:
                            writer.write_excel_sheet(sheet, row, row, column + 1, column + 1, exam_attr.score_attr_map_[key].rank_);
                        else:
                            writer.write_excel_sheet(sheet, row, row, column + 1, column + 1, "  ");
                    else:
                        writer.write_excel_sheet(sheet, row, row, column, column, "  ");
                        writer.write_excel_sheet(sheet, row, row, column + 1, column + 1, "  ");
                    column += 2
                else:

                    if exam_attr.score_attr_map_[key].get_score() >= 0:
                        writer.write_excel_sheet(sheet, row, row, column, column, exam_attr.score_attr_map_[key].get_score());
                        if exam_attr.score_attr_map_[key].rank_ > 0:
                            writer.write_excel_sheet(sheet, row, row, column + 1, column + 1, exam_attr.score_attr_map_[key].rank_);
                        else:
                            writer.write_excel_sheet(sheet, row, row, column + 1, column + 1,"  ");
                        if exam_attr.score_attr_map_[key].assign_value_ > 0:
                            writer.write_excel_sheet_bold(sheet, row, row, column + 2, column + 2, exam_attr.score_attr_map_[key].assign_value_);
                        else:
                            writer.write_excel_sheet_bold(sheet, row, row, column + 2, column + 2,"  ");
                    else:
                        writer.write_excel_sheet(sheet, row, row, column, column, "  ");
                        writer.write_excel_sheet(sheet, row, row, column + 1, column + 1, "  ");
                        writer.write_excel_sheet(sheet, row, row, column + 2, column + 2, "  ");
                    column += 3
            # 总分
            writer.write_excel_sheet(sheet, row, row, column, column, exam_attr.score_total_);
            column += 1
            writer.write_excel_sheet(sheet, row, row, column, column, int(exam_attr.score_total_class_rank_));
            column += 1

            # 语数英
            writer.write_excel_sheet(sheet, row, row, column, column, exam_attr.score_total_main_);
            column += 1
            writer.write_excel_sheet(sheet, row, row, column, column, int(exam_attr.score_total_main_rank_));
            column += 1

            # 7选3
            writer.write_excel_sheet(sheet, row, row, column, column, exam_attr.score_total_not_main_);
            column += 1
            writer.write_excel_sheet(sheet, row, row, column, column, int(exam_attr.score_total_not_main_rank_));
            column += 1

            # 考试名
            writer.write_excel_sheet(sheet, row, row, column, column, exam_name);
            column += 1

            row += 1

        # 写空行
        writer.write_excel_sheet(sheet, row, row, 0, column_len - 1, " ")
        row += 1

        # for
        #     sheet.write(row, 0, student[1].name_);
        #     sheet.write(row, 1, student[1].class_);
        #     row += 1


if __name__ == '__main__':
    excel_files = []
    json_excel_files = json_reader.read_file("./format_conf.json")
    excel_files = json_excel_files["excel_files"]
    print(excel_files)

    student_map = {}
    read_student_exam_infos(student_map, excel_files)

    title = json_excel_files["format"]["title"]
    write_student_exam_infos_to_excel(student_map, title)

# print("xxxxxxx ", len(student_map))
# for student in student_map.values():
#    print(student.name_, " ", len(student.exam_attr_map_))
