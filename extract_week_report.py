import pandas as pd
import re
import os

index_class_profile, col_class_profile = 2, 3
index_new_input_count, col_new_input_count = 5, 2
index_attention_count, col_attention_count = 4, 2
index_lose_count, col_lose_count = 3, 2
index_student_count, col_student_count = 2, 2
index_teacher, col_teacher = 0, 0
index_date, col_date = 0, 2
index_class_name = 0


def extract_subject_name(formatted_name):
    """
    提取学科名字
    :param formatted_name:格式为 xxx学科xxx阶段xxx群 的字符串
    :return:
    """
    pattern = "(.*)学科"
    try:
        sub_name = re.match(pattern, formatted_name).group(1)
    except Exception:
        sub_name = ""
    return sub_name


def extract_date(formatted_date):
    """
    提取日期
    :param formatted_date:格式为 日期：2018.9.3--9.9
    :return:
    """
    pattern = "日期.(.*)"
    try:
        d = re.match(pattern, formatted_date).group(1)
    except Exception:
        d = ""
    return d


def extract_supervisor_teacher(formatted_name):
    """
    提取导师和班主任名字
    :param formatted_name:格式为 导师：xxx   班主任：xxx
    :return:
    """
    pattern = "导师.(.*).*班主任.(.*)"
    try:
        match = re.match(pattern, formatted_name)
        sv = match.group(1)
        tc = match.group(2)
    except Exception:
        sv = ""
        tc = ""
    return sv, tc


def extract_class_name(file_name):
    """
    从文件名提取班级名称
    :param file_name:
    :return:
    """
    pattern = "(.*)周报(.*)"
    try:
        match = re.match(pattern, file_name)
        class_name = match.group(1)
    except Exception:
        class_name = ""
    return class_name


def extract_class_group_data(template_data_df, file_name):
    """
    提取班级群数据
    :param template_data_df: 模板数据 dataframe
    :param file_name: 班级群数据文件名
    :return:
    """
    # data_df = pd.read_excel(os.path.join(dir, group_dir, file_name), sheet_name=["群周报填写规范"])["群周报填写规范"]
    data_df = pd.read_excel(os.path.join(dir, group_dir, file_name))
    class_name = data_df.columns[index_class_name]  # 班级名称

    # 检查文件名和班级名是否匹配
    # class_name_from_file_name = extract_class_name(file_name)
    # if class_name != class_name_from_file_name:
    #     print(file_name, "文件文件名和班级名称不一致")
    #     return template_data_df

    subject_name = extract_subject_name(data_df.columns[index_class_name])  # 学科名称
    date = extract_date(data_df.iloc[index_date, col_date])  # 日期
    supervisor, teacher = extract_supervisor_teacher(data_df.iloc[index_teacher, col_teacher])  # 导师 班主任
    student_count = data_df.iloc[index_student_count, col_student_count]  # 学员总数
    lose_count = data_df.iloc[index_lose_count, col_lose_count]  # 失联人数
    attention_count = data_df.iloc[index_attention_count, col_attention_count]  # 重点关注学员
    new_input_count = data_df.iloc[index_new_input_count, col_new_input_count]  # 新入群学员
    class_profile = data_df.iloc[index_class_profile, col_class_profile]  # 群概况

    new_student_count = data_df.iloc[6, 2]  # 本周新进人数
    arrangement_count = data_df.iloc[7, 2]  # 布置任务人数
    complete_count = data_df.iloc[8, 2]  # 完成任务人数
    incomplete_count = data_df.iloc[9, 2]  # 历史未完成任务人数
    task_profile = data_df.iloc[6, 3]  # 总结

    portfolio_update_count = data_df.iloc[10, 2]  # 有效更新数量
    portfolio_update_profile = data_df.iloc[10, 3]  # 总结

    upgraded_count = data_df.iloc[11, 2]  # 成功晋级人数
    upgrading_count = data_df.iloc[12, 2]  # 正在完成晋级任务人数
    upgraded_profile = data_df.iloc[11, 3]  # 总结

    note_count = data_df.iloc[13, 2]  # 笔记数量（非系统）
    excellent_note_count = data_df.iloc[14, 2]  # 优秀笔记（非系统）
    note_profile = data_df.iloc[13, 3]  # 总结

    live_video_count = data_df.iloc[15, 2]  # 直播次数
    live_video_profile = data_df.iloc[15, 3]  # 总结

    QA = data_df.iloc[16, 1]  # 答疑内容
    QA_profile = data_df.iloc[16, 3]  # 总结

    others = data_df.iloc[17, 1]  # 其他内容
    others_profile = data_df.iloc[17, 3]  # 总结

    next_week_plan = data_df.iloc[18, 1]  # 下周计划
    class_data = [subject_name,
                  date,
                  supervisor,
                  teacher,
                  student_count,
                  lose_count,
                  attention_count,
                  new_input_count,
                  class_profile,
                  new_student_count,
                  arrangement_count,
                  complete_count,
                  incomplete_count,
                  task_profile,
                  portfolio_update_count,
                  portfolio_update_profile,
                  upgraded_count,
                  upgrading_count,
                  upgraded_profile,
                  note_count,
                  excellent_note_count,
                  note_profile,
                  live_video_count,
                  live_video_profile,
                  QA,
                  QA_profile,
                  others,
                  others_profile,
                  next_week_plan]
    template_data_df[class_name] = class_data
    return template_data_df


def extract_name(formatted_name):
    """
    提取老師的名字
    :param formatted_name:格式为 姓名：XXX
    :return:
    """
    pattern = "姓名：(.*)"
    try:
        tn = re.match(pattern, formatted_name).group(1)
    except Exception:
        tn = ""
    return tn


def extract_subject_data(template_data_df, file_name):
    """
    提取学科群数据
    :param template_data_df: 模板数据 dataframe
    :param file_name: 班级群数据文件名
    :return:
    """
    data_df = pd.read_excel(os.path.join(dir, subject_dir, file_name))
    subject_name = extract_subject_name(data_df.columns[0])  # 学科名称
    name = extract_name(data_df.iloc[0, 0])  # 姓名

    subject_summary = data_df.iloc[1, 1]  # 学科总结
    now_problem = data_df.iloc[2, 1]  # 现存问题
    subject_plan = data_df.iloc[3, 1]  # 学科计划

    subject_data = [name, subject_summary, now_problem, subject_plan]
    template_data_df[subject_name] = subject_data
    return template_data_df

def prompt_choice():
    print("""
欢迎使用周报数据合并工具，本工具可以按模板合并群周报和学科周报
1. 合并群周报请输入1
2. 合并学科周报请输入2
3. 退出工具请输入3
    """)


if __name__ == "__main__":
    # dir = os.getcwd()
    dir = "/Users/bravian/Desktop/9月17日各位导师提交周报"
    group_dir = "群周报"
    subject_dir = "学科周报"
    group_template_file_name = "群周报汇总模板V0.1.xlsx"
    subject_template_file_name = "学科周报汇总模板V0.1.xlsx"

    print("*" * 40)
    print("{:^30}".format("周报数据合并工具"))
    print("*" * 40)
    prompt_choice()
    while True:
        choice = input("")

        if choice == "1":
            if not os.path.isdir(os.path.join(dir, group_dir)):
                print("当前目录下群周报文件夹不存在")
                continue
            if not os.path.exists(os.path.join(dir, group_template_file_name)):
                print("当前目录下群周报模板文件不存在")
                continue

            template_df = pd.read_excel(os.path.join(dir, group_template_file_name))

            listdir = os.listdir(os.path.join(dir, group_dir))
            listdir.sort()
            for group_file_name in listdir:
                try:
                    template_df = extract_class_group_data(template_df, group_file_name)
                except Exception:
                    print(group_file_name + " 文件不符合模板要求")
            template_df = template_df.transpose()
            template_df.to_excel("群周报汇总V0.1.xlsx")
            print("已经生成《群周报汇总V0.1.xlsx》")
        elif choice == "2":
            if not os.path.isdir(os.path.join(dir, subject_dir)):
                print("当前目录下学科周报文件夹不存在")
                continue

            if not os.path.exists(os.path.join(dir, subject_template_file_name)):
                print("当前目录下学科周报模板文件不存在")
                continue

            template_df = pd.read_excel(os.path.join(dir, subject_template_file_name))
            del template_df["数据"]
            for subject_file_name in os.listdir(os.path.join(dir, subject_dir)):
                try:
                    template_df = extract_subject_data(template_df, subject_file_name)
                except Exception as e:
                    print(subject_file_name + " 文件不符合模板要求")
            template_df = template_df.transpose()
            template_df.to_excel("学科周报汇总V0.1.xlsx")
            print("已经生成《学科周报汇总V0.1.xlsx》")
        elif choice == "3":
            break
        else:
            print("请输入正确的选项")
            prompt_choice()
