"""
Description:

-*- Encoding: UTF-8 -*-
@File     ：砼试块抗压强度.py
@Author   ：King Songtao
@Time     ：2024/9/9 下午11:15
@Contact  ：king.songtao@gmail.com
"""
import csv
import random
from utils import *


def update_report(report_ids, assignment_ids, assignment_party, project_name, witnessing_agency, strength_grade, maintenance_condition, contractor, project_section, molding_date, sampling_date, report_date, batch_size, curing_age,
                  witness, witness_id, sampler, sampler_id, break_power_1, strength_1, break_power_2, strength_2, break_power_3, strength_3, mean_strength, strength_percentage, assignment_type, reports_path):
    """对检测报告内容，根据占位符进行更新"""
    # 获取文件夹中的所有 Word 文件
    files = [f for f in os.listdir(reports_path) if f.endswith('.docx')]
    # 检查给定的报告编号列表长度是否于文件夹中的文件个数相等。
    if len(files) != len(report_ids) or len(files) != len(assignment_ids) or len(files) != len(assignment_party) or len(files) != len(project_name) or len(files) != len(witnessing_agency) or len(files) != len(strength_grade) or len(files) != len(
            maintenance_condition) or len(files) != len(contractor) or len(files) != len(project_section) or len(files) != len(molding_date) or len(files) != len(sampling_date) or len(files) != len(report_date) or len(files) != len(batch_size) or len(
        files) != len(curing_age) or len(files) != len(witness) or len(files) != len(witness_id) or len(files) != len(sampler) or len(files) != len(sampler_id) or len(files) != len(break_power_1) or len(files) != len(strength_1) or len(files) != len(
        break_power_2) or len(files) != len(strength_2) or len(files) != len(break_power_3) or len(files) != len(strength_3) or len(files) != len(mean_strength) or len(files) != len(strength_percentage) or len(files) != len(assignment_type):
        logger.error("报告编号有重复！")
        return
    # 若文件数量与列表长度匹配，则更新所有文档的报告编号
    for i, file_name in enumerate(files):
        try:
            file_path = os.path.join(reports_path, file_name)
            update_content(new_text=report_ids[i], old_text="$report_id$", reports_path=file_path)
            update_content(new_text=assignment_ids[i], old_text="$assignment_id$", reports_path=file_path)
            update_content(new_text=assignment_party[i], old_text="$assignment_party$", reports_path=file_path)
            update_content(new_text=project_name[i], old_text="$project_name$", reports_path=file_path)
            update_content(new_text=witnessing_agency[i], old_text="$witnessing_agency$", reports_path=file_path)
            update_content(new_text=strength_grade[i], old_text="$strength_grade$", reports_path=file_path)
            update_content(new_text=maintenance_condition[i], old_text="$maintenance_condition$", reports_path=file_path)
            update_content(new_text=contractor[i], old_text="$contractor$", reports_path=file_path)
            update_content(new_text=project_section[i], old_text="$project_section$", reports_path=file_path)
            update_content(new_text=molding_date[i], old_text="$molding_date$", reports_path=file_path)
            update_content(new_text=sampling_date[i], old_text="$sampling_date$", reports_path=file_path)
            update_content(new_text=report_date[i], old_text="$report_date$", reports_path=file_path)
            update_content(new_text=batch_size[i], old_text="$batch_size$", reports_path=file_path)
            update_content(new_text=witness[i], old_text="$witness$", reports_path=file_path)
            update_content(new_text=witness_id[i], old_text="$witness_id$", reports_path=file_path)
            update_content(new_text=sampler[i], old_text="$sampler$", reports_path=file_path)
            update_content(new_text=sampler_id[i], old_text="$sampler_id$", reports_path=file_path)
            update_content(new_text=assignment_type[i], old_text="$assignment_type$", reports_path=file_path)
            update_content(new_text=curing_age[i], old_text="$curing_age$", center=True, reports_path=file_path)
            update_content(new_text=break_power_1[i], old_text="$break_power_1$", center=True, reports_path=file_path)
            update_content(new_text=strength_1[i], old_text="$strength_1$", center=True, reports_path=file_path)
            update_content(new_text=break_power_2[i], old_text="$break_power_2$", center=True, reports_path=file_path)
            update_content(new_text=strength_2[i], old_text="$strength_2$", center=True, reports_path=file_path)
            update_content(new_text=break_power_3[i], old_text="$break_power_3$", center=True, reports_path=file_path)
            update_content(new_text=strength_3[i], old_text="$strength_3$", center=True, reports_path=file_path)
            update_content(new_text=mean_strength[i], old_text="$mean_strength$", center=True, reports_path=file_path)
            update_content(new_text=strength_percentage[i], old_text="$strength_percentage$", center=True, reports_path=file_path)
            logger.info(f"检测报告 >>>{file_name}<<< 中的所有数据已成功修改!")
        except Exception as e:
            logger.error(f"修改报告内容时发生错误！错误信息：{e}")

    logger.success(f"所有报告已完成！输出目录 >>> {output_dir}")


def load_data(source_data):
    """从CSV文件中加载委托信息，构建数据列表群"""
    # 读取csv委托台账，从台账中抽取对应数据
    data = {
        "委托日期": [], "委托编号": [], "报告编号": [], "样品名称": [], "检测项目": [], "委托单位": [],
        "工程名称": [], "工程部位": [], "强度等级": [], "成型日期": [], "龄期（天）": [], "代表批量（m3）": [],
        "养护条件": [], "见证员": [], "见证员证号": [], "取样员": [], "取样员证号": [], "施工单位": [],
        "见证单位": [], "见证日期": [], "报告领取": [], "备注": [], "检验类别": []
    }
    try:
        with open(source_data, "r", encoding="GBK") as f:
            reader = csv.DictReader(f)
            for row in reader:
                for key, value in row.items():
                    if key in ["见证员", "取样员"]:
                        value += "；"
                    elif key == "代表批量（m3）":
                        value = "/" if value == "" else f"{value}m³"
                    data[key].append(value)
        logger.success(f"数据加载成功！")
    except Exception as e:
        logger.error(f"数据加载错误！错误信息：{e}")

    # 定义每钟强度的混凝土对应的破坏和在范围
    ranges = {
        "C35": (401, 459), "C30": (381, 429), "C25": (301, 359), "C15": (201, 259),
        "C40": (481, 529), "C50": (581, 629), "C45": (501, 559), "C20": (281, 329),
    }

    # 获取所有报告对应的数据
    report_ids = data["报告编号"]
    assignment_ids = data["委托编号"]
    assignment_party = data["委托单位"]
    project_name = data["工程名称"]
    witnessing_agency = data["见证单位"]
    strength_grade = data["强度等级"]
    maintenance_condition = data["养护条件"]
    contractor = data["施工单位"]
    project_section = data["工程部位"]
    molding_date = data["成型日期"]
    sampling_date = data["委托日期"]
    report_date = convert_sampling_date_to_report_date(sampling_date, 1)
    batch_size = data["代表批量（m3）"]
    curing_age = data["龄期（天）"]
    witness = data["见证员"]
    witness_id = data["见证员证号"]
    sampler = data["取样员"]
    sampler_id = data["取样员证号"]
    assignment_type = data["检验类别"]
    test_type = data["检测项目"]
    test_type_name = data["样品名称"]

    # 根据强度等级确定对应的破坏荷载实验数据
    break_power_1 = []
    break_power_2 = []
    break_power_3 = []
    for grade in strength_grade:
        if grade in ranges:
            min_val, max_val = ranges[grade]
            random_number = round(random.uniform(min_val, max_val), 2)
            random_number_2 = round(random.uniform(min_val, max_val), 2)
            random_number_3 = round(random.uniform(min_val, max_val), 2)
            break_power_1.append(f"{random_number:.2f}")
            break_power_2.append(f"{random_number_2:.2f}")
            break_power_3.append(f"{random_number_3:.2f}")

        else:
            # 处理不在预定义范围内的情况，这里假设直接跳过
            break_power_1.append(None)
            break_power_2.append(None)
            break_power_3.append(None)
    # 根据破坏和在计算强度
    strength_1 = [f"{round(float(num) * 0.095, 1):.1f}" for num in break_power_1]
    strength_2 = [f"{round(float(num) * 0.095, 1):.1f}" for num in break_power_2]
    strength_3 = [f"{round(float(num) * 0.095, 1):.1f}" for num in break_power_3]

    mean_strength = [
        f"{round((float(s1) + float(s2) + float(s3)) / 3, 1):.1f}"
        for s1, s2, s3 in zip(strength_1, strength_2, strength_3)
    ]

    # 根据强度平均值和强度等级计算强度百分比
    strength_percentage = []
    for mean, grade in zip(mean_strength, strength_grade):
        # 提取 strength_grade 中的数字部分
        numeric_value = int(grade[1:])
        # 转换 mean_strength 中的值为浮点数
        mean_value = float(mean)
        # 计算百分比
        percentage = (mean_value / numeric_value) * 100
        # 将结果格式化为字符串并保留两位小数
        formatted_percentage = f"{int(percentage)}"
        # 添加到新列表中
        strength_percentage.append(formatted_percentage)

    return test_type, report_ids, assignment_ids, assignment_party, project_name, witnessing_agency, strength_grade, maintenance_condition, contractor, project_section, molding_date, sampling_date, report_date, batch_size, curing_age, witness, witness_id, sampler, sampler_id, break_power_1, strength_1, break_power_2, strength_2, break_power_3, strength_3, mean_strength, strength_percentage, assignment_type


if __name__ == '__main__':
    # 第一步 -> 读取数据源csv文件
    test_type, report_ids, assignment_ids, assignment_party, project_name, witnessing_agency, strength_grade, maintenance_condition, contractor, project_section, molding_date, sampling_date, report_date, batch_size, curing_age, witness, witness_id, sampler, sampler_id, break_power_1, strength_1, break_power_2, strength_2, break_power_3, strength_3, mean_strength, strength_percentage, assignment_type = load_data(
        param.concrete_strength_source_data)
    # 第二步 -> 根据读取到的报告编号构建检测报告word文件
    output_dir = create_reports(report_ids, project_name, test_type, param.concrete_strength_template)
    # 第三步 -> 更新报告模板中的数据
    update_report(report_ids, assignment_ids, assignment_party, project_name, witnessing_agency, strength_grade, maintenance_condition, contractor, project_section, molding_date, sampling_date, report_date, batch_size, curing_age, witness, witness_id, sampler,
                  sampler_id, break_power_1, strength_1, break_power_2, strength_2, break_power_3, strength_3, mean_strength, strength_percentage, assignment_type, output_dir)
