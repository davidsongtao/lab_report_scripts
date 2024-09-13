"""
Description: 
    
-*- Encoding: UTF-8 -*-
@File     ：钢筋焊接.py
@Author   ：King Songtao
@Time     ：2024/9/10 下午5:46
@Contact  ：king.songtao@gmail.com
"""
from utils import *
import random
import csv


def load_data(source_data):
    """加载钢筋焊接委托台账，获取数据"""

    # 构建数据集
    data_loader = {
        "委托日期": [],
        "委托编号": [],
        "报告编号": [],
        "样品名称": [],
        "检测项目": [],
        "检验类别": [],
        "委托单位": [],
        "工程名称": [],
        "工程部位": [],
        "规格（mm）": [],
        "钢筋牌号": [],
        "焊工姓名": [],
        "焊工证号": [],
        "代表批量（个）": [],
        "检验形式": [],
        "见证员": [],
        "见证员证号": [],
        "取样员": [],
        "取样员证号": [],
        "施工单位": [],
        "见证单位": [],
        "报告领取": [],
        "备注": [],
    }

    # 读取源文件，获取数据，将数据加载进数据集
    try:
        with open(source_data, "r", encoding="GBK") as f:
            reader = csv.DictReader(f)
            for row in reader:
                for key, value in row.items():
                    if key == "见证员":
                        value += "；"
                    if key == "取样员":
                        value += "；"
                    if key == "代表批量（个）" and value == "":
                        value = "/"
                    if key == "代表批量（个）":
                        value += "个"
                    data_loader[key].append(value)
        logger.success(f"数据加载成功！")
    except Exception as e:
        logger.error(f"数据加载错误！错误信息：{e}")

    # 构建需要在报告中更新的数据(可以直接获取的)
    assignment_id = data_loader["委托编号"]
    report_id = data_loader["报告编号"]
    assignment_party = data_loader["委托单位"]
    project_name = data_loader["工程名称"]
    witnessing_agency = data_loader["见证单位"]
    assignment_type = data_loader["检验类别"]
    type = data_loader["样品名称"]
    sampling_date = data_loader["委托日期"]
    steel_grade = data_loader["钢筋牌号"]
    steel_model = data_loader["规格（mm）"]
    operator_name = data_loader["焊工姓名"]
    operator_id = data_loader["焊工证号"]
    test_type = data_loader["检验形式"]
    contractor = data_loader["施工单位"]
    project_section = data_loader["工程部位"]
    batch_size = data_loader["代表批量（个）"]
    witness = data_loader["见证员"]
    witness_id = data_loader["见证员证号"]
    sampler = data_loader["取样员"]
    sampler_id = data_loader["取样员证号"]
    test_type_list = data_loader["检测项目"]

    # 构建需要在报告中更新的数据(需要通过计算获得的)
    report_date = convert_sampling_date_to_report_date(sampling_date, 1)
    break_power_1 = []
    break_power_2 = []
    break_power_3 = []
    position_1 = []
    position_2 = []
    position_3 = []
    break_power_selections = ["630", "635", "640", "645", "650", "655"]
    position_selections = ["23", "24", "25", "26", "27", "28", "29", "30", "31", "32", "33", "34", "35", "36", "37", "38"]

    for model in steel_model:
        random_number_1 = random.choice(break_power_selections)
        random_number_2 = random.choice(break_power_selections)
        random_number_3 = random.choice(break_power_selections)
        break_power_1.append(random_number_1)
        break_power_2.append(random_number_2)
        break_power_3.append(random_number_3)

    # 构建断口位置数据
    for model in steel_model:
        random_number_4 = random.choice(position_selections)
        random_number_5 = random.choice(position_selections)
        random_number_6 = random.choice(position_selections)
        position_1.append(random_number_4)
        position_2.append(random_number_5)
        position_3.append(random_number_6)

    return test_type_list, assignment_id, report_id, assignment_party, project_name, witnessing_agency, assignment_type, type, sampling_date, report_date, batch_size, witness, witness_id, sampler, sampler_id, break_power_1, break_power_2, break_power_3, position_1, position_2, position_3, test_type, contractor, project_section, steel_grade, steel_model, operator_name, operator_id


def update_report(assignment_id, report_id, assignment_party, project_name, witnessing_agency, assignment_type, type, sampling_date, report_date, batch_size, witness, witness_id, sampler, sampler_id, break_power_1, break_power_2, break_power_3, position_1,
                  position_2, position_3, test_type, contractor, project_section, steel_grade, steel_model, operator_name, operator_id, reports_path):
    """对钢筋焊接检测报告内容，根据占位符进行更新"""
    # 获取文件夹中的所有 Word 文件
    files = [f for f in os.listdir(reports_path) if f.endswith('.docx')]
    # 检查给定的报告编号列表长度是否与文件夹中的文件个数相等。
    if len(files) != len(assignment_id) or len(files) != len(report_id) or len(files) != len(assignment_party) or len(files) != len(project_name) or len(files) != len(witnessing_agency) or len(files) != len(type) or len(files) != len(sampling_date) or len(
            files) != len(report_date) or len(files) != len(batch_size) or len(files) != len(witness) or len(files) != len(witness_id) or len(files) != len(sampler) or len(files) != len(sampler_id) or len(
        files) != len(break_power_1) or len(files) != len(break_power_2) or len(files) != len(break_power_3) or len(files) != len(position_1) or len(files) != len(position_2) or len(files) != len(position_3) or len(files) != len(test_type) or len(
        files) != len(contractor) or len(files) != len(project_section) or len(files) != len(steel_grade) or len(files) != len(steel_model) or len(files) != len(operator_name) or len(files) != len(operator_id):
        logger.error("文件数量与列表长度不匹配！")
        return
    # 若文件数量与列表长度匹配，则更新所有文档的报告编号
    for i, file_name in enumerate(files):
        try:
            file_path = os.path.join(reports_path, file_name)
            update_content(new_text=assignment_id[i], old_text="$assignment_id$", reports_path=file_path)
            update_content(new_text=report_id[i], old_text="$report_id$", reports_path=file_path)
            update_content(new_text=assignment_party[i], old_text="$assignment_party$", reports_path=file_path)
            update_content(new_text=project_name[i], old_text="$project_name$", reports_path=file_path)
            update_content(new_text=witnessing_agency[i], old_text="$witnessing_agency$", reports_path=file_path)
            update_content(new_text=assignment_type[i], old_text="$assignment_type$", reports_path=file_path)
            update_content(new_text=type[i], old_text="$type$", reports_path=file_path)
            update_content(new_text=sampling_date[i], old_text="$sampling_date$", reports_path=file_path)
            update_content(new_text=report_date[i], old_text="$report_date$", reports_path=file_path)
            update_content(new_text=batch_size[i], old_text="$batch_size$", reports_path=file_path)
            update_content(new_text=witness[i], old_text="$witness$", reports_path=file_path)
            update_content(new_text=witness_id[i], old_text="$witness_id$", reports_path=file_path)
            update_content(new_text=sampler[i], old_text="$sampler$", reports_path=file_path)
            update_content(new_text=sampler_id[i], old_text="$sampler_id$", reports_path=file_path)
            update_content(new_text=steel_grade[i], old_text="$steel_grade$", reports_path=file_path)
            update_content(new_text=steel_model[i], old_text="$steel_model$", reports_path=file_path)
            update_content(new_text=contractor[i], old_text="$contractor$", reports_path=file_path)
            update_content(new_text=project_section[i], old_text="$project_section$", reports_path=file_path)
            update_content(new_text=operator_name[i], old_text="$operator_name$", reports_path=file_path)
            update_content(new_text=operator_id[i], old_text="$operator_id$", reports_path=file_path)
            update_content(new_text=test_type[i], old_text="$test_type$", reports_path=file_path)
            update_content(new_text=break_power_1[i], old_text="$break_power_1$", center=True, reports_path=file_path)
            update_content(new_text=break_power_2[i], old_text="$break_power_2$", center=True, reports_path=file_path)
            update_content(new_text=break_power_3[i], old_text="$break_power_3$", center=True, reports_path=file_path)
            update_content(new_text=position_1[i], old_text="$position_1$", center=True, reports_path=file_path)
            update_content(new_text=position_2[i], old_text="$position_2$", center=True, reports_path=file_path)
            update_content(new_text=position_3[i], old_text="$position_3$", center=True, reports_path=file_path)
            logger.info(f"文件 {file_path} 中的所有数据已成功修改!")
        except Exception as e:
            logger.error(f"修改报告内容时发生错误！错误信息：{e}")

    logger.success(f"所有报告已完成！输出目录 >>> {reports_path}")


if __name__ == '__main__':
    test_type_list, assignment_id, report_id, assignment_party, project_name, witnessing_agency, assignment_type, type, sampling_date, report_date, batch_size, witness, witness_id, sampler, sampler_id, break_power_1, break_power_2, break_power_3, position_1, position_2, position_3, test_type, contractor, project_section, steel_grade, steel_model, operator_name, operator_id = load_data(
        param.rebar_welding_source_data)
    output_dir = create_reports(report_id, project_name, test_type_list, param.rebar_welding_template)
    update_report(assignment_id, report_id, assignment_party, project_name, witnessing_agency, assignment_type, type, sampling_date, report_date, batch_size, witness, witness_id, sampler, sampler_id, break_power_1, break_power_2, break_power_3, position_1,
                  position_2, position_3, test_type, contractor, project_section, steel_grade, steel_model, operator_name, operator_id, output_dir)
