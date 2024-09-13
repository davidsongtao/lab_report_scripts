"""
Description: 
    
-*- Encoding: UTF-8 -*-
@File     ：钢筋原材.py
@Author   ：King Songtao
@Time     ：2024/9/11 下午11:51
@Contact  ：king.songtao@gmail.com
"""
from utils import *
import random
import csv


def load_data(source_data):
    """
    加载钢筋原材委托台账，获取数据
    """
    # 构建数据集
    data_loader = {
        "委托日期": [],
        "委托编号": [],
        "报告编号": [],
        "钢筋种类": [],
        "检测项目": [],
        "检验类别": [],
        "委托单位": [],
        "工程名称": [],
        "工程部位": [],
        "规格（mm）": [],
        "钢筋牌号": [],
        "炉批号": [],
        "抗震等级": [],
        "代表批量（T）": [],
        "生产厂家": [],
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
                    if key == "代表批量（T）" and value == "":
                        value = "/"
                    if key == "代表批量（T）":
                        value += "T"
                    data_loader[key].append(value)
        logger.success(f"数据加载成功！")
    except Exception as e:
        logger.error(f"数据加载错误！错误信息：{e}")

    # 构建需要在报告中更新的数据(可以直接获取的)
    assignment_id = data_loader["委托编号"]
    report_id = data_loader["报告编号"]
    project_name = data_loader["工程名称"]
    witnessing_agency = data_loader["见证单位"]
    assignment_type = data_loader["检验类别"]
    assignment_party = data_loader["委托单位"]
    contractor = data_loader["施工单位"]
    batch_id = data_loader["炉批号"]
    sampling_date = data_loader["委托日期"]
    report_date = convert_sampling_date_to_report_date(sampling_date, 1)
    steel_type = data_loader["钢筋种类"]
    steel_level = data_loader["抗震等级"]
    steel_grade = data_loader["钢筋牌号"]
    steel_model = data_loader["规格（mm）"]
    manufacturer = data_loader["生产厂家"]
    project_section = data_loader["工程部位"]
    batch_size = data_loader["代表批量（T）"]
    witness = data_loader["见证员"]
    witness_id = data_loader["见证员证号"]
    sampler = data_loader["取样员"]
    sampler_id = data_loader["取样员证号"]
    test_type_list = data_loader["检测项目"]

    # 构建需要在报告中更新的数据(需要通过计算获得的)
    rel_1 = []  # 屈服强度_1
    rel_2 = []  # 屈服强度_2
    rm_1 = []  # 抗拉强度_1
    rm_2 = []  # 抗拉强度_2
    a_per_1 = []  # 断后伸长率_1
    a_per_2 = []  # 断后伸长率_2
    agt_1 = []  # 最大总伸长率_1
    agt_2 = []  # 最大总伸长率_2
    qbb_1 = []  # 屈标比_1（屈服强度实测值与强度标准值的比，值不应大于1.3）
    qbb_2 = []  # 屈标比_2
    qqb_1 = []  # 强屈比_1（抗拉强度实测值/屈服强度实测值，结果不能小于1.25）
    qqb_2 = []  # 强屈比_2
    weight = []  # 重量偏差（取值4%以内）

    # 构建屈服强度数据
    for model in steel_model:
        random_num_1 = random.randint(429, 442)
        random_num_2 = random.randint(429, 442)
        if random_num_1 == random_num_2:
            random_num_2 += 1
        rel_1.append(str(random_num_1))
        rel_2.append(str(random_num_2))

    # 构建抗拉强度数据
    for model in steel_model:
        random_num_1 = random.randint(585, 611)
        random_num_2 = random.randint(585, 611)
        if random_num_1 == random_num_2:
            random_num_2 += 1
        rm_1.append(str(random_num_1))
        rm_2.append(str(random_num_2))

    # 构建断后伸长率数据
    for model in steel_model:
        random_num_1 = round(random.randint(50, 55) * 5 / 10)
        random_num_2 = round(random.randint(50, 55) * 5 / 10)
        if random_num_1 == random_num_2:
            random_num_2 += 1
        a_per_1.append(str(random_num_1))
        a_per_2.append(str(random_num_2))

    # 构建最大总伸长率数据
    for model in steel_model:
        random_num_1 = round(random.uniform(12.1, 13.8), 1)
        random_num_2 = round(random.uniform(12.1, 13.8), 1)
        if random_num_1 == random_num_2:
            random_num_2 += 0.1
        agt_1.append(str(random_num_1))
        agt_2.append(str(random_num_2))

    # 构建屈标比数据
    for model in steel_model:
        random_num_1 = round(random.uniform(1.06, 1.14), 2)
        random_num_2 = round(random.uniform(1.06, 1.14), 2)
        if random_num_1 == random_num_2:
            random_num_2 += 0.01
        qbb_1.append(str(random_num_1))
        qbb_2.append(str(random_num_2))

    # 构建强屈比数据
    for i, model in enumerate(steel_model):
        qqb_num_1 = round(float(rm_1[i]) / float(rel_1[i]), 2)
        qqb_num_2 = round(float(rm_2[i]) / float(rel_2[i]), 2)
        qqb_1.append(str(qqb_num_1))
        qqb_2.append(str(qqb_num_2))

    # 构建重量偏差数据
    for model in steel_model:
        random_num = str(round(random.uniform(-3.9, 3.9), 1))
        weight.append(random_num)

    return assignment_id, report_id, project_name, witnessing_agency, assignment_type, assignment_party, contractor, batch_id, sampling_date, report_date, steel_type, steel_level, steel_grade, steel_model, manufacturer, project_section, batch_size, witness, witness_id, sampler, sampler_id, test_type_list, rel_1, rel_2, rm_1, rm_2, a_per_1, a_per_2, agt_1, agt_2, qbb_1, qbb_2, qqb_1, qqb_2, weight


def update_report(assignment_id, report_id, project_name, witnessing_agency, assignment_type, assignment_party, contractor, batch_id, sampling_date, report_date, steel_type, steel_level, steel_grade, steel_model, manufacturer, project_section, batch_size,
                  witness, witness_id, sampler, sampler_id, test_type_list, rel_1, rel_2, rm_1, rm_2, a_per_1, a_per_2, agt_1, agt_2, qbb_1, qbb_2, qqb_1, qqb_2, weight, reports_path):
    """对钢筋原材检测报告内容，根据占位符进行更新"""
    # 获取文件夹中的所有 Word 文件
    files = [f for f in os.listdir(reports_path) if f.endswith('.docx')]
    # 检查给定的报告编号列表长度是否与文件夹中的文件个数相等。
    if len(files) != len(assignment_id) or len(files) != len(report_id) or len(files) != len(project_name) or len(files) != len(witnessing_agency) or len(files) != len(assignment_type) or len(files) != len(assignment_party) or len(files) != len(
            contractor) or len(files) != len(batch_id) or len(files) != len(sampling_date) or len(files) != len(report_date) or len(files) != len(steel_type) or len(files) != len(steel_level) or len(files) != len(steel_grade) or len(
        files) != len(steel_model) or len(files) != len(manufacturer) or len(files) != len(project_section) or len(files) != len(batch_size) or len(files) != len(witness) or len(files) != len(witness_id) or len(files) != len(sampler) or len(
        files) != len(sampler_id) or len(files) != len(test_type_list) or len(files) != len(rel_1) or len(files) != len(rel_2) or len(files) != len(rm_1) or len(files) != len(rm_2) or len(files) != len(a_per_1) or len(files) != len(a_per_2) or len(
        files) != len(agt_2) or len(files) != len(qbb_1) or len(files) != len(qbb_2) or len(files) != len(weight):
        logger.error("文件数量与列表长度不匹配！")
        return

    # 若文件数量与列表长度匹配，则更新所有文档的报告编号
    for i, file_name in enumerate(files):
        try:
            file_path = os.path.join(reports_path, file_name)
            update_content(new_text=assignment_id[i], old_text="$assignment_id$", reports_path=file_path)
            update_content(new_text=assignment_party[i], old_text="$assignment_party$", reports_path=file_path)
            update_content(new_text=project_name[i], old_text="$project_name$", reports_path=file_path)
            update_content(new_text=witnessing_agency[i], old_text="$witnessing_agency$", reports_path=file_path)
            update_content(new_text=assignment_type[i], old_text="$assignment_type$", reports_path=file_path)
            update_content(new_text=contractor[i], old_text="$contractor$", reports_path=file_path)
            update_content(new_text=batch_id[i], old_text="$batch_id$", reports_path=file_path)
            update_content(new_text=sampling_date[i], old_text="$sampling_date$", reports_path=file_path)
            update_content(new_text=report_date[i], old_text="$report_date$", reports_path=file_path)
            update_content(new_text=steel_type[i], old_text="$steel_type$", reports_path=file_path)
            update_content(new_text=steel_level[i], old_text="$steel_level$", reports_path=file_path)
            update_content(new_text=report_id[i], old_text="$report_id$", font_size=8, reports_path=file_path)
            update_content(new_text=steel_grade[i], old_text="$steel_grade$", font_size=8, reports_path=file_path)
            update_content(new_text=steel_model[i], old_text="$steel_model$", font_size=8, reports_path=file_path)
            update_content(new_text=manufacturer[i], old_text="$manufacturer$", font_size=8, reports_path=file_path)
            update_content(new_text=project_section[i], old_text="$project_section$", font_size=8, reports_path=file_path)
            update_content(new_text=batch_size[i], old_text="$batch_size$", center=True, reports_path=file_path)
            update_content(new_text=weight[i], old_text="$weight$", center=True, reports_path=file_path)
            update_content(new_text=rel_1[i], old_text="$rel_1$", center=True, reports_path=file_path)
            update_content(new_text=rel_2[i], old_text="$rel_2$", center=True, reports_path=file_path)
            update_content(new_text=rm_1[i], old_text="$rm_1$", center=True, reports_path=file_path)
            update_content(new_text=rm_2[i], old_text="$rm_2$", center=True, reports_path=file_path)
            update_content(new_text=a_per_1[i], old_text="$a_per_1$", center=True, reports_path=file_path)
            update_content(new_text=a_per_2[i], old_text="$a_per_2$", center=True, reports_path=file_path)
            update_content(new_text=agt_1[i], old_text="$agt_1$", center=True, reports_path=file_path)
            update_content(new_text=agt_2[i], old_text="$agt_2$", center=True, reports_path=file_path)
            update_content(new_text=qbb_1[i], old_text="$qbb_1$", font_size=8, reports_path=file_path)
            update_content(new_text=qbb_2[i], old_text="$qbb_2$", font_size=8, reports_path=file_path)
            update_content(new_text=qqb_1[i], old_text="$qqb_1$", font_size=8, reports_path=file_path)
            update_content(new_text=qqb_2[i], old_text="$qqb_2$", font_size=8, reports_path=file_path)
            update_content(new_text=witness[i], old_text="$witness$", reports_path=file_path)
            update_content(new_text=witness_id[i], old_text="$witness_id$", reports_path=file_path)
            update_content(new_text=sampler[i], old_text="$sampler$", reports_path=file_path)
            update_content(new_text=sampler_id[i], old_text="$sampler_id$", reports_path=file_path)
            logger.info(f"文件 {file_path} 中的所有数据已成功修改!")
        except ValueError as e:
            logger.error(f"修改报告内容时发生错误！错误信息：{e}")

    logger.success(f"所有报告已完成！输出目录 >>> {output_dir}")


if __name__ == '__main__':
    assignment_id, report_id, project_name, witnessing_agency, assignment_type, assignment_party, contractor, batch_id, sampling_date, report_date, steel_type, steel_level, steel_grade, steel_model, manufacturer, project_section, batch_size, witness, witness_id, sampler, sampler_id, test_type_list, rel_1, rel_2, rm_1, rm_2, a_per_1, a_per_2, agt_1, agt_2, qbb_1, qbb_2, qqb_1, qqb_2, weight = load_data(
        param.steel_source_data)
    output_dir = create_reports(report_id, project_name, test_type_list, param.steel_template)
    update_report(assignment_id, report_id, project_name, witnessing_agency, assignment_type, assignment_party, contractor, batch_id, sampling_date, report_date, steel_type, steel_level, steel_grade, steel_model, manufacturer, project_section, batch_size,
                  witness, witness_id, sampler, sampler_id, test_type_list, rel_1, rel_2, rm_1, rm_2, a_per_1, a_per_2, agt_1, agt_2, qbb_1, qbb_2, qqb_1, qqb_2, weight, output_dir)
