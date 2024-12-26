"""
Description: 
    
-*- Encoding: UTF-8 -*-
@File     ：灰剂量.py
@Author   ：King Songtao
@Time     ：2024/12/26 下午3:03
@Contact  ：king.songtao@gmail.com
"""

from utils import *
import csv
import random
from decimal import Decimal, ROUND_HALF_UP


def load_data(source_data):
    """
    加载灰剂量委托台账，获取数据。
    :param source_data:
    :return:
    """

    data_loader = {
        "委托日期": [],
        "委托编号": [],
        "报告编号": [],
        "土壤种类": [],
        "结合料种类": [],
        "检测项目": [],
        "检验类别": [],
        "委托单位": [],
        "工程名称": [],
        "工程部位": [],
        "检测点数": [],
        "见证员": [],
        "见证员证号": [],
        "取样员": [],
        "取样员证号": [],
        "施工单位": [],
        "见证单位": [],
        "报告领取": [],
        "备注": [],
        "实验数据": []
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
                    data_loader[key].append(value)
        logger.success(f"数据加载成功！")
        print(data_loader)
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
    sampling_date = data_loader["委托日期"]
    report_date = convert_sampling_date_to_report_date(sampling_date, 1)
    project_section = data_loader["工程部位"]
    witness = data_loader["见证员"]
    witness_id = data_loader["见证员证号"]
    sampler = data_loader["取样员"]
    sampler_id = data_loader["取样员证号"]
    test_type_list = data_loader["检测项目"]
    sample_type = data_loader["土壤种类"]
    total_points = data_loader["检测点数"]
    material_type = data_loader["结合料种类"]

    print(total_points)

    # 构建实验数据
    sampling_data_list = []

    for index, point in enumerate(total_points):

        sampling_data = {}

        # 根据point的值处理对应数量的数据组
        if point == "1":
            # 处理sample_id
            sampling_data['sample_id_1'] = "1"
            sampling_data['sample_id_2'] = "/"
            sampling_data['sample_id_3'] = "/"
            sampling_data['sample_id_4'] = "/"
            sampling_data['sample_id_5'] = "/"

            # 处理edta_ml
            sampling_data['edta_ml_1_1'] = Decimal(str(random.uniform(5.7, 6.2))).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_ml_1_2'] = Decimal(str(random.uniform(5.7, 6.2))).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_ml_2_1'] = "/"
            sampling_data['edta_ml_2_2'] = "/"
            sampling_data['edta_ml_3_1'] = "/"
            sampling_data['edta_ml_3_2'] = "/"
            sampling_data['edta_ml_4_1'] = "/"
            sampling_data['edta_ml_4_2'] = "/"
            sampling_data['edta_ml_5_1'] = "/"
            sampling_data['edta_ml_5_2'] = "/"

            # 处理edta_avg
            sampling_data['edta_avg_1'] = ((sampling_data['edta_ml_1_1'] + sampling_data['edta_ml_1_2']) / Decimal('2.0')).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_avg_2'] = '/'
            sampling_data['edta_avg_3'] = '/'
            sampling_data['edta_avg_4'] = '/'
            sampling_data['edta_avg_5'] = '/'

            # 处理material_per
            sampling_data['material_per_1'] = (Decimal('9.4302') - sampling_data['edta_avg_1'] * Decimal('0.7059')).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['material_per_2'] = '/'
            sampling_data['material_per_3'] = '/'
            sampling_data['material_per_4'] = '/'
            sampling_data['material_per_5'] = '/'

        elif point == "2":
            # 处理sample_id
            sampling_data['sample_id_1'] = "1"
            sampling_data['sample_id_2'] = "2"
            sampling_data['sample_id_3'] = "/"
            sampling_data['sample_id_4'] = "/"
            sampling_data['sample_id_5'] = "/"

            # 处理edta_ml
            sampling_data['edta_ml_1_1'] = Decimal(str(random.uniform(5.7, 6.2))).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_ml_1_2'] = Decimal(str(random.uniform(5.7, 6.2))).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_ml_2_1'] = Decimal(str(random.uniform(5.7, 6.2))).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_ml_2_2'] = Decimal(str(random.uniform(5.7, 6.2))).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_ml_3_1'] = "/"
            sampling_data['edta_ml_3_2'] = "/"
            sampling_data['edta_ml_4_1'] = "/"
            sampling_data['edta_ml_4_2'] = "/"
            sampling_data['edta_ml_5_1'] = "/"
            sampling_data['edta_ml_5_2'] = "/"

            # 处理edta_avg
            sampling_data['edta_avg_1'] = ((sampling_data['edta_ml_1_1'] + sampling_data['edta_ml_1_2']) / Decimal('2.0')).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_avg_2'] = ((sampling_data['edta_ml_2_1'] + sampling_data['edta_ml_2_2']) / Decimal('2.0')).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_avg_3'] = '/'
            sampling_data['edta_avg_4'] = '/'
            sampling_data['edta_avg_5'] = '/'

            # 处理material_per
            sampling_data['material_per_1'] = (Decimal('9.4302') - sampling_data['edta_avg_1'] * Decimal('0.7059')).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['material_per_2'] = (Decimal('9.4302') - sampling_data['edta_avg_2'] * Decimal('0.7059')).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['material_per_3'] = '/'
            sampling_data['material_per_4'] = '/'
            sampling_data['material_per_5'] = '/'



        elif point == "3":

            # 处理sample_id
            sampling_data['sample_id_1'] = "1"
            sampling_data['sample_id_2'] = "2"
            sampling_data['sample_id_3'] = "3"
            sampling_data['sample_id_4'] = "/"
            sampling_data['sample_id_5'] = "/"

            # 处理edta_ml
            sampling_data['edta_ml_1_1'] = Decimal(str(random.uniform(5.7, 6.2))).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_ml_1_2'] = Decimal(str(random.uniform(5.7, 6.2))).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_ml_2_1'] = Decimal(str(random.uniform(5.7, 6.2))).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_ml_2_2'] = Decimal(str(random.uniform(5.7, 6.2))).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_ml_3_1'] = Decimal(str(random.uniform(5.7, 6.2))).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_ml_3_2'] = Decimal(str(random.uniform(5.7, 6.2))).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_ml_4_1'] = "/"
            sampling_data['edta_ml_4_2'] = "/"
            sampling_data['edta_ml_5_1'] = "/"
            sampling_data['edta_ml_5_2'] = "/"

            # 处理edta_avg
            sampling_data['edta_avg_1'] = ((sampling_data['edta_ml_1_1'] + sampling_data['edta_ml_1_2']) / Decimal('2.0')).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_avg_2'] = ((sampling_data['edta_ml_2_1'] + sampling_data['edta_ml_2_2']) / Decimal('2.0')).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_avg_3'] = ((sampling_data['edta_ml_3_1'] + sampling_data['edta_ml_3_2']) / Decimal('2.0')).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_avg_4'] = '/'
            sampling_data['edta_avg_5'] = '/'

            # 处理material_per
            sampling_data['material_per_1'] = (Decimal('9.4302') - sampling_data['edta_avg_1'] * Decimal('0.7059')).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['material_per_2'] = (Decimal('9.4302') - sampling_data['edta_avg_2'] * Decimal('0.7059')).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['material_per_3'] = (Decimal('9.4302') - sampling_data['edta_avg_3'] * Decimal('0.7059')).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['material_per_4'] = '/'
            sampling_data['material_per_5'] = '/'


        elif point == "4":

            # 处理sample_id
            sampling_data['sample_id_1'] = "1"
            sampling_data['sample_id_2'] = "2"
            sampling_data['sample_id_3'] = "3"
            sampling_data['sample_id_4'] = "4"
            sampling_data['sample_id_5'] = "/"

            # 处理edta_ml
            sampling_data['edta_ml_1_1'] = Decimal(str(random.uniform(5.7, 6.2))).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_ml_1_2'] = Decimal(str(random.uniform(5.7, 6.2))).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_ml_2_1'] = Decimal(str(random.uniform(5.7, 6.2))).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_ml_2_2'] = Decimal(str(random.uniform(5.7, 6.2))).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_ml_3_1'] = Decimal(str(random.uniform(5.7, 6.2))).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_ml_3_2'] = Decimal(str(random.uniform(5.7, 6.2))).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_ml_4_1'] = Decimal(str(random.uniform(5.7, 6.2))).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_ml_4_2'] = Decimal(str(random.uniform(5.7, 6.2))).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_ml_5_1'] = "/"
            sampling_data['edta_ml_5_2'] = "/"

            # 处理edta_avg
            sampling_data['edta_avg_1'] = ((sampling_data['edta_ml_1_1'] + sampling_data['edta_ml_1_2']) / Decimal('2.0')).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_avg_2'] = ((sampling_data['edta_ml_2_1'] + sampling_data['edta_ml_2_2']) / Decimal('2.0')).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_avg_3'] = ((sampling_data['edta_ml_3_1'] + sampling_data['edta_ml_3_2']) / Decimal('2.0')).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_avg_4'] = ((sampling_data['edta_ml_4_1'] + sampling_data['edta_ml_4_2']) / Decimal('2.0')).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_avg_5'] = '/'

            # 处理material_per
            sampling_data['material_per_1'] = (Decimal('9.4302') - sampling_data['edta_avg_1'] * Decimal('0.7059')).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['material_per_2'] = (Decimal('9.4302') - sampling_data['edta_avg_2'] * Decimal('0.7059')).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['material_per_3'] = (Decimal('9.4302') - sampling_data['edta_avg_3'] * Decimal('0.7059')).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['material_per_4'] = (Decimal('9.4302') - sampling_data['edta_avg_4'] * Decimal('0.7059')).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['material_per_5'] = '/'


        elif point == "5":
            # 处理sample_id
            sampling_data['sample_id_1'] = "1"
            sampling_data['sample_id_2'] = "2"
            sampling_data['sample_id_3'] = "3"
            sampling_data['sample_id_4'] = "4"
            sampling_data['sample_id_5'] = "5"

            # 处理edta_ml
            sampling_data['edta_ml_1_1'] = Decimal(str(random.uniform(5.7, 6.2))).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_ml_1_2'] = Decimal(str(random.uniform(5.7, 6.2))).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_ml_2_1'] = Decimal(str(random.uniform(5.7, 6.2))).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_ml_2_2'] = Decimal(str(random.uniform(5.7, 6.2))).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_ml_3_1'] = Decimal(str(random.uniform(5.7, 6.2))).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_ml_3_2'] = Decimal(str(random.uniform(5.7, 6.2))).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_ml_4_1'] = Decimal(str(random.uniform(5.7, 6.2))).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_ml_4_2'] = Decimal(str(random.uniform(5.7, 6.2))).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_ml_5_1'] = Decimal(str(random.uniform(5.7, 6.2))).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_ml_5_2'] = Decimal(str(random.uniform(5.7, 6.2))).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)

            # 处理edta_avg
            sampling_data['edta_avg_1'] = ((sampling_data['edta_ml_1_1'] + sampling_data['edta_ml_1_2']) / Decimal('2.0')).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_avg_2'] = ((sampling_data['edta_ml_2_1'] + sampling_data['edta_ml_2_2']) / Decimal('2.0')).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_avg_3'] = ((sampling_data['edta_ml_3_1'] + sampling_data['edta_ml_3_2']) / Decimal('2.0')).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_avg_4'] = ((sampling_data['edta_ml_4_1'] + sampling_data['edta_ml_4_2']) / Decimal('2.0')).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['edta_avg_5'] = ((sampling_data['edta_ml_5_1'] + sampling_data['edta_ml_5_2']) / Decimal('2.0')).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)

            # 处理material_per
            sampling_data['material_per_1'] = (Decimal('9.4302') - sampling_data['edta_avg_1'] * Decimal('0.7059')).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['material_per_2'] = (Decimal('9.4302') - sampling_data['edta_avg_2'] * Decimal('0.7059')).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['material_per_3'] = (Decimal('9.4302') - sampling_data['edta_avg_3'] * Decimal('0.7059')).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['material_per_4'] = (Decimal('9.4302') - sampling_data['edta_avg_4'] * Decimal('0.7059')).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)
            sampling_data['material_per_5'] = (Decimal('9.4302') - sampling_data['edta_avg_5'] * Decimal('0.7059')).quantize(Decimal('0.1'), rounding=ROUND_HALF_UP)

        sampling_data_list.append(sampling_data)

    return assignment_id, report_id, project_name, witnessing_agency, assignment_type, assignment_party, contractor, sampling_date, report_date, project_section, witness, witness_id, sampler, sampler_id, test_type_list, sample_type, total_points, sampling_data_list, material_type


def update_report(assignment_id, report_id, project_name, witnessing_agency, assignment_type, assignment_party, contractor, sampling_date, report_date, project_section, witness, witness_id, sampler, sampler_id, test_type_list, sample_type, total_points,
                  sampling_data_list, material_type, reports_path):
    # 获取文件夹中的所有 Word 文件
    files = [f for f in os.listdir(reports_path) if f.endswith('.docx')]

    # 检查给定的报告编号列表长度是否与文件夹中的文件个数相等。
    if len(files) != len(assignment_id) or len(files) != len(report_id) or len(files) != len(project_name) or len(files) != len(witnessing_agency) or len(files) != len(assignment_type) or len(files) != len(assignment_party) or len(files) != len(
            contractor) or len(files) != len(sampling_date) or len(files) != len(report_date) or len(files) != len(project_section) or len(files) != len(witness) or len(files) != len(witness_id) or len(files) != len(sampler) or len(
        files) != len(sampler_id) or len(files) != len(test_type_list) or len(files) != len(sample_type) or len(
        files) != len(sampling_data_list) or len(files) != len(total_points):
        logger.error("文件数量与列表长度不匹配！")
        return

    for i, file_name in enumerate(files):
        try:
            file_path = os.path.join(reports_path, file_name)
            update_content(new_text=assignment_id[i], old_text="$assignment_id$", reports_path=file_path)
            update_content(new_text=material_type[i], old_text="$material_type$", reports_path=file_path)
            update_content(new_text=assignment_party[i], old_text="$assignment_party$", reports_path=file_path)
            update_content(new_text=project_name[i], old_text="$project_name$", reports_path=file_path)
            update_content(new_text=witnessing_agency[i], old_text="$witnessing_agency$", reports_path=file_path)
            update_content(new_text=assignment_type[i], old_text="$assignment_type$", reports_path=file_path)
            update_content(new_text=contractor[i], old_text="$contractor$", reports_path=file_path)
            update_content(new_text=sampling_date[i], old_text="$sampling_date$", reports_path=file_path)
            update_content(new_text=report_date[i], old_text="$report_date$", reports_path=file_path)
            update_content(new_text=report_id[i], old_text="$report_id$", reports_path=file_path)
            update_content(new_text=project_section[i], old_text="$project_section$", reports_path=file_path)
            update_content(new_text=witness[i], old_text="$witness$", reports_path=file_path)
            update_content(new_text=witness_id[i], old_text="$witness_id$", reports_path=file_path)
            update_content(new_text=sampler[i], old_text="$sampler$", reports_path=file_path)
            update_content(new_text=sampler_id[i], old_text="$sampler_id$", reports_path=file_path)
            update_content(new_text=sample_type[i], old_text="$sample_type$", reports_path=file_path)
            update_content(new_text=sample_type[i], old_text="$sample_type$", reports_path=file_path)

            for j in range(1, 6):
                update_content(new_text=str(sampling_data_list[i][f'sample_id_{j}']), old_text=f"$sample_id_{j}$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i][f'edta_ml_{j}_1']), old_text=f"$edta_ml_{j}_1$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i][f'edta_ml_{j}_2']), old_text=f"$edta_ml_{j}_2$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i][f'edta_avg_{j}']), old_text=f"$edta_avg_{j}$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i][f'material_per_{j}']), old_text=f"$material_per_{j}$", reports_path=file_path)

            logger.info(f"文件 {file_path} 中的所有数据已成功修改!")
        except ValueError as e:
            logger.error(f"修改报告内容时发生错误！错误信息：{e}")
    logger.success(f"所有报告已完成！输出目录 >>> {reports_path}")


if __name__ == '__main__':
    assignment_id, report_id, project_name, witnessing_agency, assignment_type, assignment_party, contractor, sampling_date, report_date, project_section, witness, witness_id, sampler, sampler_id, test_type_list, sample_type, total_points, sampling_data_list, material_type = load_data(
        param.huijiliang_source_data)
    reports_path = create_reports(report_id, project_name, test_type_list, param.huijiliang_template)
    update_report(assignment_id, report_id, project_name, witnessing_agency, assignment_type, assignment_party, contractor, sampling_date, report_date, project_section, witness, witness_id, sampler, sampler_id, test_type_list, sample_type, total_points,
                  sampling_data_list, material_type, reports_path)
