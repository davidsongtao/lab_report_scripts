"""
Description: 
    
-*- Encoding: UTF-8 -*-
@File     ：土工密度.py
@Author   ：King Songtao
@Time     ：2024/12/15 下午9:51
@Contact  ：king.songtao@gmail.com
"""
from utils import *
import csv
import random


def load_data(source_data):
    """
    加载土工密度委托台账，获取数据
    :param source_data:
    :return:
    """
    data_loader = {
        "委托日期": [],
        "委托编号": [],
        "报告编号": [],
        "土壤种类": [],
        "检测项目": [],
        "检验类别": [],
        "委托单位": [],
        "工程名称": [],
        "工程部位": [],
        "总取样点数": [],
        "最大干密度": [],
        "设计压实度": [],
        "见证员": [],
        "见证员证号": [],
        "取样员": [],
        "取样员证号": [],
        "施工单位": [],
        "见证单位": [],
        "取样部位（1）": [],
        "取样部位（2）": [],
        "取样部位（3）": [],
        "取样部位（4）": [],
        "取样部位（5）": [],
        "取样部位（6）": [],
        "取样部位（7）": [],
        "取样部位（8）": [],
        "取样部位（9）": [],
        "取样部位（10）": [],
        "取样部位（11）": [],
        "取样部位（12）": [],
        "报告领取": [],
        "备注": [],
        "设计压实度（百分比）": [],
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
    max_density = data_loader["最大干密度"]
    compaction = data_loader["设计压实度"]
    total_sample_points = data_loader["总取样点数"]
    sampling_site_1 = data_loader["取样部位（1）"]
    sampling_site_2 = data_loader["取样部位（2）"]
    sampling_site_3 = data_loader["取样部位（3）"]
    sampling_site_4 = data_loader["取样部位（4）"]
    sampling_site_5 = data_loader["取样部位（5）"]
    sampling_site_6 = data_loader["取样部位（6）"]
    sampling_site_7 = data_loader["取样部位（7）"]
    sampling_site_8 = data_loader["取样部位（8）"]
    sampling_site_9 = data_loader["取样部位（9）"]
    sampling_site_10 = data_loader["取样部位（10）"]
    sampling_site_11 = data_loader["取样部位（11）"]
    sampling_site_12 = data_loader["取样部位（12）"]

    # 将设计压实度转换为百分比，取整
    compaction_percentage = [f"{int(float(value) * 100)}%" for value in compaction]

    # 构建取样部位及实验数据
    sampling_data_list = []

    for index, point in enumerate(total_sample_points):

        sampling_data = {}
        for i in range(1, int(point) + 1):
            # 湿密度：1.71至1.86之间，保留两位小数
            sampling_data[f'wet_{i}'] = round(random.uniform(1.71, 1.86), 2)

            # 含水率：14.7至15.6之间，保留一位小数
            sampling_data[f'water_{i}'] = round(random.uniform(14.7, 15.6), 1)

            # 干密度：在{max_density-0.8}至{max_density-0.1}之间，保留两位小数
            dry_lower_bound = float(max_density[index]) * (float(compaction[index]) + 0.01)
            dry_upper_bound = float(max_density[index]) * 0.99
            sampling_data[f'dry_{i}'] = round(random.uniform(dry_lower_bound, dry_upper_bound), 2)

            # 压实系数：干密度除以最大干密度，保留一位小数的百分比
            sampling_data[f'com_{i}'] = round(sampling_data[f'dry_{i}'] / float(max_density[index]) * 100, 1)

        sampling_data_list.append(sampling_data)

    return assignment_id, report_id, project_name, witnessing_agency, assignment_type, assignment_party, contractor, sampling_date, report_date, project_section, witness, witness_id, sampler, sampler_id, test_type_list, sample_type, max_density, compaction, total_sample_points, compaction_percentage, sampling_data_list, sampling_site_1, sampling_site_2, sampling_site_3, sampling_site_4, sampling_site_5, sampling_site_6, sampling_site_7, sampling_site_8, sampling_site_9, sampling_site_10, sampling_site_11, sampling_site_12


def update_report(assignment_id, report_id, project_name, witnessing_agency, assignment_type, assignment_party, contractor, sampling_date, report_date, project_section, witness, witness_id, sampler, sampler_id, test_type_list, sample_type, max_density,
                  compaction, total_sample_points, compaction_percentage, sampling_data_list, sampling_site_1, sampling_site_2, sampling_site_3, sampling_site_4, sampling_site_5, sampling_site_6, sampling_site_7, sampling_site_8, sampling_site_9,
                  sampling_site_10, sampling_site_11, sampling_site_12, reports_path):
    # 获取文件夹中的所有 Word 文件
    files = [f for f in os.listdir(reports_path) if f.endswith('.docx')]

    # 检查给定的报告编号列表长度是否与文件夹中的文件个数相等。
    if len(files) != len(assignment_id) or len(files) != len(report_id) or len(files) != len(project_name) or len(files) != len(witnessing_agency) or len(files) != len(assignment_type) or len(files) != len(assignment_party) or len(files) != len(
            contractor) or len(files) != len(sampling_date) or len(files) != len(report_date) or len(files) != len(project_section) or len(files) != len(witness) or len(files) != len(witness_id) or len(files) != len(sampler) or len(
        files) != len(sampler_id) or len(files) != len(test_type_list) or len(files) != len(sample_type) or len(files) != len(max_density) or len(files) != len(compaction) or len(files) != len(total_sample_points) or len(files) != len(
        compaction_percentage) or len(
        files) != len(sampling_data_list):
        logger.error("文件数量与列表长度不匹配！")
        return

    # 若文件数量与列表长度匹配，则更新所有文档的占位符
    for i, file_name in enumerate(files):
        try:
            file_path = os.path.join(reports_path, file_name)
            update_content(new_text=assignment_id[i], old_text="$assignment_id$", reports_path=file_path)
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
            update_content(new_text=max_density[i], old_text="$max_density$", reports_path=file_path)
            update_content(new_text=compaction_percentage[i], old_text="$compaction$", reports_path=file_path)
            # for j in range(len(sampling_data_list)):
            #     # 定义需要处理的字段列表
            #     fields = ['wet', 'water', 'dry', 'com']
            #     total_sites = 12
            #
            #     # 使用globals()动态获取采样点站点变量
            #     sampling_sites = [
            #         globals()[f'sampling_site_{site_num}'][j]
            #         for site_num in range(1, total_sites + 1)
            #     ]
            #
            #     # 遍历所有可能的站点
            #     for site_num in range(1, total_sites + 1):
            #         # 检查当前站点是否有数据
            #         has_data = sampling_data_list[j].get(f'wet_{site_num}', False)
            #
            #         if has_data:
            #             # 如果有数据，更新相应的占位符
            #             update_content(new_text=sampling_sites[site_num - 1], old_text=f"$sampling_site_{site_num}$", reports_path=file_path)
            #
            #             for field in fields:
            #                 update_content(new_text=str(sampling_data_list[j][f'{field}_{site_num}']),
            #                                old_text=f"${field}_{site_num}$",
            #                                reports_path=file_path)
            #         else:
            #             # 如果没有数据，将占位符替换为空
            #             update_content(new_text="", old_text=f"$sampling_site_{site_num}$", reports_path=file_path)
            #
            #             for field in fields:
            #                 update_content(new_text="", old_text=f"${field}_{site_num}$", reports_path=file_path)

            if "wet_1" in sampling_data_list[i]:
                update_content(new_text=sampling_site_1[i], old_text="$sampling_site_1$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["wet_1"]), old_text="$wet_1$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["water_1"]), old_text="$water_1$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["dry_1"]), old_text="$dry_1$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["com_1"]), old_text="$com_1$", reports_path=file_path)
            if "wet_2" in sampling_data_list[i]:
                update_content(new_text=sampling_site_2[i], old_text="$sampling_site_2$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["wet_2"]), old_text="$wet_2$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["water_2"]), old_text="$water_2$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["dry_2"]), old_text="$dry_2$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["com_2"]), old_text="$com_2$", reports_path=file_path)
            if "wet_3" in sampling_data_list[i]:
                update_content(new_text=sampling_site_3[i], old_text="$sampling_site_3$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["wet_3"]), old_text="$wet_3$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["water_3"]), old_text="$water_3$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["dry_3"]), old_text="$dry_3$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["com_3"]), old_text="$com_3$", reports_path=file_path)
            if "wet_4" in sampling_data_list[i]:
                update_content(new_text=sampling_site_4[i], old_text="$sampling_site_4$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["wet_4"]), old_text="$wet_4$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["water_4"]), old_text="$water_4$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["dry_4"]), old_text="$dry_4$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["com_4"]), old_text="$com_4$", reports_path=file_path)
            if "wet_5" in sampling_data_list[i]:
                update_content(new_text=sampling_site_5[i], old_text="$sampling_site_5$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["wet_5"]), old_text="$wet_5$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["water_5"]), old_text="$water_5$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["dry_5"]), old_text="$dry_5$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["com_5"]), old_text="$com_5$", reports_path=file_path)
            if "wet_6" in sampling_data_list[i]:
                update_content(new_text=sampling_site_6[i], old_text="$sampling_site_6$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["wet_6"]), old_text="$wet_6$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["water_6"]), old_text="$water_6$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["dry_6"]), old_text="$dry_6$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["com_6"]), old_text="$com_6$", reports_path=file_path)
            if "wet_7" in sampling_data_list[i]:
                update_content(new_text=sampling_site_7[i], old_text="$sampling_site_7$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["wet_7"]), old_text="$wet_7$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["water_7"]), old_text="$water_7$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["dry_7"]), old_text="$dry_7$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["com_7"]), old_text="$com_7$", reports_path=file_path)
            if "wet_8" in sampling_data_list[i]:
                update_content(new_text=sampling_site_8[i], old_text="$sampling_site_8$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["wet_8"]), old_text="$wet_8$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["water_8"]), old_text="$water_8$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["dry_8"]), old_text="$dry_8$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["com_8"]), old_text="$com_8$", reports_path=file_path)
            if "wet_9" in sampling_data_list[i]:
                update_content(new_text=sampling_site_9[i], old_text="$sampling_site_9$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["wet_9"]), old_text="$wet_9$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["water_9"]), old_text="$water_9$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["dry_9"]), old_text="$dry_9$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["com_9"]), old_text="$com_9$", reports_path=file_path)
            if "wet_10" in sampling_data_list[i]:
                update_content(new_text=sampling_site_10[i], old_text="$sampling_site_10$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["wet_10"]), old_text="$wet_10$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["water_10"]), old_text="$water_10$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["dry_10"]), old_text="$dry_10$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["com_10"]), old_text="$com_10$", reports_path=file_path)
            if "wet_11" in sampling_data_list[i]:
                update_content(new_text=sampling_site_11[i], old_text="$sampling_site_11$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["wet_11"]), old_text="$wet_11$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["water_11"]), old_text="$water_11$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["dry_11"]), old_text="$dry_11$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["com_11"]), old_text="$com_11$", reports_path=file_path)
            if "wet_12" in sampling_data_list[i]:
                update_content(new_text=sampling_site_12[i], old_text="$sampling_site_12$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["wet_12"]), old_text="$wet_12$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["water_12"]), old_text="$water_12$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["dry_12"]), old_text="$dry_12$", reports_path=file_path)
                update_content(new_text=str(sampling_data_list[i]["com_12"]), old_text="$com_12$", reports_path=file_path)
            elif "wet_1" in sampling_data_list[i] and "wet_2" not in sampling_data_list[i]:
                update_content(new_text="(以下空白)", old_text="$sampling_site_2$", reports_path=file_path)
            elif "wet_2" in sampling_data_list[i] and "wet_3" not in sampling_data_list[i]:
                update_content(new_text="(以下空白)", old_text="$sampling_site_3$", reports_path=file_path)
            elif "wet_3" in sampling_data_list[i] and "wet_4" not in sampling_data_list[i]:
                update_content(new_text="(以下空白)", old_text="$sampling_site_4$", reports_path=file_path)
            elif "wet_4" in sampling_data_list[i] and "wet_5" not in sampling_data_list[i]:
                update_content(new_text="(以下空白)", old_text="$sampling_site_5$", reports_path=file_path)
            elif "wet_5" in sampling_data_list[i] and "wet_6" not in sampling_data_list[i]:
                update_content(new_text="(以下空白)", old_text="$sampling_site_6$", reports_path=file_path)
            elif "wet_6" in sampling_data_list[i] and "wet_7" not in sampling_data_list[i]:
                update_content(new_text="(以下空白)", old_text="$sampling_site_7$", reports_path=file_path)
            elif "wet_7" in sampling_data_list[i] and "wet_8" not in sampling_data_list[i]:
                update_content(new_text="(以下空白)", old_text="$sampling_site_8$", reports_path=file_path)
            elif "wet_8" in sampling_data_list[i] and "wet_9" not in sampling_data_list[i]:
                update_content(new_text="(以下空白)", old_text="$sampling_site_9$", reports_path=file_path)
            elif "wet_9" in sampling_data_list[i] and "wet_10" not in sampling_data_list[i]:
                update_content(new_text="(以下空白)", old_text="$sampling_site_10$", reports_path=file_path)
            elif "wet_10" in sampling_data_list[i] and "wet_11" not in sampling_data_list[i]:
                update_content(new_text="(以下空白)", old_text="$sampling_site_11$", reports_path=file_path)
            elif "wet_11" in sampling_data_list[i] and "wet_12" not in sampling_data_list[i]:
                update_content(new_text="(以下空白)", old_text="$sampling_site_12$", reports_path=file_path)

            for k in range(1, 13):
                update_content(new_text="", old_text=f"$sampling_site_{k}$", reports_path=file_path)
                update_content(new_text="", old_text=f"$wet_{k}$", reports_path=file_path)
                update_content(new_text="", old_text=f"$water_{k}$", reports_path=file_path)
                update_content(new_text="", old_text=f"$dry_{k}$", reports_path=file_path)
                update_content(new_text="", old_text=f"$com_{k}$", reports_path=file_path)

            logger.info(f"文件 {file_path} 中的所有数据已成功修改!")

        except ValueError as e:
            logger.error(f"修改报告内容时发生错误！错误信息：{e}")

    logger.success(f"所有报告已完成！输出目录 >>> {reports_path}")


if __name__ == '__main__':
    assignment_id, report_id, project_name, witnessing_agency, assignment_type, assignment_party, contractor, sampling_date, report_date, project_section, witness, witness_id, sampler, sampler_id, test_type_list, sample_type, max_density, compaction, total_sample_points, compaction_percentage, sampling_data_list, sampling_site_1, sampling_site_2, sampling_site_3, sampling_site_4, sampling_site_5, sampling_site_6, sampling_site_7, sampling_site_8, sampling_site_9, sampling_site_10, sampling_site_11, sampling_site_12 = load_data(
        param.geotechnical_density_source_data)

    output_dir = create_reports(report_id, project_name, test_type_list, param.geotechnical_density_template)
    update_report(assignment_id, report_id, project_name, witnessing_agency, assignment_type, assignment_party, contractor, sampling_date, report_date, project_section, witness, witness_id, sampler, sampler_id, test_type_list, sample_type, max_density,
                  compaction, total_sample_points, compaction_percentage, sampling_data_list, sampling_site_1, sampling_site_2, sampling_site_3, sampling_site_4, sampling_site_5, sampling_site_6, sampling_site_7, sampling_site_8, sampling_site_9,
                  sampling_site_10, sampling_site_11, sampling_site_12, output_dir)
