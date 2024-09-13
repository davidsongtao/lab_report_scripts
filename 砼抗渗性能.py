"""
Description: 
    
-*- Encoding: UTF-8 -*-
@File     ：砼抗渗性能脚本_已优化.py
@Author   ：King Songtao
@Time     ：2024/9/12 下午8:39
@Contact  ：king.songtao@gmail.com
"""
import csv
import os
from typing import Dict, List, Tuple
from utils import logger, convert_sampling_date_to_report_date, update_content, create_reports
from basic_config import ParametersConfig


def load_data(source_data: str) -> Dict[str, List[str]]:
    """从CSV文件中加载委托信息，构建数据字典"""
    data = {
        "委托日期": [], "委托编号": [], "报告编号": [], "样品名称": [], "检测项目": [],
        "检验类别": [], "委托单位": [], "工程名称": [], "工程部位": [], "强度等级": [],
        "抗渗等级": [], "养护条件": [], "成型日期": [], "龄期（天）": [], "代表批量（m3）": [],
        "见证员": [], "见证员证号": [], "取样员": [], "取样员证号": [], "施工单位": [],
        "见证单位": [], "报告领取": [], "备注": [],
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
        logger.success("数据加载成功！")
    except Exception as e:
        logger.error(f"数据加载错误！错误信息：{e}")

    return data


def process_data(data: Dict[str, List[str]]) -> Tuple[List[str], ...]:
    """处理加载的数据，返回所需的数据列表"""
    max_p_list = [str((int(grade[1:]) + 1) / 10) for grade in data["抗渗等级"]]

    return (
        data["委托编号"], data["报告编号"], data["委托单位"], data["工程名称"],
        data["见证单位"], data["养护条件"], data["强度等级"], data["抗渗等级"],
        data["工程部位"], data["施工单位"], data["检验类别"], data["委托日期"],
        data["成型日期"], convert_sampling_date_to_report_date(data["委托日期"], 1),
        data["代表批量（m3）"], data["龄期（天）"], data["见证员"], data["见证员证号"],
        data["取样员"], data["取样员证号"], data["检测项目"], max_p_list
    )


def update_report(data: Tuple[List[str], ...], reports_path: str):
    """更新砼试块抗渗检测报告内容"""
    files = [f for f in os.listdir(reports_path) if f.endswith('.docx')]
    if len(files) != len(data[0]):
        logger.error("文件数量与列表长度不匹配！")
        return

    placeholders = [
        "$assignment_id$", "$report_id$", "$assignment_party$", "$project_name$",
        "$witnessing_agency$", "$maintenance_condition$", "$strength_grade$", "$w_grade$",
        "$project_section$", "$contractor$", "$assignment_type$", "$sampling_date$",
        "$molding_date$", "$report_date$", "$batch_size$", "$curing_age$",
        "$witness$", "$witness_id$", "$sampler$", "$sampler_id$", "$test_type$", "$max_p$"
    ]

    for i, file_name in enumerate(files):
        try:
            file_path = os.path.join(reports_path, file_name)
            for j, placeholder in enumerate(placeholders):
                center = placeholder in ["$curing_age$", "$max_p$"]
                update_content(new_text=data[j][i], old_text=placeholder, center=center, reports_path=file_path)
            logger.info(f"检测报告 >>>{file_name}<<< 中的所有数据已成功修改!")
        except Exception as e:
            logger.error(f"修改报告内容时发生错误！错误信息：{e}")

    logger.success(f"所有报告已完成！输出目录 >>> {reports_path}")


def main():
    param = ParametersConfig()
    data = load_data(param.water_source_data)
    processed_data = process_data(data)
    reports_path = create_reports(processed_data[1], processed_data[3], processed_data[20], param.water_grade_template)
    update_report(processed_data, reports_path)


if __name__ == '__main__':
    main()
