"""
Description: 
    
-*- Encoding: UTF-8 -*-
@File     ：重构砼试块抗压.py
@Author   ：King Songtao
@Time     ：2024/9/13 上午7:06
@Contact  ：king.songtao@gmail.com
"""

import os
import csv
import random
from typing import Dict, List, Tuple, Any
from utils import logger, convert_sampling_date_to_report_date, update_content, create_reports
from basic_config import ParametersConfig


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
        return data
    except Exception as e:
        logger.error(f"数据加载错误！错误信息：{e}")
        return None


def process_data(data: Dict[str, List[str]]):
    """处理加载的数据，返回所需的数据列表"""

    report_date = convert_sampling_date_to_report_date(data["委托日期"], 1)
    break_power_1 = []
    break_power_2 = []
    break_power_3 = []
    ranges = {
        "C35": (401, 459), "C30": (381, 429), "C25": (301, 359), "C15": (201, 259),
        "C40": (481, 529), "C50": (581, 629), "C45": (501, 559), "C20": (281, 329),
    }

    for grade in data["强度等级"]:
        if grade in ranges:
            min_val, max_val = ranges[grade]
            break_power_1.append(str(round(random.uniform(min_val, max_val), 2)))
            break_power_2.append(str(round(random.uniform(min_val, max_val), 2)))
            break_power_3.append(str(round(random.uniform(min_val, max_val), 2)))
    strength_1 = []
    for power_1 in break_power_1:
        # 通过破坏荷载计算抗压强度

    return (
        data["委托日期"], data["委托编号"], data["报告编号"], data["样品名称"], data["检测项目"], data["委托单位"],
        data["工程名称"], data["工程部位"], data["强度等级"], data["成型日期"], data["龄期（天）"], data["代表批量（m3）"],
        data["养护条件"], data["见证员"], data["见证员证号"], data["取样员"], data["取样员证号"], data["施工单位"],
        data["见证单位"], data["见证日期"], data["报告领取"], data["备注"], data["检验类别"], report_date,
        break_power_1, break_power_2, break_power_3
    )


def main():
    param = ParametersConfig()
    data = load_data(param.concrete_strength_source_data)
    process_data(data)


if __name__ == '__main__':
    main()
