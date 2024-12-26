"""
Description: 全局配置文件，配置全局参数
    
-*- Encoding: UTF-8 -*-
@File     ：basic_config.py
@Author   ：King Songtao
@Time     ：2024/9/9 下午11:24
@Contact  ：king.songtao@gmail.com
"""


class ParametersConfig(object):
    def __init__(self):
        # 砼试块抗压强度 -> 相关配置参数
        self.concrete_strength_template = r"E:\lab_report_scripts\templates\砼试块抗压强度模板.docx"
        self.concrete_strength_source_data = r"E:\lab_report_scripts\source_data\砼试块抗压强度.csv"
        self.output_directory = r"E:\检测报告生成\砼试块抗压强度"
        self.log_directory = r"E:\lab_report_scripts\logs"
        self.rebar_welding_source_data = r"E:\lab_report_scripts\source_data\钢筋焊接数据源.csv"
        self.rebar_welding_template = r"E:\lab_report_scripts\templates\电渣压力焊模板.docx"
        self.steel_source_data = r"E:\lab_report_scripts\source_data\钢筋原材数据源.csv"
        self.steel_template = r"E:\lab_report_scripts\templates\钢筋原材模板.docx"
        self.water_source_data = r"E:\lab_report_scripts\source_data\砼抗渗性能数据源.csv"
        self.water_grade_template = r"E:\lab_report_scripts\templates\砼抗渗性能模板.docx"
        self.geotechnical_density_source_data = r"E:\lab_report_scripts\source_data\土工密度源文件.csv"
        self.geotechnical_density_template = r"E:\lab_report_scripts\templates\土工密度模板.docx"
        self.huijiliang_source_data = r"E:\lab_report_scripts\source_data\灰剂量源文件.csv"
        self.huijiliang_template = r"E:\lab_report_scripts\templates\灰剂量模板.docx"
