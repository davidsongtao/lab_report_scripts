"""
Description: 
    
-*- Encoding: UTF-8 -*-
@File     ：土工密度重构.py
@Author   ：King Songtao
@Time     ：2024/12/16 上午10:01
@Contact  ：king.songtao@gmail.com
"""
"""
Description: 
    Geotechnical Density Report Generator
@File     ：土工密度.py
@Author   ：King Songtao
@Time     ：2024/12/15 下午9:51
@Contact  ：king.songtao@gmail.com
"""
import csv
import os
import random
import logging
from typing import List, Dict, Tuple, Any
from utils import *


class GeotechnicalDensityReportGenerator:
    def __init__(self, source_data_path: str, template_path: str):
        """
        Initialize the report generator with source data and template paths

        :param source_data_path: Path to the source CSV data file
        :param template_path: Path to the Word document template
        """
        self.source_data_path = source_data_path
        self.template_path = template_path
        self.data_loader = self._initialize_data_loader()

    def _initialize_data_loader(self) -> Dict[str, List[str]]:
        """
        Create a dictionary to store all data fields from the source CSV

        :return: Dictionary with empty lists for each data field
        """
        fields = [
            "委托日期", "委托编号", "报告编号", "土壤种类", "检测项目", "检验类别",
            "委托单位", "工程名称", "工程部位", "总取样点数", "最大干密度",
            "设计压实度", "见证员", "见证员证号", "取样员", "取样员证号",
            "施工单位", "见证单位"
        ]

        # Add sampling site fields dynamically
        fields.extend([f"取样部位（{i}）" for i in range(1, 13)])

        # Additional fields
        fields.extend(["报告领取", "备注", "设计压实度（百分比）", "实验数据"])

        return {field: [] for field in fields}

    def load_data(self) -> Tuple[Any, ...]:
        """
        Load data from the source CSV file

        :return: Tuple of processed data for report generation
        """
        try:
            with open(self.source_data_path, "r", encoding="GBK") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    for key, value in row.items():
                        if key in self.data_loader:
                            # Append semicolon for witness and sampler names
                            value = value + "；" if key in ["见证员", "取样员"] else value
                            self.data_loader[key].append(value)

                logger.info("Data loaded successfully!")

            return self._process_data()

        except Exception as e:
            logger.error(f"Data loading error: {e}")
            raise

    def _process_data(self) -> Tuple[Any, ...]:
        """
        Process and transform the loaded data

        :return: Processed data tuple for report generation
        """
        # Convert design compaction to percentage
        compaction_percentage = [
            f"{int(float(value) * 100)}%" for value in self.data_loader["设计压实度"]
        ]

        # Generate sampling data with randomized values
        sampling_data_list = self._generate_sampling_data(
            self.data_loader["最大干密度"],
            self.data_loader["设计压实度"],
            self.data_loader["总取样点数"]
        )

        # Prepare sampling sites dynamically
        sampling_sites = [
            self.data_loader[f"取样部位（{i}）"] for i in range(1, 13)
        ]

        # Combine all data to return
        return (
            self.data_loader["委托编号"],
            self.data_loader["报告编号"],
            self.data_loader["工程名称"],
            # ... (continue with the rest of the data fields)
            sampling_sites[0], sampling_sites[1], sampling_sites[2],
            sampling_sites[3], sampling_sites[4], sampling_sites[5],
            sampling_sites[6], sampling_sites[7], sampling_sites[8],
            sampling_sites[9], sampling_sites[10], sampling_sites[11]
        )

    def _generate_sampling_data(self, max_densities: List[str],
                                compactions: List[str],
                                total_sample_points: List[str]) -> List[Dict[str, float]]:
        """
        Generate randomized sampling data based on parameters

        :param max_densities: List of maximum densities
        :param compactions: List of compaction rates
        :param total_sample_points: List of total sampling points
        :return: List of sampling data dictionaries
        """
        sampling_data_list = []

        for index, point_count in enumerate(total_sample_points):
            sampling_data = {}
            point_count = int(point_count)
            max_density = float(max_densities[index])
            compaction_rate = float(compactions[index])

            for i in range(1, point_count + 1):
                # Wet density: between 1.71 and 1.86
                wet_density = round(random.uniform(1.71, 1.86), 2)

                # Water content: between 14.7 and 15.6
                water_content = round(random.uniform(14.7, 15.6), 1)

                # Dry density calculation
                dry_lower_bound = max_density * (compaction_rate + 0.01)
                dry_upper_bound = max_density * 0.99
                dry_density = round(random.uniform(dry_lower_bound, dry_upper_bound), 2)

                # Compaction coefficient
                compaction_coef = round(dry_density / max_density * 100, 1)

                sampling_data.update({
                    f'wet_{i}': wet_density,
                    f'water_{i}': water_content,
                    f'dry_{i}': dry_density,
                    f'com_{i}': compaction_coef
                })

            sampling_data_list.append(sampling_data)

        return sampling_data_list

    def update_reports(self, output_dir: str):
        """
        Update all report files in the output directory

        :param output_dir: Directory containing report files
        """
        files = [f for f in os.listdir(output_dir) if f.endswith('.docx')]

        for i, file_name in enumerate(files):
            file_path = os.path.join(output_dir, file_name)
            try:
                self._update_report_content(file_path, i)
                logger.info(f"Successfully updated report: {file_path}")
            except Exception as e:
                logger.error(f"Error updating report {file_path}: {e}")

    def _update_report_content(self, file_path: str, index: int):
        """
        Update content of a single report file

        :param file_path: Path to the report file
        :param index: Index of current data row
        """
        placeholders = {
            "$assignment_id$": self.data_loader["委托编号"][index],
            "$assignment_party$": self.data_loader["委托单位"][index],
            # ... (add more placeholders)
        }

        # Dynamically update content
        for placeholder, value in placeholders.items():
            self._replace_content(file_path, placeholder, str(value))

        # Handle sampling data
        self._handle_sampling_data(file_path, index)

    def _replace_content(self, file_path: str, old_text: str, new_text: str):
        """
        Replace content in a file

        :param file_path: Path to the file
        :param old_text: Text to be replaced
        :param new_text: Replacement text
        """
        # Implement your content replacement logic here
        pass

    def _handle_sampling_data(self, file_path: str, index: int):
        """
        Handle sampling data placeholders in the report

        :param file_path: Path to the report file
        :param index: Index of current data row
        """
        # Implement sampling data placeholder replacement
        pass


def main():
    """
    Main execution function for report generation
    """
    try:
        # Replace with your actual parameter paths
        source_data_path = param.geotechnical_density_source_data
        template_path = param.geotechnical_density_template

        generator = GeotechnicalDensityReportGenerator(source_data_path, template_path)

        # Generate reports
        data = generator.load_data()
        output_dir = create_reports(data[1], data[2], data[14], template_path)

        # Update reports
        generator.update_reports(output_dir)

        logger.info("Report generation completed successfully!")

    except Exception as e:
        logger.error(f"Report generation failed: {e}")


if __name__ == '__main__':
    main()
