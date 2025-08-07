# -*- coding: utf-8 -*-
import pandas as pd
import os
import sys
import json

# 确保控制台能正确显示中文
try:
    if sys.platform.startswith('win'):
        os.system("chcp 65001")  # 设置Windows控制台编码为UTF-8
except:
    pass

# 文件夹路径
folder_path = "C:\\Users\\Administrator\\Desktop\\data\\"

# 定义model字段值的映射关系
MODEL_MAPPING = {
    "8002": "飞米1",
    "90": "DJI Air 3"
}


def is_blank_row(row):
    """判断一行是否为空行"""
    return all(pd.isna(cell) or str(cell).strip() == "" for cell in row)


def process_excel(file_path):
    try:
        file_name = os.path.basename(file_path)
        print(f"\n正在处理文件: {file_name}")

        # 生成与Excel同名的JSON文件路径
        file_base_name = os.path.splitext(file_name)[0]
        output_json_path = os.path.join(folder_path, f"{file_base_name}.json")

        # 根据文件扩展名选择合适的引擎
        if file_path.lower().endswith('.xlsx'):
            engine = 'openpyxl'  # xlsx格式强制使用openpyxl
        else:  # .xls格式
            engine = 'xlrd'  # xls格式使用xlrd

        # 读取Sheet2，不跳过任何行，不设置表头
        try:
            df_sheet2 = pd.read_excel(file_path, sheet_name='Sheet2', header=None, engine=engine)
            print(f"Sheet2 共{len(df_sheet2)}行数据")
        except Exception as e:
            print(f"读取Sheet2时出错: {str(e)}")
            return None

        # 存储转换后的结果
        result = []

        # 解析状态变量
        # 状态: initial(初始), found_blank(发现空白行), found_general_title(发现通用标题行),
        #       found_general_data(发现通用数据行), found_specific_title(发现具体标题行), processing_data(处理数据)
        current_state = "initial"
        current_general_data = {}  # 存储当前表格的通用数据
        current_specific_headers = []  # 存储当前表格的具体数据标题
        table_count = 0  # 统计发现的表格数量
        general_titles = []  # 存储通用数据的标题行

        # 遍历Sheet2的每一行
        for row_idx, row in df_sheet2.iterrows():
            # 获取第一列的值用于判断行类型
            first_col_value = str(row.iloc[0]).strip() if (len(row) > 0 and pd.notna(row.iloc[0])) else ""
            blank_row = is_blank_row(row)

            # 调试信息，可根据需要开启
            # print(f"行{row_idx+1} - 状态: {current_state}, 首列值: '{first_col_value}', 空行: {blank_row}")

            # 处理空白行
            if blank_row:
                # 空白行是表格的分隔符
                current_state = "found_blank"
                # 重置当前表格数据，但不重置结果，以便继续处理下一个表格
                current_general_data = {}
                current_specific_headers = []
                general_titles = []
                continue

            # 状态机处理逻辑
            if current_state in ["initial", "found_blank"]:
                # 寻找通用数据标题行（第一列是"首次发现"）
                if first_col_value == "首次发现":
                    # 记录通用数据标题行
                    general_titles = [str(cell).strip() if pd.notna(cell) else "" for cell in row]
                    current_state = "found_general_title"
                    # 记录新表格发现
                    table_count += 1
                    print(f"发现第{table_count}个表格 - 行号: {row_idx + 1}")

            elif current_state == "found_general_title":
                # 处理通用数据数据行（"首次发现"标题行的下一行）
                # 提取通用数据（根据标题匹配）
                current_general_data = {"other_general_data": {}}  # 可扩展存储其他通用数据

                # 遍历通用标题，匹配需要的字段
                for col_idx, title in enumerate(general_titles):
                    if col_idx >= len(row):
                        continue

                    cell_value = row.iloc[col_idx]
                    cell_str = str(cell_value).strip() if pd.notna(cell_value) else ""

                    # 根据标题匹配通用数据字段（移除了frequency的提取）
                    if title == "ID":
                        current_general_data["targetId"] = cell_str
                    elif title == "机型":
                        current_general_data["model"] = MODEL_MAPPING.get(cell_str, cell_str)

                current_state = "found_general_data"
                print(f"提取第{table_count}个表格的通用数据 - 行号: {row_idx + 1}")

            elif current_state == "found_general_data":
                # 寻找具体数据标题行（第一列是"发现时间"）
                if first_col_value == "发现时间":
                    current_specific_headers = [str(cell).strip() if pd.notna(cell) else "" for cell in row]
                    current_state = "found_specific_title"
                    print(f"提取第{table_count}个表格的具体数据标题 - 行号: {row_idx + 1}")

            elif current_state == "found_specific_title":
                # 处理具体数据行
                # 检查是否是有效的数据行（不是空白行且不是新的标题行）
                if not blank_row and first_col_value not in ["首次发现", "发现时间"]:
                    json_obj = {
                        # 通用数据
                        "targetId": current_general_data.get("targetId", ""),
                        "model": current_general_data.get("model", ""),
                        # 固定值
                        "dataSource": 5,
                        "droneType": 4,
                        # 默认空值（新增distance字段）
                        "groundSpeed": None,
                        "longitude": None,
                        "latitude": None,
                        "altitude": None,
                        "azimuth": None,
                        "createTime": "",
                        "pilotLongitude": None,
                        "pilotLatitude": None,
                        "deviceCode": "",
                        "timestamp": "",
                        "verticalSpeed": "",
                        "heading": None,
                        "pitch": None,
                        "roll": None,
                        "currentFlightStage": "",
                        "currentFlightDuration": "",
                        "coFlag": None,
                        "uasId": "",
                        "cardNum": "",
                        "distance": None,  # 新增distance字段
                        "frequency": ""  # frequency改为从具体数据提取
                    }

                    # createTime 对应具体数据行的第1列（索引0）
                    if len(row) > 0 and pd.notna(row.iloc[0]):
                        json_obj["createTime"] = str(row.iloc[0]).strip()

                    # 根据具体数据标题匹配其他数据
                    for col_idx, header in enumerate(current_specific_headers):
                        if col_idx >= len(row):
                            continue

                        cell_value = row.iloc[col_idx]
                        cell_str = str(cell_value).strip() if pd.notna(cell_value) else ""

                        # 根据表头匹配到对应的字段
                        if header == "经度":
                            try:
                                json_obj["longitude"] = float(cell_str) if cell_str else None
                            except ValueError:
                                json_obj["longitude"] = None
                        elif header == "纬度":
                            try:
                                json_obj["latitude"] = float(cell_str) if cell_str else None
                            except ValueError:
                                json_obj["latitude"] = None
                        elif header == "海拔高度":
                            try:
                                json_obj["altitude"] = float(cell_str) if cell_str else None
                            except ValueError:
                                json_obj["altitude"] = None
                        elif header == "飞行速度":
                            try:
                                json_obj["groundSpeed"] = float(cell_str) if cell_str else None
                            except ValueError:
                                json_obj["groundSpeed"] = None
                        elif header == "方位":
                            try:
                                json_obj["azimuth"] = float(cell_str) if cell_str else None
                            except ValueError:
                                json_obj["azimuth"] = None
                        elif header == "飞手经度":
                            try:
                                json_obj["pilotLongitude"] = float(cell_str) if cell_str else None
                            except ValueError:
                                json_obj["pilotLongitude"] = None
                        elif header == "飞手纬度":
                            try:
                                json_obj["pilotLatitude"] = float(cell_str) if cell_str else None
                            except ValueError:
                                json_obj["pilotLatitude"] = None
                        elif header == "设备ID":
                            json_obj["deviceCode"] = cell_str
                        elif header == "时间戳(ms)":
                            json_obj["timestamp"] = cell_str
                        elif header == "距离":  # 新增distance字段匹配
                            try:
                                json_obj["distance"] = float(cell_str) if cell_str else None
                            except ValueError:
                                json_obj["distance"] = None
                        elif header == "频率(Mhz)":  # frequency改为从具体数据提取
                            json_obj["frequency"] = cell_str

                    result.append(json_obj)
                elif first_col_value == "首次发现":
                    # 遇到新的表格标题行，说明当前表格已结束
                    table_count += 1
                    print(f"发现第{table_count}个表格 - 行号: {row_idx + 1}")
                    # 记录新表格的通用数据标题行
                    general_titles = [str(cell).strip() if pd.notna(cell) else "" for cell in row]
                    current_state = "found_general_title"
                    current_general_data = {}

        print(f"共发现 {table_count} 个表格")
        # 写入JSON文件
        with open(output_json_path, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)

        print(f"转换完成，共生成 {len(result)} 条数据")
        print(f"JSON文件已保存至: {output_json_path}")

        return result

    except Exception as e:
        print(f"\n处理文件 {file_path} 时出错: {str(e)}")
        return None


def process_all_excel_files(folder):
    """处理文件夹中所有的xls和xlsx文件"""
    if not os.path.exists(folder):
        print(f"文件夹不存在: {folder}")
        return

    # 获取文件夹中所有xls和xlsx文件
    excel_files = []
    for file in os.listdir(folder):
        if file.lower().endswith(('.xls', '.xlsx')):
            file_path = os.path.join(folder, file)
            if os.path.isfile(file_path):
                excel_files.append(file_path)

    if not excel_files:
        print(f"在 {folder} 中未找到任何xls或xlsx文件")
        return

    print(f"共发现 {len(excel_files)} 个Excel文件（xls/xlsx），开始批量处理...")

    # 遍历处理每个Excel文件
    for file_path in excel_files:
        process_excel(file_path)

    print("\n所有文件处理完毕")


# 执行批量处理
process_all_excel_files(folder_path)
