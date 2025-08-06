# -*- coding: utf-8 -*-
import pandas as pd
import os
import sys
import json
import xml.etree.ElementTree as ET

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

        # 尝试读取Sheet1
        try:
            df_sheet1 = pd.read_excel(file_path, sheet_name='Sheet1', header=0, engine=engine)
            print(f"Sheet1 共{len(df_sheet1)}行数据")
        except Exception as e:
            print(f"读取Sheet1时出错: {str(e)}")
            return None

        # 提取Sheet1中的公共属性
        sheet1_common = df_sheet1.iloc[0] if len(df_sheet1) > 0 else None

        # 读取Sheet2
        try:
            df_sheet2 = pd.read_excel(file_path, sheet_name='Sheet2', skiprows=3, header=0, engine=engine)
            print(f"Sheet2 共{len(df_sheet2)}行数据（从第5行开始）")
        except Exception as e:
            print(f"读取Sheet2时出错: {str(e)}")
            return None

        # 获取Sheet2的列数
        sheet2_columns_count = len(df_sheet2.columns)

        # 存储转换后的结果
        result = []

        # 处理Sheet2中的每一行数据
        for sheet2_idx, sheet2_row in df_sheet2.iterrows():
            # 获取原始model值并转换
            original_model = str(sheet1_common.iloc[10]) if (
                    sheet1_common is not None and pd.notna(sheet1_common.iloc[10])) else ""
            converted_model = MODEL_MAPPING.get(original_model, original_model)

            # 构建JSON对象
            json_obj = {
                "targetId": str(sheet1_common.iloc[2]) if (
                        sheet1_common is not None and pd.notna(sheet1_common.iloc[2])) else "",
                "model": converted_model,
                "dataSource": 5,
                "droneType": 4,
                "groundSpeed": float(sheet2_row.iloc[6]) if pd.notna(sheet2_row.iloc[6]) else None,
                "longitude": float(sheet2_row.iloc[3]) if pd.notna(sheet2_row.iloc[3]) else None,
                "latitude": float(sheet2_row.iloc[4]) if pd.notna(sheet2_row.iloc[4]) else None,
                "altitude": float(sheet2_row.iloc[5]) if pd.notna(sheet2_row.iloc[5]) else None,
                "azimuth": float(sheet2_row.iloc[7]) if pd.notna(sheet2_row.iloc[7]) else None,
                "createTime": str(sheet2_row.iloc[0]) if pd.notna(sheet2_row.iloc[0]) else "",
                "pilotLongitude": float(sheet2_row.iloc[10]) if pd.notna(sheet2_row.iloc[10]) else None,
                "pilotLatitude": float(sheet2_row.iloc[11]) if pd.notna(sheet2_row.iloc[11]) else None,
                "frequency": str(sheet2_row.iloc[sheet2_columns_count - 5]) if pd.notna(
                    sheet2_row.iloc[sheet2_columns_count - 5]) else "",
                "deviceCode": str(sheet2_row.iloc[sheet2_columns_count - 2]) if pd.notna(
                    sheet2_row.iloc[sheet2_columns_count - 2]) else "",
                "timestamp": str(sheet2_row.iloc[sheet2_columns_count - 1]) if pd.notna(
                    sheet2_row.iloc[sheet2_columns_count - 1]) else "",
                "verticalSpeed": "",
                "heading": None,
                "pitch": None,
                "roll": None,
                "currentFlightStage": "",
                "currentFlightDuration": "",
                "coFlag": None,
                "uasId": "",
                "cardNum": ""
            }

            result.append(json_obj)

        # 写入JSON文件
        with open(output_json_path, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)

        print(f"转换完成，共生成 {len(result)} 条数据")
        print(f"JSON文件已保存至: {output_json_path}")

        return result

    except ET.ParseError as e:
        print(f"\n文件 {file_path} 存在XML解析错误（可能损坏）: {str(e)}")
        print("建议修复方法:")
        print("1. 尝试用Excel打开文件并另存为新文件")
        print("2. 将数据复制到新的Excel文件中")
        return None
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
