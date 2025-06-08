import pandas as pd
import openpyxl
from database import Database


class QuotaImporter:
    def __init__(self):
        self.db = Database()

    def generate_template(self, file_path):
        """生成定额导入模板"""
        try:
            # 创建Excel文件
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = '定额模板'

            # 设置表头
            headers = [
                '定额编号', '分部分项工程名称', '计量单位', '工程量',
                '主材费', '小计', '人工费', '材料费', '机械费', '主材费', '合计', '人工费', '材料费', '机械费'
            ]
            for col_idx, header in enumerate(headers, 1):
                ws.cell(row=1, column=col_idx, value=header)

            # 保存文件
            wb.save(file_path)
            return True
        except Exception as e:
            print(f"生成模板失败: {e}")
            return False

    def import_quota(self, file_path, table_name):
        """导入定额数据"""
        try:
            # 读取Excel文件
            df = pd.read_excel(file_path)

            # 验证模板
            required_headers = [
                '定额编号', '分部分项工程名称', '计量单位', '工程量',
                '主材费', '小计', '人工费', '材料费', '机械费'
            ]

            if list(df.columns) != required_headers:
                return False, '请使用模板导入'

            # 检查表是否存在
            if self.db.table_exists(table_name):
                # 检查表中是否已有数据
                result = self.db.execute_query(f"SELECT COUNT(*) FROM {table_name}")
                count = result[0][0] if result else 0

                if count > 0:
                    # 检查导入的数据是否与已有数据重复
                    first_code = df.iloc[0]['定额编号']
                    result = self.db.execute_query(f"SELECT COUNT(*) FROM {table_name} WHERE 定额编号 = '{first_code}'")
                    code_count = result[0][0] if result else 0

                    if code_count > 0:
                        return False, '定额已存在'

            # 创建表（如果不存在）
            if not self.db.table_exists(table_name):
                self.db.create_table(table_name, df.columns)

            # 插入数据
            data = df.to_dict('records')
            success = self.db.insert_data(table_name, data)

            if success:
                return True, '导入成功'
            else:
                return False, '插入数据失败'
        except Exception as e:
            print(f"导入定额失败: {e}")
            return False, str(e)