import os
import json
import uuid
from database import Database

class ProjectManager:
    def __init__(self):
        self.db = Database()
    
    def create_project(self, project_name, quota_table):
        """创建新工程"""
        try:
            # 生成工程ID
            project_id = str(uuid.uuid4())
            
            # 创建绑定表
            bind_table = f"{quota_table}_bind"
            
            # 获取定额表结构
            columns = self.get_table_columns(quota_table)
            
            # 创建绑定表
            bind_columns = [
                'project_id VARCHAR(255) NOT NULL',
                'condition_hash CHAR(32) NOT NULL'
            ] + [f"`{col}` TEXT" for col in columns if col != 'id'] + ['use_count INT DEFAULT 0']
            
            if not self.db.table_exists(bind_table):
                create_table_sql = f"CREATE TABLE IF NOT EXISTS {bind_table} ("
                create_table_sql += "id INT AUTO_INCREMENT PRIMARY KEY, "
                create_table_sql += ", ".join(bind_columns)
                create_table_sql += ", UNIQUE KEY (project_id, condition_hash)"
                create_table_sql += ")"
                
                self.db.execute_query(create_table_sql)
            
            # 创建工程对象
            project = {
                'name': project_name,
                'id': project_id,
                'quota_table': quota_table,
                'bind_table': bind_table,
                'match_file': None,
                'match_data': None
            }
            
            return project
        except Exception as e:
            print(f"创建工程失败: {e}")
            return None
    
    def save_project(self, project, file_path):
        """保存工程"""
        try:
            with open(file_path, 'w') as f:
                json.dump(project, f)
            return True
        except Exception as e:
            print(f"保存工程失败: {e}")
            return False
    
    def load_project(self, file_path):
        """加载工程"""
        try:
            with open(file_path, 'r') as f:
                project = json.load(f)
            
            # 检查绑定表是否存在
            if not self.db.table_exists(project['bind_table']):
                # 重新创建绑定表
                quota_table = project['quota_table']
                columns = self.get_table_columns(quota_table)
                
                bind_columns = [
                    'project_id VARCHAR(255) NOT NULL',
                    'condition_hash CHAR(32) NOT NULL'
                ] + [f"`{col}` TEXT" for col in columns if col != 'id'] + ['use_count INT DEFAULT 0']
                
                create_table_sql = f"CREATE TABLE IF NOT EXISTS {project['bind_table']} ("
                create_table_sql += "id INT AUTO_INCREMENT PRIMARY KEY, "
                create_table_sql += ", ".join(bind_columns)
                create_table_sql += ", UNIQUE KEY (project_id, condition_hash)"
                create_table_sql += ")"
                
                self.db.execute_query(create_table_sql)
            
            # 添加文件路径到工程数据
            project['file_path'] = file_path
            
            return project
        except Exception as e:
            print(f"加载工程失败: {e}")
            return None
    
    def get_table_columns(self, table_name):
        """获取表的列名"""
        try:
            result = self.db.execute_query(f"SHOW COLUMNS FROM {table_name}")
            return [row[0] for row in result]
        except Exception as e:
            print(f"获取表列名失败: {e}")
            return []    