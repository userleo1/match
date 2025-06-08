import pymysql
from pymysql import Error

class BindTable:
    def __init__(self, project_id, quota_table, bind_table):
        self.project_id = project_id
        self.quota_table = quota_table
        self.bind_table = bind_table
        self.connection = None
        self.connect()
    
    def connect(self):
        """连接到数据库"""
        try:
            self.connection = pymysql.connect(
                host='localhost',
                user='root',
                password='L4512678.1',
                database='coast'
            )
        except Error as e:
            print(f"数据库连接失败: {e}")
            self.connection = None
    
    def disconnect(self):
        """断开数据库连接"""
        if self.connection:
            self.connection.close()
            self.connection = None
    
    def save_fix(self, name, specification, model, work_content, project_feature, code):
        """保存修正记录"""
        if not self.connection:
            return False
            
        try:
            # 计算条件哈希
            condition = f"{name}{specification}{model}{work_content}{project_feature}"
            condition_hash = hashlib.md5(condition.encode('utf-8')).hexdigest()
            
            cursor = self.connection.cursor()
            
            # 检查表中是否已有该条件的记录
            check_query = f"SELECT id FROM {self.bind_table} WHERE project_id = %s AND condition_hash = %s"
            cursor.execute(check_query, (self.project_id, condition_hash))
            existing_record = cursor.fetchone()
            
            if existing_record:
                # 更新记录
                update_query = f"UPDATE {self.bind_table} SET use_count = use_count + 1 WHERE id = %s"
                cursor.execute(update_query, (existing_record[0],))
            else:
                # 插入新记录
                # 获取定额表结构
                cursor.execute(f"SHOW COLUMNS FROM {self.quota_table}")
                quota_columns = [col[0] for col in cursor.fetchall()]
                
                # 从定额表中查询该编码的详细信息
                query = f"SELECT * FROM {self.quota_table} WHERE 定额编号 = %s"
                cursor.execute(query, (code,))
                quota_data = cursor.fetchone()
                
                if quota_data:
                    # 构建插入语句
                    insert_columns = "project_id, condition_hash, "
                    insert_values = "%s, %s, "
                    insert_params = [self.project_id, condition_hash]
                    
                    for col_idx, col_name in enumerate(quota_columns):
                        if col_name != 'id':  # 跳过id列
                            insert_columns += f"`{col_name}`, "
                            insert_values += "%s, "
                            insert_params.append(quota_data[col_idx])
                    
                    # 添加use_count
                    insert_columns += "use_count"
                    insert_values += "1"
                    
                    insert_query = f"INSERT INTO {self.bind_table} ({insert_columns}) VALUES ({insert_values})"
                    cursor.execute(insert_query, insert_params)
            
            self.connection.commit()
            return True
        except Error as e:
            print(f"保存修正记录失败: {e}")
            self.connection.rollback()
            return False
    
    def get_fix_by_condition(self, name, specification, model, work_content, project_feature):
        """根据条件获取修正记录"""
        if not self.connection:
            return None
            
        try:
            # 计算条件哈希
            condition = f"{name}{specification}{model}{work_content}{project_feature}"
            condition_hash = hashlib.md5(condition.encode('utf-8')).hexdigest()
            
            cursor = self.connection.cursor()
            
            # 从绑定表中查找
            query = f"SELECT * FROM {self.bind_table} WHERE project_id = %s AND condition_hash = %s ORDER BY use_count DESC LIMIT 1"
            cursor.execute(query, (self.project_id, condition_hash))
            result = cursor.fetchone()
            
            return result
        except Error as e:
            print(f"获取修正记录失败: {e}")
            return None    