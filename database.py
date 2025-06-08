import pymysql
from pymysql import Error

class Database:
    def __init__(self):
        self.host = 'localhost'
        self.user = 'root'
        self.password = 'L4512678.1'
        self.database = 'coast'
        self.connection = None
    
    def connect(self):
        """连接到数据库"""
        try:
            self.connection = pymysql.connect(
                host=self.host,
                user=self.user,
                password=self.password,
                database=self.database
            )
            print("数据库连接成功")
        except Error as e:
            print(f"数据库连接失败: {e}")
            # 尝试创建数据库
            try:
                connection = pymysql.connect(
                    host=self.host,
                    user=self.user,
                    password=self.password
                )
                cursor = connection.cursor()
                cursor.execute(f"CREATE DATABASE IF NOT EXISTS {self.database}")
                connection.close()
                
                # 再次尝试连接
                self.connection = pymysql.connect(
                    host=self.host,
                    user=self.user,
                    password=self.password,
                    database=self.database
                )
                print(f"数据库 {self.database} 创建成功并连接")
            except Error as e:
                print(f"创建数据库失败: {e}")
                self.connection = None
    
    def disconnect(self):
        """断开数据库连接"""
        if self.connection:
            self.connection.close()
            self.connection = None
            print("数据库连接已断开")
    
    def create_table(self, table_name, columns):
        """创建表"""
        if not self.connection:
            self.connect()
        
        try:
            cursor = self.connection.cursor()
            
            # 构建创建表的SQL语句
            create_table_sql = f"CREATE TABLE IF NOT EXISTS {table_name} ("
            create_table_sql += "id INT AUTO_INCREMENT PRIMARY KEY, "
            
            # 添加列
            for col in columns:
                # 替换列名中的空格为下划线
                safe_col = col.replace(' ', '_')
                create_table_sql += f"`{safe_col}` TEXT, "
            
            # 移除最后一个逗号和空格
            create_table_sql = create_table_sql[:-2]
            create_table_sql += ")"
            
            cursor.execute(create_table_sql)
            self.connection.commit()
            print(f"表 {table_name} 创建成功")
            return True
        except Error as e:
            print(f"创建表失败: {e}")
            return False
    
    def insert_data(self, table_name, data):
        """插入数据"""
        if not self.connection:
            self.connect()
        
        try:
            cursor = self.connection.cursor()
            
            # 构建插入语句
            columns = ', '.join([f"`{col.replace(' ', '_')}`" for col in data[0].keys()])
            placeholders = ', '.join(['%s'] * len(data[0]))
            insert_sql = f"INSERT INTO {table_name} ({columns}) VALUES ({placeholders})"
            
            # 批量插入
            for row in data:
                values = tuple(row.values())
                cursor.execute(insert_sql, values)
            
            self.connection.commit()
            print(f"数据插入到表 {table_name} 成功")
            return True
        except Error as e:
            print(f"插入数据失败: {e}")
            self.connection.rollback()
            return False
    
    def execute_query(self, query, params=None):
        """执行查询"""
        if not self.connection:
            self.connect()
        
        try:
            cursor = self.connection.cursor()
            cursor.execute(query, params)
            result = cursor.fetchall()
            return result
        except Error as e:
            print(f"执行查询失败: {e}")
            return None
    
    def execute_update(self, query, params=None):
        """执行更新操作"""
        if not self.connection:
            self.connect()
        
        try:
            cursor = self.connection.cursor()
            cursor.execute(query, params)
            self.connection.commit()
            return cursor.rowcount
        except Error as e:
            print(f"执行更新失败: {e}")
            self.connection.rollback()
            return 0
    
    def table_exists(self, table_name):
        """检查表是否存在"""
        if not self.connection:
            self.connect()
        
        try:
            cursor = self.connection.cursor()
            cursor.execute(f"SHOW TABLES LIKE '{table_name}'")
            result = cursor.fetchone()
            return result is not None
        except Error as e:
            print(f"检查表是否存在失败: {e}")
            return False    