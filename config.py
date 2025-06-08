# 数据库配置
DB_CONFIG = {
    'host': 'localhost',
    'user': 'root',
    'password': 'L4512678.1',
    'database': 'coast'
}

# 匹配文件模板列名
MATCH_TEMPLATE_COLUMNS = [
    '序号', '名称', '规格', '型号', '工作内容', '项目特征', '修正项目'
]

# 定额导入模板列名
QUOTA_TEMPLATE_COLUMNS = [
    '定额编号', '分部分项工程名称', '计量单位', '工程量', 
    '主材费', '小计', '人工费', '材料费', '机械费'
]    