from PyQt5.QtCore import QThread, pyqtSignal
import pymysql
import hashlib
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

class MatchWorker(QThread):
    progress_updated = pyqtSignal(int)
    match_completed = pyqtSignal(bool, str)
    match_result = pyqtSignal(list)
    
    def __init__(self, project, match_data):
        super().__init__()
        self.project = project
        self.match_data = match_data
    
    def run(self):
        try:
            conn = pymysql.connect(
                host='localhost',
                user='root',
                password='L4512678.1',
                database='coast'
            )
            cursor = conn.cursor()
            
            total_rows = len(self.match_data)
            result = []
            
            # 获取定额表结构
            cursor.execute(f"SHOW COLUMNS FROM {self.project['quota_table']}")
            quota_columns = [col[0] for col in cursor.fetchall()]
            
            # 对于每一行匹配数据
            for i, row in enumerate(self.match_data):
                # 更新进度
                if i % 10 == 0:  # 每10行更新一次进度
                    progress = int((i + 1) / total_rows * 100)
                    self.progress_updated.emit(progress)
                
                # 复制当前行数据
                result_row = row.copy()
                
                # 如果修正项目已有编码，直接使用
                if row.get('修正项目'):
                    code = row['修正项目']
                    
                    # 从定额表中查询该编码的详细信息
                    query = f"SELECT * FROM {self.project['quota_table']} WHERE 定额编号 = %s"
                    cursor.execute(query, (code,))
                    quota_data = cursor.fetchone()
                    
                    if quota_data:
                        # 将定额数据添加到结果中
                        for col_idx, col_name in enumerate(quota_columns):
                            if col_name != 'id':  # 跳过id列
                                result_row[f'定额_{col_name}'] = quota_data[col_idx]
                else:
                    # 计算匹配条件的哈希值
                    condition = f"{row.get('名称', '')}{row.get('规格', '')}{row.get('型号', '')}{row.get('工作内容', '')}{row.get('项目特征', '')}"
                    condition_hash = hashlib.md5(condition.encode('utf-8')).hexdigest()
                    
                    # 先从绑定表中查找
                    bind_query = f"SELECT * FROM {self.project['bind_table']} WHERE project_id = %s AND condition_hash = %s ORDER BY use_count DESC LIMIT 1"
                    cursor.execute(bind_query, (self.project['id'], condition_hash))
                    bind_data = cursor.fetchone()
                    
                    if bind_data:
                        # 使用绑定表中的编码
                        code = bind_data[quota_columns.index('定额编号')]  # 假设定额编号在第2列
                        result_row['修正项目'] = code
                        
                        # 更新使用次数
                        update_query = f"UPDATE {self.project['bind_table']} SET use_count = use_count + 1 WHERE id = %s"
                        cursor.execute(update_query, (bind_data[0],))
                        conn.commit()
                        
                        # 将定额数据添加到结果中
                        for col_idx, col_name in enumerate(quota_columns):
                            if col_name != 'id':  # 跳过id列
                                result_row[f'定额_{col_name}'] = bind_data[col_idx + 3]  # 绑定表中多了project_id, condition_hash, use_count三列
                    else:
                        # 没有绑定数据，使用文本匹配
                        # 拼接匹配条件
                        match_text = f"{row.get('名称', '')} {row.get('规格', '')} {row.get('型号', '')} {row.get('工作内容', '')} {row.get('项目特征', '')}"
                        
                        # 获取所有定额数据
                        cursor.execute(f"SELECT * FROM {self.project['quota_table']}")
                        quota_rows = cursor.fetchall()
                        
                        # 计算相似度
                        best_similarity = -1
                        best_code = None
                        best_quota_data = None
                        
                        for quota_row in quota_rows:
                            # 拼接定额文本
                            quota_text = f"{quota_row[1]} {quota_row[2]} {quota_row[3]} {quota_row[4]} {quota_row[5]}"  # 假设前5列是名称、规格、型号、工作内容、项目特征
                            
                            # 计算相似度
                            similarity = self.calculate_similarity(match_text, quota_text)
                            
                            if similarity > best_similarity:
                                best_similarity = similarity
                                best_code = quota_row[1]  # 假设定额编号在第2列
                                best_quota_data = quota_row
                        
                        if best_code:
                            result_row['修正项目'] = best_code
                            
                            # 将定额数据添加到结果中
                            for col_idx, col_name in enumerate(quota_columns):
                                if col_name != 'id':  # 跳过id列
                                    result_row[f'定额_{col_name}'] = best_quota_data[col_idx]
                
                result.append(result_row)
            
            conn.close()
            self.match_result.emit(result)
            self.match_completed.emit(True, '匹配完成')
        except Exception as e:
            self.match_completed.emit(False, str(e))
    
    def calculate_similarity(self, text1, text2):
        """计算两个文本的相似度"""
        if not text1 or not text2:
            return 0
        
        # 使用简单的余弦相似度
        vectorizer = TfidfVectorizer()
        tfidf_matrix = vectorizer.fit_transform([text1, text2])
        similarity = cosine_similarity(tfidf_matrix[0:1], tfidf_matrix[1:2])[0][0]
        
        return similarity    