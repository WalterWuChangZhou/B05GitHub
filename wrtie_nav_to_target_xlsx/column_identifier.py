"""
智能列识别器 - 自动识别Excel/CSV中的列位置
"""
from difflib import get_close_matches


class ColumnIdentifier:
    """智能列识别器类"""

    def __init__(self, config):
        """
        初始化列识别器

        Args:
            config: ConfigManager实例
        """
        self.config = config
        keywords = config.get_column_keywords()
        self.code_keywords = keywords['code_keywords']
        self.date_keywords = keywords['date_keywords']
        self.nav_keywords = keywords['nav_keywords']

    def identify_columns(self, df, sheet_name=""):
        """
        智能识别DataFrame中的列

        Args:
            df: pandas DataFrame
            sheet_name: 工作表名称（用于显示）

        Returns:
            dict: {'code_col': 列名, 'date_col': 列名, 'nav_col': 列名}
        """
        columns = df.columns.tolist()

        if sheet_name:
            print(f"\n🔍 智能识别【{sheet_name}】列位置...")
        else:
            print(f"\n🔍 智能识别列位置...")
        print(f"   表头: {', '.join(columns)}")

        result = {}

        # 1. 识别代码列
        code_col = self._find_column(columns, self.code_keywords, '代码')
        if code_col:
            result['code_col'] = code_col
            print(f"   ✅ 代码列: [{code_col}] (第{columns.index(code_col) + 1}列)")
        else:
            # 尝试模糊匹配
            for col in columns:
                if '代码' in col or 'code' in col.lower() or '编号' in col:
                    result['code_col'] = col
                    print(f"   ⚠️ 代码列(模糊匹配): [{col}] (第{columns.index(col) + 1}列)")
                    break

        # 2. 识别日期列
        date_col = self._find_column(columns, self.date_keywords, '日期')
        if date_col:
            result['date_col'] = date_col
            print(f"   ✅ 日期列: [{date_col}] (第{columns.index(date_col) + 1}列)")
        else:
            for col in columns:
                if '日期' in col or 'date' in col.lower() or '时间' in col:
                    result['date_col'] = col
                    print(f"   ⚠️ 日期列(模糊匹配): [{col}] (第{columns.index(col) + 1}列)")
                    break

        # 3. 识别NAV列
        nav_col = self._find_column(columns, self.nav_keywords, '净值')
        if nav_col:
            result['nav_col'] = nav_col
            print(f"   ✅ NAV列: [{nav_col}] (第{columns.index(nav_col) + 1}列)")
        else:
            for col in columns:
                if '净值' in col or 'nav' in col.lower() or '价格' in col:
                    result['nav_col'] = col
                    print(f"   ⚠️ NAV列(模糊匹配): [{col}] (第{columns.index(col) + 1}列)")
                    break

        # 检查是否所有列都找到了
        missing = []
        if 'code_col' not in result:
            missing.append('代码')
        if 'date_col' not in result:
            missing.append('日期')
        if 'nav_col' not in result:
            missing.append('NAV')

        if missing:
            print(f"   ⚠️ 警告: 未找到以下列: {', '.join(missing)}")
            # 使用默认位置
            if 'code_col' not in result and len(columns) > 0:
                result['code_col'] = columns[0]
                print(f"   🔄 使用第一列作为代码列: {columns[0]}")
            if 'date_col' not in result and len(columns) > 1:
                result['date_col'] = columns[1]
                print(f"   🔄 使用第二列作为日期列: {columns[1]}")
            if 'nav_col' not in result and len(columns) > 2:
                result['nav_col'] = columns[2]
                print(f"   🔄 使用第三列作为NAV列: {columns[2]}")

        return result

    def _find_column(self, columns, keywords, label):
        """
        在列名中查找匹配的关键词

        Args:
            columns: 列名列表
            keywords: 关键词列表
            label: 标签（用于显示）

        Returns:
            str: 匹配的列名，未找到返回None
        """
        # 1. 精确匹配
        for keyword in keywords:
            for col in columns:
                if col.strip() == keyword.strip():
                    return col

        # 2. 包含匹配（大小写不敏感）
        for keyword in keywords:
            keyword_lower = keyword.lower()
            for col in columns:
                col_lower = col.lower()
                if keyword_lower in col_lower:
                    return col

        # 3. 模糊匹配（使用difflib）
        for keyword in keywords:
            matches = get_close_matches(keyword, columns, n=1, cutoff=0.6)
            if matches:
                return matches[0]

        return None