"""
配置管理器 - 负责读取和管理所有配置
"""
import configparser
import os
from datetime import datetime


class ConfigManager:
    """配置管理器类"""
    
    def __init__(self, config_file='config.ini'):
        """
        初始化配置管理器
        
        Args:
            config_file: 配置文件路径
        """
        self.config_file = config_file
        self.config = configparser.ConfigParser()
        self._load_config()
    
    def _load_config(self):
        """加载配置文件"""
        if not os.path.exists(self.config_file):
            print(f"⚠️ 配置文件 {self.config_file} 不存在，创建默认配置...")
            self._create_default_config()
        
        self.config.read(self.config_file, encoding='utf-8')
        print(f"✅ 配置文件加载成功: {self.config_file}")
    
    def _create_default_config(self):
        """创建默认配置文件"""
        self.config['Paths'] = {
            'source_folder': 'E:\\B01Python\\PythonProject\\fund_nav_data\\fund_nav_data',
            'target_file': 'E:\\20260606套利计算.xlsm',
            'backup_folder': './backup',
            'log_folder': './logs'
        }
        
        self.config['Sheets'] = {
            'sheet_names': '赎回,申购'
        }
        
        self.config['ColumnMapping'] = {
            'code_keywords': '代码,基金代码,基金编号,code,fund_code',
            'date_keywords': '日期,交易日期,净值日期,date,trade_date',
            'nav_keywords': '净值,单位净值,NAV,nav,unit_nav'
        }
        
        self.config['Matching'] = {
            'nav_tolerance': '0.00001',
            'date_match_mode': 'date_only'
        }
        
        self.config['Logging'] = {
            'log_level': 'INFO',
            'show_debug': 'true',
            'max_display_records': '20',
            'save_log': 'true'
        }
        
        self.config['Backup'] = {
            'auto_backup': 'true',
            'backup_count': '5'
        }
        
        # 确保目录存在
        os.makedirs('./backup', exist_ok=True)
        os.makedirs('./logs', exist_ok=True)
        
        with open(self.config_file, 'w', encoding='utf-8') as f:
            self.config.write(f)
        print(f"✅ 默认配置文件已创建: {self.config_file}")
    
    def get(self, section, key, fallback=None):
        """获取配置值"""
        try:
            return self.config.get(section, key, fallback=fallback)
        except:
            return fallback
    
    def get_int(self, section, key, fallback=0):
        """获取整数配置值"""
        try:
            return self.config.getint(section, key, fallback=fallback)
        except:
            return fallback
    
    def get_float(self, section, key, fallback=0.0):
        """获取浮点数配置值"""
        try:
            return self.config.getfloat(section, key, fallback=fallback)
        except:
            return fallback
    
    def get_boolean(self, section, key, fallback=False):
        """获取布尔配置值"""
        try:
            return self.config.getboolean(section, key, fallback=fallback)
        except:
            return fallback
    
    def get_list(self, section, key, fallback=''):
        """获取列表配置值"""
        value = self.get(section, key, fallback)
        if not value:
            return []
        return [item.strip() for item in value.split(',') if item.strip()]
    
    def get_path(self, key):
        """获取路径配置"""
        return self.get('Paths', key)
    
    def get_sheet_names(self):
        """获取工作表名称列表"""
        return self.get_list('Sheets', 'sheet_names')
    
    def get_column_keywords(self):
        """获取列关键词配置"""
        return {
            'code_keywords': self.get_list('ColumnMapping', 'code_keywords'),
            'date_keywords': self.get_list('ColumnMapping', 'date_keywords'),
            'nav_keywords': self.get_list('ColumnMapping', 'nav_keywords')
        }
    
    def get_matching_config(self):
        """获取匹配配置"""
        return {
            'nav_tolerance': self.get_float('Matching', 'nav_tolerance', 0.00001),
            'date_match_mode': self.get('Matching', 'date_match_mode', 'date_only')
        }
    
    def get_logging_config(self):
        """获取日志配置"""
        return {
            'log_level': self.get('Logging', 'log_level', 'INFO'),
            'show_debug': self.get_boolean('Logging', 'show_debug', True),
            'max_display_records': self.get_int('Logging', 'max_display_records', 20),
            'save_log': self.get_boolean('Logging', 'save_log', True)
        }
    
    def get_backup_config(self):
        """获取备份配置"""
        return {
            'auto_backup': self.get_boolean('Backup', 'auto_backup', True),
            'backup_count': self.get_int('Backup', 'backup_count', 5)
        }
    
    def update_path(self, key, value):
        """更新路径配置"""
        self.config.set('Paths', key, value)
        with open(self.config_file, 'w', encoding='utf-8') as f:
            self.config.write(f)
        print(f"✅ 路径配置已更新: {key} = {value}")