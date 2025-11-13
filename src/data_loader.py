"""
数据加载模块
负责从CSV文件加载共享单车数据并进行基本的数据检查
"""

import pandas as pd
import os
from typing import Tuple, Optional


class DataLoader:
    """数据加载器类"""
    
    def __init__(self, data_dir: str = "data"):
        """
        初始化数据加载器
        
        Args:
            data_dir: 数据文件所在目录
        """
        self.data_dir = data_dir
    
    def load_day_data(self) -> pd.DataFrame:
        """
        加载按天聚合的数据
        
        Returns:
            包含每日数据的DataFrame
            
        Raises:
            FileNotFoundError: 如果数据文件不存在
        """
        file_path = os.path.join(self.data_dir, "day.csv")
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"数据文件不存在: {file_path}")
        
        df = pd.read_csv(file_path)
        print(f"成功加载每日数据: {len(df)} 条记录")
        return df
    
    def load_hour_data(self) -> pd.DataFrame:
        """
        加载按小时聚合的数据
        
        Returns:
            包含每小时数据的DataFrame
            
        Raises:
            FileNotFoundError: 如果数据文件不存在
        """
        file_path = os.path.join(self.data_dir, "hour.csv")
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"数据文件不存在: {file_path}")
        
        df = pd.read_csv(file_path)
        print(f"成功加载每小时数据: {len(df)} 条记录")
        return df
    
    def get_data_info(self, df: pd.DataFrame) -> None:
        """
        打印数据基本信息
        
        Args:
            df: 要检查的DataFrame
        """
        print("\n" + "="*50)
        print("数据基本信息")
        print("="*50)
        print(f"数据形状: {df.shape}")
        print(f"\n列名: {list(df.columns)}")
        print(f"\n前5行数据:")
        print(df.head())
        print(f"\n数据类型:")
        print(df.dtypes)
        print(f"\n缺失值统计:")
        print(df.isnull().sum())
        print(f"\n数据统计摘要:")
        print(df.describe())
        print("="*50 + "\n")

