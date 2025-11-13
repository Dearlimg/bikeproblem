"""
数据预处理模块
负责特征工程和数据清洗
"""

import pandas as pd
import numpy as np
from typing import Tuple, Optional
from sklearn.preprocessing import StandardScaler


class DataPreprocessor:
    """数据预处理器类"""
    
    def __init__(self):
        """初始化预处理器"""
        self.scaler = StandardScaler()
        self.feature_columns = None
    
    def prepare_features(self, df: pd.DataFrame, target: str = "cnt") -> Tuple[pd.DataFrame, pd.Series]:
        """
        准备特征和目标变量
        
        Args:
            df: 原始数据DataFrame
            target: 目标变量列名
            
        Returns:
            (特征DataFrame, 目标变量Series)
        """
        # 复制数据避免修改原始数据
        data = df.copy()
        
        # 删除不需要的列
        columns_to_drop = ['instant', 'dteday', 'casual', 'registered']
        if target in columns_to_drop:
            columns_to_drop.remove(target)
        
        # 确保目标列存在
        if target not in data.columns:
            raise ValueError(f"目标列 '{target}' 不存在于数据中")
        
        # 分离特征和目标
        X = data.drop(columns=[col for col in columns_to_drop if col in data.columns])
        X = X.drop(columns=[target])
        y = data[target]
        
        # 保存特征列名
        self.feature_columns = X.columns.tolist()
        
        print(f"特征数量: {len(self.feature_columns)}")
        print(f"特征列表: {self.feature_columns}")
        print(f"目标变量统计: 均值={y.mean():.2f}, 标准差={y.std():.2f}, 最小值={y.min()}, 最大值={y.max()}")
        
        return X, y
    
    def scale_features(self, X_train: pd.DataFrame, X_test: Optional[pd.DataFrame] = None) -> Tuple[pd.DataFrame, Optional[pd.DataFrame]]:
        """
        标准化特征
        
        Args:
            X_train: 训练集特征
            X_test: 测试集特征（可选）
            
        Returns:
            (标准化后的训练集, 标准化后的测试集)
        """
        X_train_scaled = pd.DataFrame(
            self.scaler.fit_transform(X_train),
            columns=X_train.columns,
            index=X_train.index
        )
        
        if X_test is not None:
            X_test_scaled = pd.DataFrame(
                self.scaler.transform(X_test),
                columns=X_test.columns,
                index=X_test.index
            )
            return X_train_scaled, X_test_scaled
        
        return X_train_scaled, None
    
    def get_feature_importance_data(self, feature_importance: np.ndarray) -> pd.DataFrame:
        """
        获取特征重要性数据框（用于可视化）
        
        Args:
            feature_importance: 特征重要性数组
            
        Returns:
            包含特征名称和重要性的DataFrame
        """
        if self.feature_columns is None:
            raise ValueError("特征列未定义，请先调用 prepare_features")
        
        importance_df = pd.DataFrame({
            'feature': self.feature_columns,
            'importance': feature_importance
        }).sort_values('importance', ascending=False)
        
        return importance_df

