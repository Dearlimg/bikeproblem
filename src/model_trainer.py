"""
模型训练模块
负责训练和评估回归模型
"""

import numpy as np
import pandas as pd
from typing import Dict, Tuple
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestRegressor
from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_squared_error, mean_absolute_error, r2_score


class ModelTrainer:
    """模型训练器类"""
    
    def __init__(self, random_state: int = 42):
        """
        初始化模型训练器
        
        Args:
            random_state: 随机种子
        """
        self.random_state = random_state
        self.models = {}
        self.best_model = None
        self.best_model_name = None
    
    def train_models(self, X: pd.DataFrame, y: pd.Series, test_size: float = 0.2) -> Dict[str, Dict]:
        """
        训练多个模型并比较性能
        
        Args:
            X: 特征数据
            y: 目标变量
            test_size: 测试集比例
            
        Returns:
            包含各模型评估指标的字典
        """
        # 划分训练集和测试集
        X_train, X_test, y_train, y_test = train_test_split(
            X, y, test_size=test_size, random_state=self.random_state
        )
        
        print(f"\n数据划分: 训练集 {len(X_train)} 条, 测试集 {len(X_test)} 条\n")
        
        # 定义要训练的模型
        models_to_train = {
            'Linear Regression': LinearRegression(),
            'Random Forest': RandomForestRegressor(
                n_estimators=1000,
                max_depth=10,
                random_state=self.random_state,
                n_jobs=-1
            )
        }
        
        results = {}
        
        # 训练每个模型
        for name, model in models_to_train.items():
            print(f"正在训练 {name}...")
            
            # 训练模型
            model.fit(X_train, y_train)
            
            # 预测
            y_train_pred = model.predict(X_train)
            y_test_pred = model.predict(X_test)
            
            # 计算评估指标
            train_metrics = self._calculate_metrics(y_train, y_train_pred, "训练集")
            test_metrics = self._calculate_metrics(y_test, y_test_pred, "测试集")
            
            results[name] = {
                'model': model,
                'train_metrics': train_metrics,
                'test_metrics': test_metrics,
                'y_test': y_test,
                'y_test_pred': y_test_pred
            }
            
            self.models[name] = model
        
        # 选择最佳模型（基于测试集R²分数）
        best_name = max(results.keys(), key=lambda k: results[k]['test_metrics']['r2_score'])
        self.best_model = results[best_name]['model']
        self.best_model_name = best_name
        
        print(f"\n最佳模型: {self.best_model_name}")
        print(f"测试集 R² 分数: {results[best_name]['test_metrics']['r2_score']:.4f}")
        
        return results
    
    def _calculate_metrics(self, y_true: np.ndarray, y_pred: np.ndarray, dataset_name: str = "") -> Dict[str, float]:
        """
        计算回归评估指标
        
        Args:
            y_true: 真实值
            y_pred: 预测值
            dataset_name: 数据集名称（用于打印）
            
        Returns:
            包含各种评估指标的字典
        """
        mse = mean_squared_error(y_true, y_pred)
        rmse = np.sqrt(mse)
        mae = mean_absolute_error(y_true, y_pred)
        r2 = r2_score(y_true, y_pred)
        
        metrics = {
            'mse': mse,
            'rmse': rmse,
            'mae': mae,
            'r2_score': r2
        }
        
        if dataset_name:
            print(f"\n{dataset_name}评估指标:")
            print(f"  MSE (均方误差): {mse:.2f}")
            print(f"  RMSE (均方根误差): {rmse:.2f}")
            print(f"  MAE (平均绝对误差): {mae:.2f}")
            print(f"  R² 分数: {r2:.4f}")
        
        return metrics
    
    def get_feature_importance(self) -> np.ndarray:
        """
        获取最佳模型的特征重要性
        
        Returns:
            特征重要性数组
            
        Raises:
            ValueError: 如果模型不是RandomForest或未训练
        """
        if self.best_model is None:
            raise ValueError("模型尚未训练")
        
        if hasattr(self.best_model, 'feature_importances_'):
            return self.best_model.feature_importances_
        elif hasattr(self.best_model, 'coef_'):
            # 对于线性回归，使用系数的绝对值作为重要性
            return np.abs(self.best_model.coef_)
        else:
            raise ValueError("模型不支持特征重要性提取")
    
    def predict(self, X: pd.DataFrame) -> np.ndarray:
        """
        使用最佳模型进行预测
        
        Args:
            X: 特征数据
            
        Returns:
            预测值数组
        """
        if self.best_model is None:
            raise ValueError("模型尚未训练")
        
        return self.best_model.predict(X)

