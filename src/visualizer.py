"""
可视化模块
负责绘制模型评估结果和特征重要性
"""

# 关键修改：设置非交互式后端（避免tkinter线程冲突）
import matplotlib
matplotlib.use('Agg')  # 必须在导入pyplot之前设置，使用Agg后端（无GUI，适合保存图片）

import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import pandas as pd
from typing import Dict
import os

# 关键修改：优化字体配置（确保支持上标等符号）
# 优先使用支持更多符号的中文字体，按优先级排序
plt.rcParams['font.sans-serif'] = [
    'Microsoft YaHei',  # 微软雅黑（支持上标²等符号，Windows系统常见）
    'Arial Unicode MS',  # 跨平台字体，支持多符号
    'SimHei',
    'DejaVu Sans'
]
plt.rcParams['axes.unicode_minus'] = False  # 解决负号显示问题


class Visualizer:
    """可视化类"""

    def __init__(self, output_dir: str = "output"):
        """
        初始化可视化器

        Args:
            output_dir: 输出目录
        """
        self.output_dir = output_dir
        os.makedirs(output_dir, exist_ok=True)

    def plot_predictions(self, y_true: np.ndarray, y_pred: np.ndarray,
                         model_name: str, save_path: str = None) -> None:
        """
        绘制预测值 vs 真实值散点图
        """
        plt.figure(figsize=(10, 6))

        plt.scatter(y_true, y_pred, alpha=0.5, s=20)

        min_val = min(y_true.min(), y_pred.min())
        max_val = max(y_true.max(), y_pred.max())
        plt.plot([min_val, max_val], [min_val, max_val], 'r--', lw=2, label='完美预测线')

        plt.xlabel('真实值', fontsize=12)
        plt.ylabel('预测值', fontsize=12)
        plt.title(f'{model_name} - 预测值 vs 真实值', fontsize=14, fontweight='bold')
        plt.legend()
        plt.grid(True, alpha=0.3)

        if save_path is None:
            save_path = os.path.join(self.output_dir, f'{model_name}_predictions.png')

        plt.tight_layout()
        plt.savefig(save_path, dpi=300, bbox_inches='tight')
        print(f"预测图已保存至: {save_path}")
        plt.close()  # 确保关闭图形，释放资源

    def plot_residuals(self, y_true: np.ndarray, y_pred: np.ndarray,
                       model_name: str, save_path: str = None) -> None:
        """
        绘制残差图
        """
        residuals = y_true - y_pred

        fig, axes = plt.subplots(1, 2, figsize=(14, 5))

        axes[0].scatter(y_pred, residuals, alpha=0.5, s=20)
        axes[0].axhline(y=0, color='r', linestyle='--', lw=2)
        axes[0].set_xlabel('预测值', fontsize=12)
        axes[0].set_ylabel('残差', fontsize=12)
        axes[0].set_title('残差 vs 预测值', fontsize=12, fontweight='bold')
        axes[0].grid(True, alpha=0.3)

        axes[1].hist(residuals, bins=30, edgecolor='black', alpha=0.7)
        axes[1].axvline(x=0, color='r', linestyle='--', lw=2)
        axes[1].set_xlabel('残差', fontsize=12)
        axes[1].set_ylabel('频数', fontsize=12)
        axes[1].set_title('残差分布', fontsize=12, fontweight='bold')
        axes[1].grid(True, alpha=0.3)

        plt.suptitle(f'{model_name} - 残差分析', fontsize=14, fontweight='bold')

        if save_path is None:
            save_path = os.path.join(self.output_dir, f'{model_name}_residuals.png')

        plt.tight_layout()
        plt.savefig(save_path, dpi=300, bbox_inches='tight')
        print(f"残差图已保存至: {save_path}")
        plt.close()  # 确保关闭图形，释放资源

    def plot_feature_importance(self, importance_df: pd.DataFrame,
                                top_n: int = 10, save_path: str = None) -> None:
        """
        绘制特征重要性图
        """
        top_features = importance_df.head(top_n)

        plt.figure(figsize=(10, 6))
        sns.barplot(data=top_features, x='importance', y='feature', hue='feature', palette='viridis', legend=False)
        plt.xlabel('重要性', fontsize=12)
        plt.ylabel('特征', fontsize=12)
        plt.title(f'特征重要性 (Top {top_n})', fontsize=14, fontweight='bold')
        plt.grid(True, alpha=0.3, axis='x')

        if save_path is None:
            save_path = os.path.join(self.output_dir, 'feature_importance.png')

        plt.tight_layout()
        plt.savefig(save_path, dpi=300, bbox_inches='tight')
        print(f"特征重要性图已保存至: {save_path}")
        plt.close()  # 确保关闭图形，释放资源

    def plot_model_comparison(self, results: Dict, save_path: str = None) -> None:
        """
        绘制模型对比图
        """
        model_names = list(results.keys())
        metrics = ['rmse', 'mae', 'r2_score']
        metric_labels = ['RMSE', 'MAE', 'R² Score']  # 这里的²会被正确显示

        fig, axes = plt.subplots(1, 3, figsize=(15, 5))

        for idx, (metric, label) in enumerate(zip(metrics, metric_labels)):
            test_values = [results[name]['test_metrics'][metric] for name in model_names]

            bars = axes[idx].bar(model_names, test_values, alpha=0.7,
                                 color=['#3498db', '#2ecc71'])
            axes[idx].set_ylabel(label, fontsize=12)
            axes[idx].set_title(f'测试集 {label}', fontsize=12, fontweight='bold')
            axes[idx].grid(True, alpha=0.3, axis='y')

            for bar, val in zip(bars, test_values):
                height = bar.get_height()
                axes[idx].text(bar.get_x() + bar.get_width()/2., height,
                               f'{val:.4f}', ha='center', va='bottom', fontsize=10)

        plt.suptitle('模型性能对比', fontsize=14, fontweight='bold')

        if save_path is None:
            save_path = os.path.join(self.output_dir, 'model_comparison.png')

        plt.tight_layout()
        plt.savefig(save_path, dpi=300, bbox_inches='tight')
        print(f"模型对比图已保存至: {save_path}")
        plt.close()  # 确保关闭图形，释放资源