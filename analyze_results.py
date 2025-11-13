#!/usr/bin/env python3
"""
结果分析脚本
详细分析模型训练结果并提供深入见解
"""

import sys
import os
import pandas as pd
import numpy as np

# 添加src目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from src.data_loader import DataLoader
from src.data_preprocessor import DataPreprocessor
from src.model_trainer import ModelTrainer
from src.visualizer import Visualizer


def analyze_model_performance(results):
    """分析模型性能"""
    print("\n" + "="*70)
    print("模型性能详细分析")
    print("="*70)
    
    for model_name, result in results.items():
        print(f"\n【{model_name}】")
        print("-" * 70)
        
        train_metrics = result['train_metrics']
        test_metrics = result['test_metrics']
        
        # 计算过拟合程度
        train_r2 = train_metrics['r2_score']
        test_r2 = test_metrics['r2_score']
        overfitting = train_r2 - test_r2
        
        print(f"训练集性能:")
        print(f"  R² 分数: {train_r2:.4f}")
        print(f"  RMSE: {train_metrics['rmse']:.2f}")
        print(f"  MAE: {train_metrics['mae']:.2f}")
        
        print(f"\n测试集性能:")
        print(f"  R² 分数: {test_r2:.4f}")
        print(f"  RMSE: {test_metrics['rmse']:.2f}")
        print(f"  MAE: {test_metrics['mae']:.2f}")
        
        print(f"\n过拟合分析:")
        if overfitting < 0.05:
            print(f"  过拟合程度: {overfitting:.4f} (轻微，模型泛化能力良好)")
        elif overfitting < 0.15:
            print(f"  过拟合程度: {overfitting:.4f} (中等，模型可能略有过拟合)")
        else:
            print(f"  过拟合程度: {overfitting:.4f} (严重，模型存在明显过拟合)")
        
        # 预测误差分析
        y_test = result['y_test']
        y_pred = result['y_test_pred']
        errors = np.abs(y_test - y_pred)
        relative_errors = errors / (y_test + 1) * 100  # 避免除零
        
        print(f"\n预测误差分析:")
        print(f"  平均绝对误差: {errors.mean():.2f}")
        print(f"  误差中位数: {np.median(errors):.2f}")
        print(f"  最大误差: {errors.max():.2f}")
        print(f"  平均相对误差: {relative_errors.mean():.2f}%")
        print(f"  误差标准差: {errors.std():.2f}")


def analyze_feature_importance(trainer, preprocessor):
    """分析特征重要性"""
    print("\n" + "="*70)
    print("特征重要性分析")
    print("="*70)
    
    if trainer.best_model is None:
        print("模型尚未训练")
        return
    
    feature_importance = trainer.get_feature_importance()
    importance_df = preprocessor.get_feature_importance_data(feature_importance)
    
    print(f"\n最佳模型: {trainer.best_model_name}")
    print("\n特征重要性排序 (从高到低):")
    print("-" * 70)
    
    total_importance = importance_df['importance'].sum()
    
    for idx, row in importance_df.iterrows():
        importance_pct = (row['importance'] / total_importance) * 100
        bar_length = int(importance_pct / 2)  # 每2%一个字符
        bar = "█" * bar_length
        print(f"{row['feature']:15s} | {bar:50s} | {row['importance']:.4f} ({importance_pct:.1f}%)")
    
    # 分析最重要的特征
    top_features = importance_df.head(5)
    print(f"\n最重要的5个特征:")
    for idx, row in top_features.iterrows():
        print(f"  {idx+1}. {row['feature']}: {row['importance']:.4f}")


def analyze_prediction_quality(results):
    """分析预测质量"""
    print("\n" + "="*70)
    print("预测质量分析")
    print("="*70)
    
    for model_name, result in results.items():
        print(f"\n【{model_name}】")
        print("-" * 70)
        
        y_test = result['y_test']
        y_pred = result['y_test_pred']
        
        # 按真实值范围分析误差
        ranges = [
            (0, 2000, "低需求 (0-2000)"),
            (2000, 4000, "中低需求 (2000-4000)"),
            (4000, 6000, "中高需求 (4000-6000)"),
            (6000, float('inf'), "高需求 (6000+)")
        ]
        
        print("不同需求水平的预测误差:")
        for min_val, max_val, label in ranges:
            mask = (y_test >= min_val) & (y_test < max_val)
            if mask.sum() > 0:
                errors = np.abs(y_test[mask] - y_pred[mask])
                mae = errors.mean()
                mape = (errors / (y_test[mask] + 1) * 100).mean()
                print(f"  {label:25s}: MAE={mae:.2f}, MAPE={mape:.2f}%, 样本数={mask.sum()}")


def compare_models(results):
    """对比模型"""
    print("\n" + "="*70)
    print("模型对比总结")
    print("="*70)
    
    comparison_data = []
    for model_name, result in results.items():
        test_metrics = result['test_metrics']
        train_metrics = result['train_metrics']
        comparison_data.append({
            '模型': model_name,
            '测试集R²': test_metrics['r2_score'],
            '测试集RMSE': test_metrics['rmse'],
            '测试集MAE': test_metrics['mae'],
            '训练集R²': train_metrics['r2_score'],
            '过拟合程度': train_metrics['r2_score'] - test_metrics['r2_score']
        })
    
    df = pd.DataFrame(comparison_data)
    print("\n模型性能对比表:")
    print(df.to_string(index=False))
    
    # 找出最佳模型
    best_r2 = df.loc[df['测试集R²'].idxmax()]
    best_rmse = df.loc[df['测试集RMSE'].idxmin()]
    best_mae = df.loc[df['测试集MAE'].idxmin()]
    
    print(f"\n最佳R²分数: {best_r2['模型']} ({best_r2['测试集R²']:.4f})")
    print(f"最低RMSE: {best_rmse['模型']} ({best_rmse['测试集RMSE']:.2f})")
    print(f"最低MAE: {best_mae['模型']} ({best_mae['测试集MAE']:.2f})")


def generate_recommendations(results, trainer):
    """生成改进建议"""
    print("\n" + "="*70)
    print("改进建议")
    print("="*70)
    
    recommendations = []
    
    # 检查过拟合
    for model_name, result in results.items():
        train_r2 = result['train_metrics']['r2_score']
        test_r2 = result['test_metrics']['r2_score']
        overfitting = train_r2 - test_r2
        
        if overfitting > 0.15:
            recommendations.append(f"⚠️  {model_name}存在明显过拟合(差异{overfitting:.3f})，建议:")
            recommendations.append("   - 增加正则化参数")
            recommendations.append("   - 减少模型复杂度")
            recommendations.append("   - 增加训练数据或使用交叉验证")
    
    # 检查模型性能
    best_result = results[trainer.best_model_name]
    best_r2 = best_result['test_metrics']['r2_score']
    
    if best_r2 < 0.7:
        recommendations.append("⚠️  模型R²分数较低，建议:")
        recommendations.append("   - 进行更深入的特征工程")
        recommendations.append("   - 尝试更复杂的模型(如XGBoost、神经网络)")
        recommendations.append("   - 检查数据质量和特征选择")
    elif best_r2 < 0.85:
        recommendations.append("💡 模型性能良好，但仍有改进空间:")
        recommendations.append("   - 尝试特征交互项")
        recommendations.append("   - 进行超参数调优")
        recommendations.append("   - 尝试集成学习方法")
    else:
        recommendations.append("✅ 模型性能优秀!")
        recommendations.append("   - 可以考虑模型部署")
        recommendations.append("   - 可以尝试模型压缩以提升推理速度")
    
    # 特征工程建议
    if trainer.best_model is not None:
        feature_importance = trainer.get_feature_importance()
        if len(feature_importance) > 0:
            max_importance = feature_importance.max()
            min_importance = feature_importance.min()
            if max_importance / min_importance > 100:
                recommendations.append("💡 特征重要性差异较大，建议:")
                recommendations.append("   - 考虑移除重要性极低的特征")
                recommendations.append("   - 对重要特征进行更精细的特征工程")
    
    if recommendations:
        for rec in recommendations:
            print(rec)
    else:
        print("模型表现良好，暂无特殊建议。")


def main():
    """主分析函数"""
    print("="*70)
    print("共享单车租赁预测 - 结果分析报告")
    print("="*70)
    
    # 1. 加载数据
    print("\n[步骤 1] 加载数据...")
    data_loader = DataLoader(data_dir="data")
    # df = data_loader.load_day_data()
    df = data_loader.load_hour_data()
    
    # 2. 数据预处理
    print("\n[步骤 2] 数据预处理...")
    preprocessor = DataPreprocessor()
    X, y = preprocessor.prepare_features(df, target="cnt")
    X_scaled, _ = preprocessor.scale_features(X)
    
    # 3. 模型训练
    print("\n[步骤 3] 模型训练...")
    trainer = ModelTrainer(random_state=42)
    results = trainer.train_models(X_scaled, y, test_size=0.2)
    
    # 4. 详细分析
    analyze_model_performance(results)
    analyze_feature_importance(trainer, preprocessor)
    analyze_prediction_quality(results)
    compare_models(results)
    generate_recommendations(results, trainer)
    
    print("\n" + "="*70)
    print("分析完成！")
    print("="*70)


if __name__ == "__main__":
    main()

