#!/usr/bin/env python3
"""
共享单车租赁预测主程序
使用机器学习模型预测共享单车租赁数量
"""

import sys
import os

# 添加src目录到路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from src.data_loader import DataLoader
from src.data_preprocessor import DataPreprocessor
from src.model_trainer import ModelTrainer
from src.visualizer import Visualizer


def main():
    """主函数"""
    print("="*60)
    print("共享单车租赁预测系统")
    print("="*60)
    
    # 1. 加载数据
    print("\n[步骤 1] 加载数据...")
    data_loader = DataLoader(data_dir="data")
    
    # 选择使用每日数据或每小时数据
    use_hourly = False  # 设置为True使用每小时数据，False使用每日数据
    
    if use_hourly:
        df = data_loader.load_hour_data()
    else:
        df = data_loader.load_day_data()
    
    # 显示数据信息
    data_loader.get_data_info(df)
    
    # 2. 数据预处理
    print("\n[步骤 2] 数据预处理...")
    preprocessor = DataPreprocessor()
    X, y = preprocessor.prepare_features(df, target="cnt")
    
    # 特征标准化
    X_scaled, _ = preprocessor.scale_features(X)
    
    # 3. 模型训练
    print("\n[步骤 3] 模型训练...")
    trainer = ModelTrainer(random_state=42)
    results = trainer.train_models(X_scaled, y, test_size=0.2)
    
    # 4. 可视化结果
    print("\n[步骤 4] 生成可视化结果...")
    visualizer = Visualizer(output_dir="output")
    
    # 绘制每个模型的预测结果
    for model_name, result in results.items():
        visualizer.plot_predictions(
            result['y_test'],
            result['y_test_pred'],
            model_name
        )
        visualizer.plot_residuals(
            result['y_test'],
            result['y_test_pred'],
            model_name
        )
    
    # 绘制模型对比
    visualizer.plot_model_comparison(results)
    
    # 绘制特征重要性
    if trainer.best_model is not None:
        feature_importance = trainer.get_feature_importance()
        importance_df = preprocessor.get_feature_importance_data(feature_importance)
        visualizer.plot_feature_importance(importance_df, top_n=10)
    
    # 5. 总结
    print("\n" + "="*60)
    print("训练完成！")
    print("="*60)
    print(f"\n最佳模型: {trainer.best_model_name}")
    best_result = results[trainer.best_model_name]
    print(f"测试集 R² 分数: {best_result['test_metrics']['r2_score']:.4f}")
    print(f"测试集 RMSE: {best_result['test_metrics']['rmse']:.2f}")
    print(f"测试集 MAE: {best_result['test_metrics']['mae']:.2f}")
    print(f"\n所有可视化结果已保存至 output/ 目录")
    print("="*60)


if __name__ == "__main__":
    main()

