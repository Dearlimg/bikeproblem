#!/usr/bin/env python3
"""
创建项目演示PPT
风格：简洁大气，蓝色主题
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

# 蓝色主题色
BLUE_PRIMARY = RGBColor(0, 51, 102)      # 深蓝色
BLUE_SECONDARY = RGBColor(0, 102, 204)    # 中蓝色
BLUE_LIGHT = RGBColor(173, 216, 230)     # 浅蓝色
BLUE_ACCENT = RGBColor(0, 153, 255)      # 强调蓝
WHITE = RGBColor(255, 255, 255)
GRAY = RGBColor(128, 128, 128)

def create_title_slide(prs):
    """创建标题页"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # 空白布局
    
    # 背景色
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = BLUE_PRIMARY
    
    # 标题
    title_box = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(1.5))
    title_frame = title_box.text_frame
    title_frame.text = "共享单车租赁预测系统"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(54)
    title_para.font.bold = True
    title_para.font.color.rgb = WHITE
    title_para.alignment = PP_ALIGN.CENTER
    
    # 副标题
    subtitle_box = slide.shapes.add_textbox(Inches(1), Inches(4), Inches(8), Inches(1))
    subtitle_frame = subtitle_box.text_frame
    subtitle_frame.text = "基于机器学习的需求预测分析"
    subtitle_para = subtitle_frame.paragraphs[0]
    subtitle_para.font.size = Pt(28)
    subtitle_para.font.color.rgb = BLUE_LIGHT
    subtitle_para.alignment = PP_ALIGN.CENTER

def create_content_slide(prs, title, content_items, layout_idx=1):
    """创建内容页"""
    slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
    
    # 设置背景
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = WHITE
    
    # 标题
    title_shape = slide.shapes.title
    title_shape.text = title
    title_para = title_shape.text_frame.paragraphs[0]
    title_para.font.size = Pt(36)
    title_para.font.bold = True
    title_para.font.color.rgb = BLUE_PRIMARY
    
    # 内容
    if layout_idx == 1:  # 标题和内容布局
        content_shape = slide.placeholders[1]
        tf = content_shape.text_frame
        tf.text = content_items[0]
        
        for item in content_items[1:]:
            p = tf.add_paragraph()
            p.text = item
            p.level = 0
            p.font.size = Pt(18)
            p.font.color.rgb = GRAY
            p.space_after = Pt(12)
    
    return slide

def create_two_column_slide(prs, title, left_items, right_items):
    """创建两列内容页"""
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    
    # 背景
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = WHITE
    
    # 标题
    title_shape = slide.shapes.title
    title_shape.text = title
    title_para = title_shape.text_frame.paragraphs[0]
    title_para.font.size = Pt(36)
    title_para.font.bold = True
    title_para.font.color.rgb = BLUE_PRIMARY
    
    # 左列
    left_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.8), Inches(4.5), Inches(5))
    left_frame = left_box.text_frame
    left_frame.text = left_items[0]
    left_para = left_frame.paragraphs[0]
    left_para.font.size = Pt(18)
    left_para.font.color.rgb = BLUE_SECONDARY
    left_para.font.bold = True
    
    for item in left_items[1:]:
        p = left_frame.add_paragraph()
        p.text = item
        p.font.size = Pt(16)
        p.font.color.rgb = GRAY
        p.space_after = Pt(8)
    
    # 右列
    right_box = slide.shapes.add_textbox(Inches(5.5), Inches(1.8), Inches(4.5), Inches(5))
    right_frame = right_box.text_frame
    right_frame.text = right_items[0]
    right_para = right_frame.paragraphs[0]
    right_para.font.size = Pt(18)
    right_para.font.color.rgb = BLUE_SECONDARY
    right_para.font.bold = True
    
    for item in right_items[1:]:
        p = right_frame.add_paragraph()
        p.text = item
        p.font.size = Pt(16)
        p.font.color.rgb = GRAY
        p.space_after = Pt(8)

def create_metrics_slide(prs, title, metrics):
    """创建指标展示页"""
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    
    # 背景
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = WHITE
    
    # 标题
    title_shape = slide.shapes.title
    title_shape.text = title
    title_para = title_shape.text_frame.paragraphs[0]
    title_para.font.size = Pt(36)
    title_para.font.bold = True
    title_para.font.color.rgb = BLUE_PRIMARY
    
    # 创建指标框
    y_start = 1.8
    box_height = 1.2
    box_width = 2.8
    spacing = 0.2
    
    for i, (label, value, desc) in enumerate(metrics):
        x = 0.5 + (i % 3) * (box_width + spacing)
        y = y_start + (i // 3) * (box_height + spacing)
        
        # 指标框
        box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(box_width), Inches(box_height))
        frame = box.text_frame
        frame.text = value
        para = frame.paragraphs[0]
        para.font.size = Pt(32)
        para.font.bold = True
        para.font.color.rgb = BLUE_ACCENT
        para.alignment = PP_ALIGN.CENTER
        
        # 标签
        label_box = slide.shapes.add_textbox(Inches(x), Inches(y + 0.5), Inches(box_width), Inches(0.3))
        label_frame = label_box.text_frame
        label_frame.text = label
        label_para = label_frame.paragraphs[0]
        label_para.font.size = Pt(14)
        label_para.font.color.rgb = GRAY
        label_para.alignment = PP_ALIGN.CENTER
        
        # 描述
        desc_box = slide.shapes.add_textbox(Inches(x), Inches(y + 0.8), Inches(box_width), Inches(0.3))
        desc_frame = desc_box.text_frame
        desc_frame.text = desc
        desc_para = desc_frame.paragraphs[0]
        desc_para.font.size = Pt(10)
        desc_para.font.color.rgb = GRAY
        desc_para.alignment = PP_ALIGN.CENTER

def main():
    """创建PPT"""
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    
    # 1. 标题页
    create_title_slide(prs)
    
    # 2. 项目简介
    create_content_slide(prs, "项目简介", [
        "• 使用机器学习方法预测共享单车租赁数量",
        "• 基于历史数据中的环境因素和季节信息",
        "• 帮助优化车辆调度，预测需求高峰",
        "• 提高运营效率和用户体验"
    ])
    
    # 3. 数据概览
    create_two_column_slide(prs, "数据概览", 
        ["数据来源", "• Capital Bikeshare系统", "• 华盛顿特区", "• 2011-2012年数据"],
        ["数据规模", "• 每小时数据：17,379条", "• 每日数据：731条", "• 12个特征维度"]
    )
    
    # 4. 模型性能
    create_metrics_slide(prs, "模型性能指标（随机森林）", [
        ("R² 分数", "0.9207", "解释92%变异"),
        ("RMSE", "50.11", "均方根误差"),
        ("MAE", "31.10", "平均绝对误差"),
        ("过拟合", "2.3%", "训练-测试差异")
    ])
    
    # 5. 特征重要性
    create_content_slide(prs, "特征重要性分析", [
        "1. hr (小时) - 64.6% ⭐ 最重要",
        "   • 一天中不同时段需求差异巨大",
        "   • 早高峰、晚高峰是需求高峰",
        "",
        "2. temp (温度) - 12.0%",
        "   • 温度越高，租车需求越高",
        "",
        "3. yr (年份) - 8.7%",
        "   • 2012年比2011年需求显著增长",
        "",
        "4. workingday (工作日) - 6.1%",
        "   • 工作日和周末的需求模式不同"
    ])
    
    # 6. 核心发现
    create_content_slide(prs, "核心发现", [
        "✅ 小时因素是影响需求的最重要因素（64.6%）",
        "   • 时段因素远超温度、天气等因素",
        "",
        "✅ 模型预测精度高（R² = 0.92）",
        "   • 能够准确预测中高需求水平",
        "",
        "✅ 运营策略建议",
        "   • 按时段动态调度车辆",
        "   • 早高峰、晚高峰增加投放",
        "   • 深夜时段减少投放"
    ])
    
    # 7. 模型对比
    create_two_column_slide(prs, "模型对比", 
        ["线性回归", "• R²: 0.3880", "• RMSE: 139.21", "• 简单但性能较低"],
        ["随机森林", "• R²: 0.9207", "• RMSE: 50.11", "• 性能优秀"]
    )
    
    # 8. 技术架构
    create_content_slide(prs, "技术架构", [
        "数据层",
        "  • 数据加载与清洗",
        "  • 特征工程与标准化",
        "",
        "模型层",
        "  • 随机森林回归（1000棵树）",
        "  • 最大深度：10层",
        "",
        "评估层",
        "  • 多种评估指标（R², RMSE, MAE）",
        "  • 残差分析",
        "  • 特征重要性分析"
    ])
    
    # 9. 结论
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = BLUE_PRIMARY
    
    title_shape = slide.shapes.title
    title_shape.text = "结论与展望"
    title_para = title_shape.text_frame.paragraphs[0]
    title_para.font.size = Pt(36)
    title_para.font.bold = True
    title_para.font.color.rgb = WHITE
    
    content_box = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(4))
    content_frame = content_box.text_frame
    content_frame.text = "• 成功建立了高精度的共享单车需求预测模型"
    para = content_frame.paragraphs[0]
    para.font.size = Pt(20)
    para.font.color.rgb = WHITE
    
    items = [
        "• 小时因素是核心驱动因素，占64.6%的重要性",
        "• 模型R²达到0.92，预测精度高",
        "• 为运营决策提供了数据支持",
        "",
        "未来优化方向：",
        "• 超参数调优",
        "• 特征工程优化",
        "• 实时预测系统"
    ]
    
    for item in items:
        p = content_frame.add_paragraph()
        p.text = item
        p.font.size = Pt(18)
        p.font.color.rgb = WHITE
        p.space_after = Pt(10)
    
    # 保存
    output_path = "../doc/共享单车租赁预测系统.pptx"
    prs.save(output_path)
    print(f"✅ PPT已创建：{output_path}")

if __name__ == "__main__":
    main()

