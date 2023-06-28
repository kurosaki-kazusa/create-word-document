import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
import numpy as np
from docx import Document
from docx.shared import Cm


def month_chart():
    # 清除之前的绘图状态
    plt.clf()
    # 创建一个包含月份的列表
    months = ['六月', '五月', '四月', '三月', '二月', '一月']

    # 创建一个包含品牌名称的列表
    brands = ['三星', '苹果', '小米', '华为']

    data=[]


    excel_file = pd.ExcelFile('sale.xlsx')
    sheet_names=excel_file.sheet_names

    for sheet in sheet_names:
        df = pd.read_excel('sale.xlsx', sheet_name=sheet)
        brand_labels = df.columns.tolist()[1:]
        brand_sales=[sum(df[brand]) for brand in brand_labels]
        data.append(brand_sales)

    # 设置颜色
    colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b']

    # 设置中文字体
    font = FontProperties(fname='simsunb.ttf', size=12)  # 替换为你的中文字体文件路径
    plt.rcParams['font.family'] = 'SimSun'

    # 创建柱状图
    fig, ax = plt.subplots()
    width = 0.15  # 设置每个柱子的宽度

    # 循环遍历每个品牌，并绘制对应的柱状图
    for i, brand in enumerate(months):
        x = np.arange(len(data[i]))  # 每个品牌的横坐标位置
        ax.bar(x + (i * width), data[i], width, label=brand)

    # 设置图表标题和坐标轴标签
    ax.set_title('Monthly Sales by Brand')
    ax.set_ylabel('Sales')

    for container in ax.containers:
        ax.bar_label(container, fmt='%.0f', label_type='edge', color='black', fontsize=7)

    # 设置横坐标刻度标签
    ax.set_xticks(np.arange(len(brands)) + (len(brands) * width) / 2)
    ax.set_xticklabels(brands)


    legend_labels = [f'{month}' for month in months]
    legend_handles = [plt.Rectangle((0, 0), 1, 1, color=color) for color in colors]
    plt.legend(legend_handles, legend_labels, loc='lower center', bbox_to_anchor=(0.5, -0.2), ncol=6)

    # 调整图表布局，增加底部空白
    plt.subplots_adjust(bottom=0.2)

    # 可选：将图形保存为图像文件
    plt.savefig('plot3.png')

    # 打开现有的Word文档
    doc = Document('output.docx')

    # 在指定的段落或表格中添加图像
    doc.add_picture('plot3.png', width=Cm(10), height=Cm(7))

    # 保存Word文档
    doc.save('output.docx')


if __name__=="__mian__":
    # 展示图表
    plt.show()











