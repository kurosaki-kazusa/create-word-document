import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
from docx import Document
from docx.shared import Cm


def build_chart(excel_name, sheet):

    # 创建DataFrame对象
    try:
        df=pd.read_excel(excel_name,sheet_name=sheet)
    except Exception:
        print('文件没有找到！')
        return

    column1=df.iloc[:,0]
    row1=df.columns.tolist()[1:]

    df.set_index(column1)
    df = df.rename(columns={'Unnamed: 0': '姓名'})

    # 设置中文字体
    font = FontProperties(fname='simsunb.ttf', size=12)  # 替换为你的中文字体文件路径
    plt.rcParams['font.family'] = 'SimSun'

    # 设置图表数据
    bar_data = df.set_index('姓名')[row1]

    # 设置柱状图颜色
    colors = ['#FFC000', '#5B9BD5', '#ED7D31', '#A5A5A5']

    # 绘制横向柱状图
    ax = bar_data.plot(kind='barh', figsize=(8, 6), color=colors)

    # 设置图表标题和轴标签
    ax.set_title(sheet+'销售额')
    ax.set_ylabel('销售员')

    # 添加数据标签
    for container in ax.containers:
        ax.bar_label(container, fmt='%.0f', label_type='edge', color='black', fontsize=10)

    # 添加基线
    # 添加基线和标注
    for column, color in zip(bar_data.columns, colors):
        ax.axvline(bar_data[column].mean(), color=color, linestyle='--', label=f'{column}平均值')
        ax.annotate(int(bar_data[column].mean()), xy=(bar_data[column].mean(), 0), xytext=(0, -55),
                    textcoords='offset points', ha='right', va='bottom', rotation=0, color=color)

    # 显示图例
    ax.legend(ncol=4, loc='lower center', bbox_to_anchor=(0.5, -0.30))


    # 调整图表布局，增加底部空白
    plt.subplots_adjust(bottom=0.3)

    # 可选：将图形保存为图像文件
    plt.savefig('plot.png')

    # 打开现有的Word文档
    doc = Document('output.docx')

    # 在指定的段落或表格中添加图像
    doc.add_picture('plot.png', width=Cm(10), height=Cm(7))

    # 保存Word文档
    doc.save('output.docx')

    # 清除之前的绘图状态
    plt.clf()

if __name__ == "__main__":
    excel_name = input('请输入excel名字：')
    sheet = input('请输入要读取的表单：')
    build_chart(excel_name, sheet)
    # 显示图表
    plt.show()

