import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
import pandas as pd
from docx import Document
from docx.shared import Cm

def pichart2(excel_name,sheet):
    # 清除之前的绘图状态
    plt.clf()

    try:
        df = pd.read_excel(excel_name, sheet)
    except Exception:
        print('文件没有找到！')
        return

    n = df.iloc[:,0]  # 获取除了姓名列之外的列作为品牌标签
    names=n.tolist()
    person_colors = ['#FFC000', '#5B9BD5', '#ED7D31', '#A5A5A5', '#F15A24']  # 每个品牌对应的颜色

    # 设置中文字体
    font = FontProperties(fname='simsunb.ttf', size=12)  # 替换为你的中文字体文件路径
    plt.rcParams['font.family'] = 'SimSun'

    person_sales = df.iloc[:, 1:].sum(axis=1)  # 计算每个人的销售量之和

    plt.pie(person_sales, labels=names, autopct='%1.1f%%', colors=person_colors)
    plt.title(sheet+'个人每月销售量占比')
    plt.axis('equal')

    # 添加品牌名称及颜色方块标签
    legend_labels = [f'{brand}: {sales}' for brand, sales in zip(names, person_sales)]
    legend_handles = [plt.Rectangle((0, 0), 1, 1, color=color) for color in person_colors]
    plt.legend(legend_handles, legend_labels, loc='lower center', bbox_to_anchor=(0.5, -0.2), ncol=4)

    # 调整图表布局，增加底部空白
    plt.subplots_adjust(bottom=0.2)

    # 可选：将图形保存为图像文件
    plt.savefig('plot2.png')

    # 打开现有的Word文档
    doc = Document('output.docx')

    # 在指定的段落或表格中添加图像
    doc.add_picture('plot2.png', width=Cm(10), height=Cm(7))

    # 保存Word文档
    doc.save('output.docx')


if __name__=="__main__":
    excel_name = input('请输入excel名字：')
    sheet = input('请输入要读取的表单：')
    pichart2(excel_name, sheet)
    plt.show()
