from docx import Document
import create_chart
import create_pychart1
import create_pychart2
import create_monthly_chart

doc = Document()
doc.save('output.docx')
create_monthly_chart.month_chart()

excel_name = input('请输入excel名字：')

while True:
    sheet = input('请输入要读取的表单：')
    create_chart.build_chart(excel_name, sheet)
    create_pychart1.pichart1(excel_name,sheet)
    create_pychart2.pichart2(excel_name,sheet)

    choice = input("是否继续读取其他表单？(y/n): ")
    if choice.lower() == 'n':
        break


