import pandas as pd

df1=pd.read_excel('sale.xlsx',sheet_name='六月')
df2=pd.read_excel('sale.xlsx',sheet_name='五月')

column1=df1.iloc[:,0].tolist()

# 创建DataFrame对象
df = pd.DataFrame(column1, columns=['姓名'])

num_columns = df1.shape[1]

for i in range(1, num_columns):
    column=df2.iloc[:,i].tolist()
    c_name=df2.iloc[:,i].name
    df[c_name]=column

person_sales = df1.iloc[:, 1:].sum(axis=1).tolist()
df['六月'+'销售额']=person_sales

person_sales2=df2.iloc[:,1:].sum(axis=1).tolist()
df['五月'+'销售额']=person_sales2

result = []
for i in range(len(person_sales)):
    diff = person_sales2[i] - person_sales[i]
    result.append(diff)

df['增长值']=result

# 创建Excel Writer对象，并指定文件名
with pd.ExcelWriter('example.xlsx', engine='xlsxwriter') as writer:
    # 将DataFrame写入Excel表格
    df.to_excel(writer, sheet_name='Sheet1', startrow=0, startcol=0, index=False)









