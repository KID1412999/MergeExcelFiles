# coding: utf-8
import pandas as pd
import numpy as np
import os
path_='C://Users//Administrator//Desktop//edata'
path=os.listdir(path_)
dsheet=[]
s=pd.DataFrame()

for i in path:#批量读取数据表
	key='姓名'
	excel_path =path_+'\\'+i
	d=pd.read_excel(excel_path, sheet_name=None)
	dsheet.append(d['Sheet1'].drop_duplicates(subset=key, keep='first', inplace=False))

def djoin(d1,d2):#合并数据表
	d_new=pd.merge(d2,d1,on=key,how='outer').fillna('nan')
	c=[]
	for i in d1.columns:
		if i in d2.columns and i!=key:
			c.append(i)
	for j in c:
		w=j
		n=[]
		for i in range(len(d_new[w+'_x'])):
			if d_new[w+'_x'][i]!='nan':
				n.append(d_new[w+'_x'][i])
			else:
				n.append(str(d_new[w+'_y'][i]))
		d_new.loc[:,w+'_x']=n
		del d_new[w+'_y']
		d_new=d_new.rename(columns={w+'_x':w})
	return d_new

for i in range(len(dsheet)):#两两结合
	if i==0:
		s=djoin(dsheet[i],dsheet[i+1])
	else:
		s=djoin(s,dsheet[i])

#将合并结果存入Excel文件
writer = pd.ExcelWriter(path_+'//output.xlsx')
df1 = pd.DataFrame(data=s)
df1.to_excel(writer,'Sheet1')
writer.save()
print(s)
