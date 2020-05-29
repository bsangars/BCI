import pandas
from pandas.plotting import scatter_matrix
import matplotlib.pyplot as plt
import numpy as np

location =r"C:\BhargavsWorkspace\FlatFiles\IPLdata.csv"
names =['Match_sk','Match_id','Team1','Team2','Match_date','Season_Year','Venue','City','Country','Toss_Winner','Match_Winner','Toss_Name','Win_Type','Outcome','MOM','WinMargin','Country_id']
dataset = pandas.read_csv(location,skiprows=[0],names=[0])
df=pandas.DataFrame(dataset, columns =names)
# print(df.head(20))
# file_name =r"C:\Workspace\FlatFiles\IPLdatatranspose.csv"
# with open(file_name,'a') as t:
#     df.transpose().to_csv(t, header=True)  
df1 =df['MOM'].value_counts().head(10)
df1['Count']= df.groupby('Team1').count()
df1.plot(kind='bar',x='Team1', y='Count')
#df1.plot(kind='bar',x='MOM', y='Team1')
plt.show()
# df1= df.loc[['Bangalore','MOM']]
# print(df1)