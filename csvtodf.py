import pandas
from pandas.plotting import scatter_matrix
import matplotlib.pyplot as plt

def importcsv(location):
    dataset = pandas.read_csv(location, skiprows=[0],header=[0])
    df=pandas.DataFrame(dataset)
    return df

for col in dFrame.columns:
    print(col)

def getColumns(dataset):
    columnarray=[]
    for col in dataset.columns:
         columnarray.append(col)
    
    return columnarray



