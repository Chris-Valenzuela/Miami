import pandas as pd
from pandas import DataFrame
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.pyplot import figure
df = pd.read_excel('datasets/LEGO.xlsx', sheet_name = 'T2')

newdf = pd.DataFrame()

newdf = df.copy().loc[5:]
# newdf.dropna(how='all', inplace=True)

# newdf.replace(0, np.NaN).dropna(how='all', inplace=True)
# newdf


rowName = newdf.loc[5][1:]
rowName
rowNameDict = rowName.to_dict()


#this works
for x in range(1, len(df.columns)):
    
    newdf = newdf.rename(columns={f'Unnamed: {x}': rowNameDict[f'Unnamed: {x}']})
    
newdf.dropna(inplace = True)



newdf.set_index('Unnamed: 0', inplace= True)
# newdf

newdf = newdf.iloc[0:]
newdf = newdf.replace('-', 0)

column_max = newdf.max()
column_max
df_max = column_max.max()
normalized_df = newdf /df_max

# newdf



# FOR SOME BIZZARE REASON IF I DONT CONVERT FROM FLOAT TO NUMPY.FLOT64 THNE I GET A ISNAN ERROR WHEN I TRY TO PLOT MY NUMBERS. THIS IS HOW I CONVERT VIA GOOGLE
for col in normalized_df.columns:
    normalized_df[col] = pd.to_numeric(normalized_df[col], errors='coerce')
# newnormal_float64 = np.float64(newnormal.copy())

for col in newdf.columns:
    newdf[col] = pd.to_numeric(newdf[col], errors='coerce')

# normalized_df


# fig = plt.figure(figsize=(12,12))
# ax = fig.add_subplot(111)
# ax.matshow(newnormal, cmap=plt.cm.RdYlGn)

# tester.to_excel('datasets/playgrounds.xlsx')
# newnormal.to_excel('datasets/playground.xlsx')
plt.figure(figsize=(15, 15), dpi=100)
plt.pcolor(newdf)
plt.yticks(np.arange(0, len(newdf.index), 1), newdf.index)
plt.xticks(np.arange(0, len(newdf.columns), 1), newdf.columns, rotation=90)
plt.show()
# plt.matshow(normalized_df)