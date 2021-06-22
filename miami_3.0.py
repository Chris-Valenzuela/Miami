import pandas as pd
from pandas import DataFrame
from openpyxl import load_workbook


filename = 'datasets/ChaseCAExample.xlsx'

xls = pd.ExcelFile(filename)
for sheet in xls.sheet_names[2:4]:
    # print(sheet)
    sheet_name = sheet

    # sheet_name = 'T1'



    df = pd.read_excel(filename, sheet_name = sheet_name)

    newdf = pd.DataFrame()
    newdf = df.copy().loc[5:]
    # newdf.dropna(how='all', inplace=True)

    # newdf.replace(0, np.NaN).dropna(how='all', inplace=True)
    newdf

    #to do: make this more dynamic
    rowName = newdf.loc[5][1:]
    rowName
    rowNameDict = rowName.to_dict()


    #this works
    for x in range(1, len(df.columns)):

        newdf = newdf.rename(columns={f'Unnamed: {x}': rowNameDict[f'Unnamed: {x}']})

    newdf.dropna(inplace = True)



    newdf.set_index('Unnamed: 0', inplace= True)
    

    newdf = newdf.iloc[0:]
    newdf = newdf.replace('-', 0)

    # column_max = newdf.max()
    # column_max
    # df_max = column_max.max()
    # normalized_df = newdf /df_max

    normalized_df = newdf
    
    # FOR SOME BIZZARE REASON IF I DONT CONVERT FROM FLOAT TO NUMPY.FLOT64 THNE I GET A ISNAN ERROR WHEN I TRY TO PLOT MY NUMBERS. THIS IS HOW I CONVERT VIA GOOGLE
    for col in normalized_df.columns:
        normalized_df[col] = pd.to_numeric(normalized_df[col], errors='coerce')
    # newnormal_float64 = np.float64(newnormal.copy())

    # for col in newdf.columns:
    #     newdf[col] = pd.to_numeric(newdf[col], errors='coerce')

    x = len(normalized_df.columns)
    y = len(normalized_df.index)

    # normalized_df
    

    normalCol = list(normalized_df.columns)

    #TO DO: make sure each index is unique. temp fix is "reset_inndex(drop=True)"
    normalized_df = normalized_df.reset_index(drop=True).style.background_gradient(cmap='viridis')
    
    
    book = load_workbook(filename)
    writer = pd.ExcelWriter(filename, engine='openpyxl') 

    # workbook = writer.book

    writer.book = book

    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
    normalized_df.to_excel(writer, sheet_name + '_heatmap', columns=normalCol)
    writer.save()

