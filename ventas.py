import pandas as pd


archivo='ventas_dia2.xlsx'
df=pd.read_excel(archivo,sheet_name= 'Sheet1')

df['ID']=range(1,len(df)+1)
df=df[["ID","FechaVenta","Zona","Producto","PrecioVenta","Costo"]]


print(df)
#print(df.info())
#print(df.dtypes)
#print(df.shape)

