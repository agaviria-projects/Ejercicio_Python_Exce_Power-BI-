import pandas as pd


archivo='ventas_dia2.xlsx'
df=pd.read_excel(archivo,sheet_name= 'Sheet1')

#Generar un ID
df['ID']=range(1,len(df)+1)

#Reordenar columnas
df=df[["ID","FechaVenta","Zona","Producto","PrecioVenta","Costo"]]



print(df)
#print(df.info())
#print(df.dtypes)
#print(df.shape)

