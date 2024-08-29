import pandas as pd
import os
import openpyxl

os.chdir(r'd:\users\USER\miniconda3\lib\site-packages')

excel_one = pd.read_excel(r'C:\Users\USER\Documents\Comparativo Automático Excel\excel_uno.xlsx')
excel_two = pd.read_excel(r'C:\Users\USER\Documents\Comparativo Automático Excel\excel_dos.xlsx')


# Añadimos prefijo a las columnas del segundo excel para distinguirlas de las primeras
excel_two.columns = excel_two.columns.map(lambda x :'dos_' + x)

# Miramos que ambos excel tengan las columnas en los mismos formatos antes de realizar la comparativa
print("\n\nPrimer fichero:\n")
print(excel_one.info())
print("\n")
print("\n\nSegundo fichero:\n")
print(excel_two.info())

# Juntamos ambas tablas
join_tables= pd.concat([excel_one, excel_two], axis=1)
print(join_tables)


# Define the ranges to distinguis the columns of each table
xslx_two_col = join_tables.columns[join_tables.columns.str.contains('dos_')]

from_tabA= 0
to_tabA = len(xslx_two_col)

from_tabB= len(xslx_two_col)
to_tabB= len(join_tables.columns)

# Convert the ranges to list

list1 = pd.Series(range(from_tabA, to_tabA)).to_list()
list2 = pd.Series(range(from_tabB, to_tabB)).to_list()

# Create a third List based on the iterations of the values from List1 and List2
column_order = [val for pair in zip(list1, list2) for val in pair]

# Check column order
print(column_order)

# Apply the new column order to the df that contains both tables
df = join_tables.iloc[:,column_order]
print(df.head(4))

# Compare if the columns are equal
def add_equal_columns(df):
  for i in range(0, len(df.columns), 2):
      coll = df.columns[i]
      col2 = df.columns[i+1]
      new_col_name = f'{coll}_{col2}_equal'
      df = df.assign(**{new_col_name: df[coll] == df[col2]})
  return df

df = add_equal_columns(df.copy())  

print(df.head(4))

# Reorder the columns before saving the file

list_a = pd.Series(range(0, len(join_tables.columns))).to_list()
list_b = pd.Series(range(len(join_tables.columns), len(df.columns) )).to_list()

def new_order(list_a, list_b):
    new_order = []
    index_b = 0
    for i in range(0, len(list_a), 2):
        new_order.extend([list_a[i], list_a[i+1], list_b[index_b]])
        index_b += 1
    return new_order

new_order(list_a, list_b)


# Apply final order to the df
df_final= df.iloc[:, new_order(list_a, list_b)]

print(df_final.head())


# Export the df to an excel file
with pd.ExcelFile(r'C:\Users\USER\Documents\excel_equal.xlsx') as writer:
   df_final.to_excel(writer, index=False)