import pandas as pd
import numpy as np

import pandas.io.sql
import pyodbc

import xlrd

#mappin formats of excel possibilities (some problems with that, best way to go in the future)
values_type = {
    "object": "varchar(255)",
    "float64": "float",
    "int64": "BIGINT",
    "datetime64[ns]": "date"
}

print('- Initialize')

#SQL connection configs
server = 'YOUR SQL SERVER' 
database = 'YOUR DATABASE' 
username = 'YOUR USERNAME' 
password = 'YOUR PASSWORD' 

# create Connection and Cursor objects
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = conn.cursor()


cursor.execute("Select @@version")
results = cursor.fetchone()
    # Check if anything at all is returned
if results:
    print('- Db Connection Sucessfull')
else:
    print('- Db Connection Fail')       


# read data
file_excel = pd.ExcelFile(YOUR_EXCEL_FILE.xlsx)
data1 = pd.read_excel(file_excel)

#Transform all data types to string ( i have a problem with that :c ) (alternative way)
data = data1.astype(str)

n_rows = data[data.columns[0]].count()


#create a list of excel headers
new_vet = list(data.columns.values)
new_vet_types = list(data.dtypes)

vetor = []
vetor_aux = []
aux_querry_create = ''
aux_querry_insert = ''
aux_querry_insert2 = ''

for x in range(len(new_vet)):

    if x == len(new_vet) - 1:
        vetor.append(new_vet[x] + ' ' + values_type[new_vet_types[x].name])
        vetor_aux.append(new_vet[x])
        aux_querry_insert += ' ?'
    else:
        vetor.append(new_vet[x] + ' ' + values_type[new_vet_types[x].name] + ', ')  
        vetor_aux.append(new_vet[x] + ',')  
        aux_querry_insert += ' ?,'  
    aux_querry_create += vetor[x]
    aux_querry_insert2 += vetor_aux[x]
    
    
#name of table you will create
name_new_table = 'YOUR_TABLE_NAME'

#make the comands in sql
make_querry = 'CREATE TABLE [' + name_new_table + '] (' + aux_querry_create + ')'
make_insert = 'INSERT INTO [' + name_new_table + '] (' + aux_querry_insert2 + ') VALUES ('+aux_querry_insert+')'

print('- Creating Table')

# execute create table
try:
    cursor.execute(make_querry)
    conn.commit()
    print('- Table Create Sucessfull')
except pyodbc.ProgrammingError:
    print('- Table Create Ignored')
    pass

all_values = []

print('- Copying Values')


print('- Saving Values')

for r in range(1, n_rows):
    
    cells_values = []  

    for x in range(len(new_vet)):
   

        cells_values.append(data.iloc[r, x])
 
 
    all_values.append(cells_values)

# Execute sql Query

cursor.fast_executemany = True
param = list(tuple(row) for row in all_values)

for y in range(len(all_values)):
    cursor.execute(make_insert, param[y])

# Commit the transaction
conn.commit()

print('- Values Saved')


# Close the database connection
conn.close()

print('- Finish')
