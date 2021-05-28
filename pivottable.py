import pyodbc
import xlsxwriter
import pandas as pd

# Connection to SQL Server and SQL Script
conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=test01;'
                      'Database=Test;'
                      'Trusted_Connection=yes;')

sql_query = pd.read_sql_query('''
SELECT TOP (1000) [Agent]
      ,[Manager]
      ,[Date]
      ,[Accepted]
      ,[Rejected]
      ,[Total Talk Time]
  FROM [Test].[dbo].[TC_TEST]
  WHERE [Agent] = 'Test'
                              '''
                              ,conn)

# Save Data pulled from SQL Server into Excel file
df = pd.DataFrame(sql_query)
writer = pd.ExcelWriter(r'C:\Users\Deskop\ar_urltime.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='data', startrow=1, header=False, index=False)

# Get data from SQL Server
workbook = writer.book
worksheet = writer.sheets['data']
(max_row, max_col) = df.shape
column_settings = [{'header': column} for column in df.columns]

# Create Excel Table
worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings})
workbook.add_worksheet('pivot summary')

# Create Pivot Table
pivot= pd.pivot_table(df, index=['Agent'], values=['Total Talk Time'], aggfunc='sum')
pivot.to_excel(writer, sheet_name='pivot')
writer.save()
