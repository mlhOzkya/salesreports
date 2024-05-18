import pyodbc
import pandas as pd
import os

def get_sales_from_db(year,month,d1,d2,next_month,next_year):
 
    server = 'xx'
    database = 'xx'
    driver = '{xx}'

    conn_str = f"DRIVER={driver};SERVER={server};DATABASE={database};Trusted_connection=yes"

    conn = pyodbc.connect(conn_str)

    
    date1 = f"{year}-{month:02d}-{d1:02d}"
    date2 = f"{next_year}-{next_month:02d}-{d2:02d}"

    # Example query to retrieve sales data from the database

    query = f"""

DECLARE @date1 AS DATETIME;
DECLARE @date2 AS DATETIME;

SET @date1 = '{date1}';
SET @date2 = '{date2}';

SELECT 
	br.StoreName AS [Store],
	t.OrderDate AS [Date],
    DATEPART(HOUR, t.OrderDateTime) AS [Hour],
    t.Check AS [Check],
    t.MenuKey AS [MainMenuKey],
    t.MenuText AS [MainMenuText],
    t.TransactionID as [TransactionID],
    t.Menu as [Menu],
    h.OrderSource AS [OrderSource],
	(
	CASE ISNULL(h.Type, 0)
		WHEN 1 THEN 'Alacarte'
		WHEN 2 THEN 'Package'
		ELSE 'UNKNOWN'
	END
	) AS [Order Type],
	t.Product AS [Product],
	t.Quantity AS [Quantity],
	(t.ExtendedPrice * ISNULL((h.AmountDue / NULLIF(h.SubTotal,0)), 0)) AS [Sales]
FROM
    TransactionDetails t
    JOIN OrderHeaders o ON o.OrderKey = t.TransactionKey
    JOIN Stores b ON b.StoreID = t.StoreID
WHERE
	t.OrderDateTime BETWEEN @date1 AND @date2
   
    """
    cursor = conn.cursor()
    cursor.execute(query)	

    columns = [column[0] for column in cursor.description]

    rows = cursor.fetchall()
    data = pd.DataFrame.from_records(rows, columns=columns)

 

    conn.close()


    data['Sales'] = data['Sales'].astype(float)
    data['Quantity'] = data['Quantity'].astype(float)
    data['Date'] = pd.to_datetime(data['Date'], format='%Y-%m-%d')
    data = data[(data['Date'].dt.month == month) ]

    output_folder = output_folder = os.path.join(os.getcwd(),   'Output',f'{year}',f'{month:02d}' ,'RawData')
    os.makedirs(output_folder, exist_ok=True)  
    csv_path = os.path.join(output_folder, 'data.csv')
    data.to_csv(csv_path, index=False)
    print(f"Data exported to '{csv_path}'.")


    data['MainMenuKey'] = data['MainMenuKey'].fillna('NonMenu')
    data['MainMenuText'] = data['MainMenuText'].fillna('NonMenu')
    data['Menu'] = data['Menu'].fillna('NonMenu')
    data['OrderSource'] = data['OrderSource'].fillna('Alacarte')
    data['Menu'] = data['Menu'].str.split('-').str[0]
    data['Date'] = pd.to_datetime(data['Date'])
    data.reset_index()
    data.info()

    

    return data
