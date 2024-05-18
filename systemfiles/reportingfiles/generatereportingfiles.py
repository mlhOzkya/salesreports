import os
import pandas as pd

def generate_product_mix(year,month,data):
    product_mix = data.groupby(['Product', 'Type', 'Category', 'Group', 'Detail', 'TDFR', 'Order Type']).agg({
        'Quantity': 'sum',
        'Sales': 'sum'
    }).reset_index()

    td_alacarte_sales = product_mix.loc[(product_mix['Order Type'] == 'Alacarte') & (product_mix['TDFR'] == 'TD')].copy()
    td_alacarte_sales.rename(columns={'Quantity': 'TD Alacarte Quantity', 'Sales': 'TD Alacarte Sales'}, inplace=True)

    fr_alacarte_sales = product_mix.loc[(product_mix['Order Type'] == 'Alacarte') & (product_mix['TDFR'] == 'FR')].copy()
    fr_alacarte_sales.rename(columns={'Quantity': 'FR Alacarte Quantity', 'Sales': 'FR Alacarte Sales'}, inplace=True)

    total_alacarte_sales = product_mix.loc[product_mix['Order Type'] == 'Alacarte'].copy()
    total_alacarte_sales.rename(columns={'Quantity': 'Total Alacarte Quantity', 'Sales': 'Total Alacarte Sales'}, inplace=True)
    totalalacartevisitors = total_alacarte_sales.loc[total_alacarte_sales['Type'] == 'Visitor', 'Total Alacarte Quantity'].sum()

    td_package_sales = product_mix.loc[(product_mix['Order Type'] == 'Package') & (product_mix['TDFR'] == 'TD')].copy()
    td_package_sales.rename(columns={'Quantity': 'TD Package Quantity', 'Sales': 'TD Package Sales'}, inplace=True)
    

    fr_package_sales = product_mix.loc[(product_mix['Order Type'] == 'Package') & (product_mix['TDFR'] == 'FR')].copy()
    fr_package_sales.rename(columns={'Quantity': 'FR Package Quantity', 'Sales': 'FR Package Sales'}, inplace=True)

    total_package_sales= product_mix.loc[product_mix['Order Type'] == 'Package'].copy()
    total_package_sales.rename(columns={'Quantity': 'Total Package Quantity', 'Sales': 'Total Package Sales'}, inplace=True)
    totalpackagevisitors = total_package_sales.loc[total_package_sales['Type'] == 'Visitor', 'Total Package Quantity'].sum()

    td_total_sales = product_mix.loc[product_mix['TDFR'] == 'TD'].copy()
    td_total_sales.rename(columns={'Quantity': 'TD Total Quantity', 'Sales': 'TD Total Sales'}, inplace=True)
    totaltdvisitors = td_total_sales.loc[td_total_sales['Type'] == 'Visitor', 'TD Total Quantity'].sum()

    fr_total_sales = product_mix.loc[product_mix['TDFR'] == 'FR'].copy()
    fr_total_sales.rename(columns={'Quantity': 'FR Total Quantity', 'Sales': 'FR Total Sales'}, inplace=True)  
    totalfrvisitors = fr_total_sales.loc[fr_total_sales['Type'] == 'Visitor', 'FR Total Quantity'].sum()  

    total_sales = product_mix.copy()
    total_sales.rename(columns={'Quantity': 'Total Quantity', 'Sales': 'Total Sales'}, inplace=True)  

    merged_df = pd.merge(product_mix, td_alacarte_sales, on=['Product', 'Type', 'Category', 'Group', 'Detail', 'TDFR', 'Order Type'], how='left')
    merged_df = pd.merge(merged_df, fr_alacarte_sales, on=['Product', 'Type', 'Category', 'Group', 'Detail', 'TDFR', 'Order Type'], how='left')
    merged_df = pd.merge(merged_df, total_alacarte_sales, on=['Product', 'Type', 'Category', 'Group', 'Detail', 'TDFR', 'Order Type'], how='left')
    merged_df = pd.merge(merged_df, td_package_sales, on=['Product', 'Type', 'Category', 'Group', 'Detail', 'TDFR', 'Order Type'], how='left')
    merged_df = pd.merge(merged_df, fr_package_sales, on=['Product', 'Type', 'Category', 'Group', 'Detail', 'TDFR', 'Order Type'], how='left')
    merged_df = pd.merge(merged_df, total_package_sales, on=['Product', 'Type', 'Category', 'Group', 'Detail', 'TDFR', 'Order Type'], how='left')
    merged_df = pd.merge(merged_df, td_total_sales, on=['Product', 'Type', 'Category', 'Group', 'Detail', 'TDFR', 'Order Type'], how='left')
    merged_df = pd.merge(merged_df, fr_total_sales, on=['Product', 'Type', 'Category', 'Group', 'Detail', 'TDFR', 'Order Type'], how='left')
    merged_df = pd.merge(merged_df, total_sales, on=['Product', 'Type', 'Category', 'Group', 'Detail', 'TDFR', 'Order Type'], how='left')

    merged_df.fillna(0, inplace=True)

  
    merged_df = merged_df[['Product', 'Type', 'Category', 'Group', 'Detail',
                           'TD Alacarte Quantity', 'FR Alacarte Quantity', 'Total Alacarte Quantity',
                           'TD Package Quantity', 'FR Package Quantity', 'Total Package Quantity',
                           'TD Alacarte Sales', 'FR Alacarte Sales', 'Total Alacarte Sales',
                           'TD Package Sales', 'FR Package Sales', 'Total Package Sales',
                           'TD Total Quantity', 'FR Total Quantity', 'Total Quantity',
                           'TD Total Sales', 'FR Total Sales', 'Total Sales',
                           ]]
    

    merged_df = merged_df.groupby(['Product', 'Type', 'Category', 'Group', 'Detail']).sum().reset_index()

    totalSystemSales = merged_df['Total Sales'].sum()
    totaltdsales = merged_df['TD Total Sales'].sum()
    totalfrsales = merged_df['FR Total Sales'].sum()
    totalalacartesales = merged_df['Total Alacarte Sales'].sum()
    totalpackagesales = merged_df['Total Package Sales'].sum()
    
    print(f"Total System Sales: {totalSystemSales:,.2f}\n"
      f"Total TD Sales: {totaltdsales:,.2f}\n"
      f"Total FR Sales: {totalfrsales:,.2f}\n"
      f"Total Alacarte Sales: {totalalacartesales:,.2f}\n"
      f"Total Package Sales: {totalpackagesales:,.2f}\n"
      f"Total TD Visitors: {totaltdvisitors:,.2f}\n"
      f"Total FR Visitors: {totalfrvisitors:,.2f}\n"
      f"Total Alacarte Visitors: {totalalacartevisitors:,.2f}\n"
      f"Total Package Visitors: {totalpackagevisitors:,.2f}\n"
      f"Total Visitors: {totaltdvisitors + totalfrvisitors:,.2f}")


    output_folder = output_folder = os.path.join(os.getcwd(),   'Output',f'{year}',f'{month:02d}' ,'ReportingFiles','Product Mix')
    os.makedirs(output_folder, exist_ok=True) 
    excel_path = os.path.join(output_folder, 'product_mix.xlsx')


    merged_df.to_excel(excel_path, index=False)
    print(f"Product Mix exported to '{excel_path}'.")

    return merged_df

def generate_store_mix(data,year,month):

    product_mix = data.groupby(['Product', 'Store','Type', 'Category', 'Group', 'Detail', 'TDFR', 'Order Type']).agg({
        'Quantity': 'sum',
        'Sales': 'sum'
    }).reset_index()

  
    
    pivot_table = product_mix.pivot_table(index=['Product','Order Type','Type', 'Category', 'Group', 'Detail'], columns='Store', values=['Quantity', 'Sales'], fill_value=0)

    
    pivot_table.columns = [f'{col[0]}_{col[1]}' for col in pivot_table.columns]

    
    pivot_table.reset_index(inplace=True)

    output_folder = output_folder = os.path.join(os.getcwd(),   'Output',f'{year}' ,f'{month:02d}','ReportingFiles','Product Mix')
    os.makedirs(output_folder, exist_ok=True)  
    excel_path = os.path.join(output_folder, 'Store_mix.xlsx')

    
    pivot_table.to_excel(excel_path, index=False)
    generate_product_mix(data=data,year=year,month=month)
    print(f"Store Mix exported to '{excel_path}'.")

def export_reporting_nd_db_to_excel(year, month,data, max_rows_per_sheet=1000000):
    
    output_folder = os.path.join(os.getcwd(), 'Output', f'{year}',f'{month:02d}', 'ReportingFiles')
    os.makedirs(output_folder, exist_ok=True)
    excel_path = os.path.join(output_folder, 'reportingDatabaseNoDate.xlsx')
    data = data.groupby(['Store',"TDFR","Region",  'Order Type', 'Product','Type', 'Detail', 'Category', 'Group']).agg({'Quantity': 'sum', 'Sales': 'sum'}).reset_index()

    
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        start_row = 0
        total_rows = data.shape[0]
        sheet_count = 1

        while start_row < total_rows:
            end_row = min(start_row + max_rows_per_sheet, total_rows)

            
            data.iloc[start_row:end_row].to_excel(writer, index=False, 
                                                          sheet_name=f'Sheet{sheet_count}')

            start_row = end_row
            sheet_count += 1

            
    generate_store_mix(data=data,year=year,month=month)
    print(f"Reporting DB exported to '{excel_path}' with {sheet_count - 1} sheets.")
    
def export_reporting_dbs_to_excel(year, month, data, max_rows_per_sheet=1000000):
    
    output_folder = os.path.join(os.getcwd(), 'Output', f'{year}',f'{month:02d}', 'ReportingFiles')
    os.makedirs(output_folder, exist_ok=True)
    excel_path = os.path.join(output_folder, 'reportingDatabase.xlsx')
    data = data.groupby(['Store',"TDFR","Region", 'Date', 'Order Type', 'Product','Type', 'Detail', 'Category', 'Group']).agg({'Quantity': 'sum', 'Sales': 'sum'}).reset_index()

    
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        start_row = 0
        total_rows = data.shape[0]
        sheet_count = 1

        while start_row < total_rows:
            end_row = min(start_row + max_rows_per_sheet, total_rows)

            
            data.iloc[start_row:end_row].to_excel(writer, index=False, 
                                                          sheet_name=f'Sheet{sheet_count}')

            start_row = end_row
            sheet_count += 1


    export_reporting_nd_db_to_excel(year = year, month = month,data=data)

    print(f"Reporting DB exported to '{excel_path}' with {sheet_count - 1} sheets.")
    
