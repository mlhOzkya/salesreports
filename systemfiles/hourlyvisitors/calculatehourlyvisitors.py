import pandas as pd

import os



def calculate_hourly_total_sales(df):
    
    hourlysales = df.groupby(['Store', 'Region', 'Order Type', 'DayofWeek', 'Hour'])['Sales'].sum().reset_index()

    hourlysales['WorkingDays'] = hourlysales['Order Type'] + "_" + hourlysales['DayofWeek']

    hourlysales = hourlysales.pivot_table(index=['Store', 'Region','Hour'],
                                              columns='WorkingDays', 
                                              values='Sales',
                                              fill_value=0).reset_index()
    
    

    return hourlysales

def calculate_hourly_total_visitors(df):
    
    visitors = df[df['Type'] == 'Visitor']

    hourlyvisitors = visitors.groupby(['Store', 'Region', 'Order Type', 'DayofWeek', 'Hour'])['Quantity'].sum().reset_index()

    hourlyvisitors['WorkingDays'] = hourlyvisitors['Order Type'] + "_" + hourlyvisitors['DayofWeek']

    hourlyvisitors = hourlyvisitors.pivot_table(index=['Store', 'Region','Hour'],
                                              columns='WorkingDays', 
                                              values='Quantity',
                                              fill_value=0).reset_index()
    
 

    return hourlyvisitors



def calculate_active_days(df):
   
    visitors = df[df['Type'] == 'Visitor']

    workingdays = visitors.groupby(['Store', 'Region', 'Date','Order Type', 'DayofWeek',])['Quantity'].sum().reset_index()

    activedays  =  workingdays[workingdays['Quantity']>0]

    activeweekdays = activedays.groupby(['Store', 'Region', 'Order Type', 'DayofWeek']).size().reset_index(name='Active Days')

    totalactivedays = activedays.groupby(['Store', 'Region', 'Order Type']).size().reset_index(name='Active Days')

    activeweekdays['WorkingDays'] = activeweekdays['Order Type'] + "_" + activeweekdays['DayofWeek']

    totalactivedays = totalactivedays.pivot_table(index=['Store', 'Region'],
                                              columns='Order Type', 
                                              values='Active Days',
                                              fill_value=0).reset_index()

    activeweekdays = activeweekdays.pivot_table(index=['Store', 'Region'],
                                              columns='WorkingDays', 
                                              values='Active Days',
                                              fill_value=0).reset_index()
    
    merged_data = pd.merge(activeweekdays, totalactivedays, on=['Store', 'Region'], how='left')
    
    

    return merged_data 



    
def calculate_hourly_sales(df, year, month):
    output_folder = os.path.join(os.getcwd(), 'Output', f'{year}', f'{month:02d}', 'ReportingFiles', 'Hourly Sales')
    os.makedirs(output_folder, exist_ok=True)  
    

    excel_path = os.path.join(output_folder, 'Hourly Sales.xlsx')


    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:

        hourly_sales_df = calculate_hourly_total_sales(df)
        hourly_sales_df.to_excel(writer, sheet_name='Hourly Sales', index=False)


        hourly_visitors_df = calculate_hourly_total_visitors(df)
        hourly_visitors_df.to_excel(writer, sheet_name='Hourly Total Visitors', index=False)

  
        working_days_df = calculate_active_days(df)
        working_days_df.to_excel(writer, sheet_name='Working Days', index=False)