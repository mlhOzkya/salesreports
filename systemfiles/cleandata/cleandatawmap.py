import os
import xlwings as xw
import pandas as pd

def get_maps():

    folder_path = os.path.join(os.getcwd(), 'Data')
    file_path = os.path.join(folder_path, 'map.xlsx')


    app = xw.App(visible=False)
    workbook = app.books.open(file_path)

   
    replacement_sheet = workbook.sheets['Map']
    replacement_range = replacement_sheet.range('Replacement')
    replacement = replacement_range.value


    products_sheet = workbook.sheets['Map']
    products_range = products_sheet.range('Products')
    products = products_range.value


    tdFr_sheet = workbook.sheets['Map']
    tdFr_range = tdFr_sheet.range('tdFR')
    tdFr = tdFr_range.value

    workbook.close()
    app.quit()

    return replacement, products, tdFr

def assign_time_group(day):
    if day in [0, 1, 2, 3]:
         return 'Weekday'
    elif day == 4:
            return 'Friday'
    elif day in [5, 6]:
            return 'Weekend'

def update_order_source(data, index):

    if data.at[index, 'Order Type'] == 'Package':
        order_source_value = data.at[index, 'OrderSource'].lower()
        if 'trend' in order_source_value:
            data.at[index, 'OrderSource'] = 'Trendyol'
        elif 'sepet' in order_source_value:
            data.at[index, 'OrderSource'] = 'Yemek Sepeti'
        elif 'getir' in order_source_value:
            data.at[index, 'OrderSource'] = 'Getir'
        elif 'migros' in order_source_value:
            data.at[index, 'OrderSource'] = 'Migros'
        else:
            data.at[index, 'OrderSource'] = 'Sef'
    else:
        data.at[index, 'OrderSource'] = 'Alacarte'

def update_product_info(data, i, replacement, products):
    product_match = False
    for old_value, new_value in replacement:
        if old_value is not None and data.at[i, 'Product'].lower() == old_value.lower():
            data.at[i, 'Product'] = new_value
            break

    for p in products[1:]:
        if p[0] is not None and data.at[i, 'Product'].lower() == p[0].lower():
            product_match = True
            data.at[i, 'Type'] = p[1]
            data.at[i, 'Category'] = p[2]
            data.at[i, 'Group'] = p[3]
            data.at[i, 'Detail'] = p[4]
            break
    return product_match

def update_store_info(data, i, tdFr):
    store_match = False
    for td in tdFr[1:]:
        if td[0] is not None and data.at[i, 'Store'].lower() == td[0].lower():
            data.at[i, 'TDFR'] = td[2]
            data.at[i, 'Region'] = td[1]
            store_match = True
            break
    return store_match

def save_unmatched_data(unmatched, unmatchedstores,ordersourcecontrol):
    output_folder = os.path.join(os.getcwd(), 'UnmatchedData')
    os.makedirs(output_folder, exist_ok=True)

    if unmatched:
        new_sales_df = pd.DataFrame({'Product': list(set(unmatched))})
        csv_path = os.path.join(output_folder, 'unmatched_products.csv')
        new_sales_df.to_csv(csv_path, index=False)

    if unmatchedstores:
        new_stores_df = pd.DataFrame({'Stores': list(set(unmatchedstores))})
        csv_path = os.path.join(output_folder, 'unmatched_stores.csv')
        new_stores_df.to_csv(csv_path, index=False)

    if ordersourcecontrol:
        unmatched_source_df = pd.DataFrame({'Check': list(set(ordersourcecontrol))})
        csv_path = os.path.join(output_folder, 'ordersourceCheck.csv')
        unmatched_source_df.to_csv(csv_path, index=False)

def ordersource_control(data, i):
    ordersource = False
    if data.at[i, 'OrderSource'] == 'Alacarte' and data.at[i, 'Order Type'] != 'Alacarte':
        ordersource = True
    return ordersource



def datacleaner(data):
    data.sort_values(by='Check', inplace=True)
    data.reset_index(drop=True, inplace=True)
    replacement, products, tdFr = get_maps()
    unmatched = []
    unmatchedstores = []
    ordersourcecontrol = []
    day = data['Date'].dt.weekday

    for i in range(len(data)):
        data.at[i, 'DayofWeek'] = assign_time_group(day[i])
        update_order_source(data, i)
        product_match = update_product_info(data, i, replacement, products)
        store_match = update_store_info(data, i, tdFr)
        ordercontrol = ordersource_control(data, i)
        
        if not product_match:
            unmatched.append(data.at[i, 'Product'])
        if not store_match:
            unmatchedstores.append(data.at[i, 'Store'])
        if ordercontrol:
            ordersourcecontrol.append(data.at[i, 'Check'])

    if unmatched or unmatchedstores or ordersourcecontrol:
        save_unmatched_data(unmatched, unmatchedstores, ordersourcecontrol)

        raise Exception(f'Unmatched data found in data. The program has been stopped.')

    data.info()
   

  
    return data

