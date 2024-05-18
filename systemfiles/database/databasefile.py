import os
import pandas as pd

def read_database_file():
    input_file = 'data.csv'
    data_folder = os.path.join(os.getcwd(), "Data")


    file_path = os.path.join(data_folder, input_file)
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File '{input_file}' not found in the 'Data' folder.")

 
    data = pd.read_csv(file_path)

    data['MainMenuKey'] = data['MainMenuKey'].fillna('NonMenu')
    data['MainMenuText'] = data['MainMenuText'].fillna('NonMenu')
    data['Menu'] = data['Menu'].fillna('NonMenu')
    data['OrderSource'] = data['OrderSource'].fillna('Alacarte')
    data.sort_values(by='TransactionID')
    data['Menu'] = data['Menu'].str.split('-').str[0]
    data['Date'] = pd.to_datetime(data['Date'])
    

    data.info()


    return data