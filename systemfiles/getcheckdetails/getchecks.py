import pandas as pd


class ProductSale:
    def __init__(self, name, type, category, group, detail,quantity,sales,transactionid) -> None:
        self.name = name
        self.type = type
        self.category = category
        self.group = group
        self.detail = detail
        self.quantity = quantity
        self.sales = sales
        self.transactionid = transactionid 


class ComboProductSale:
    def __init__(self, name, type, category, group, detail,quantity,sales,mainmenukey,mainmenutext,transactionid) -> None:
        self.name = name
        self.type = type
        self.category = category
        self.group = group
        self.detail = detail
        self.quantity = quantity
        self.sales = sales
        self.mainmenukey = mainmenukey
        self.mainmenutext = mainmenutext
        self.transactionid = transactionid 


class Checks:
    def __init__(self, id, ordertype,store,date,hour,tdfr,ordersource,dayofweek):
        self.id = id
        self.ordertype = ordertype
        self.store = store
        self.date = date
        self.hour = hour
        self.tdfr = tdfr
        self.ordersource = ordersource
        self.dayofweek = dayofweek
        self.nonmenuproducts = []
        self.combomenuproducts = []
        

    def add_nonmenuproduct_sale(self, non_menu_product_sale):
        self.nonmenuproducts.append(non_menu_product_sale)
        self.nonmenuproducts.sort(key=lambda x: x.transactionid)

    def add_combomenuproduct_sale(self, combo_menu_product_sale):
        self.combomenuproducts.append(combo_menu_product_sale)
        self.combomenuproducts.sort(key=lambda x: x.transactionid)


def export_checks_to_excel(checks_list, file_name="output_by_ordersource.xlsx"):
    
    data_dict = {}

    for check in checks_list:
        for product in check.nonmenuproducts + check.combomenuproducts:
            order_source = product.ordersource if hasattr(product, 'ordersource') else check.ordersource
            if order_source not in data_dict:
                
                data_dict[order_source] = {
                    'Check ID': [],
                    'Order Type': [],
                    'Store': [],
                    'Date': [],
                    'Hour': [],
                    'TDFR': [],
                    'Day of Week': [],
                    'Product Name': [],
                    'Product Type': [],
                    'Category': [],
                    'Group': [],
                    'Detail': [],
                    'Quantity': [],
                    'Sales': [],
                    'Main Menu Key': [],
                    'Main Menu Text': [],
                    'Transaction ID': []
                }

           
            data = data_dict[order_source]
            data['Check ID'].append(check.id)
            data['Order Type'].append(check.ordertype)
            data['Store'].append(check.store)
            data['Date'].append(check.date)
            data['Hour'].append(check.hour)
            data['TDFR'].append(check.tdfr)
            data['Day of Week'].append(check.dayofweek)
            data['Product Name'].append(product.name)
            data['Product Type'].append(product.type)
            data['Category'].append(product.category)
            data['Group'].append(product.group)
            data['Detail'].append(product.detail)
            data['Quantity'].append(product.quantity)
            data['Sales'].append(product.sales)
            data['Transaction ID'].append(product.transactionid)
            if isinstance(product, ComboProductSale):
                data['Main Menu Key'].append(product.mainmenukey)
                data['Main Menu Text'].append(product.mainmenutext)
            else:
                data['Main Menu Key'].append(None)
                data['Main Menu Text'].append(None)

    
    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        
        for order_source, data in data_dict.items():
            df = pd.DataFrame(data)
            df.to_excel(writer, sheet_name=order_source, index=False)

 
def get_checks(clean_data):
    check_dict = {}
    for index, row in clean_data.iterrows():
        check_id = row['Check']
        if check_id not in check_dict:
            check_dict[check_id] = Checks(
                id=check_id,
                ordertype=row['Order Type'],
                store=row['Store'],
                date=row['Date'],
                hour=row['Hour'],
                tdfr=row['TDFR'],
                ordersource=row['OrderSource'],
                dayofweek=row['DayofWeek']
            )

        current_check = check_dict[check_id]
        if row['Menu'] == 'NonMenu':
            new_product_sale = ProductSale(
                name=row['Product'],
                type=row['Type'],
                category=row['Category'],
                group=row['Group'],
                detail=row['Detail'],
                quantity=row['Quantity'],
                sales=row['Sales'],
                transactionid=row['TransactionID']
            )
            current_check.add_nonmenuproduct_sale(new_product_sale)
        else:
            new_product_sale = ComboProductSale(
                name=row['Product'],
                type=row['Type'],
                category=row['Category'],
                group=row['Group'],
                detail=row['Detail'],
                quantity=row['Quantity'],
                sales=row['Sales'],
                mainmenukey=row['MainMenuKey'],
                mainmenutext=row['MainMenuText'],
                transactionid=row['TransactionID']
            )
            current_check.add_combomenuproduct_sale(new_product_sale)

    export_checks_to_excel(list(check_dict.values()))

    return list(check_dict.values())

    

    
    
    
    
