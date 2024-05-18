from systemfiles.database.databaseconnection import get_sales_from_db
from systemfiles.database.databasefile import read_database_file
from systemfiles.cleandata.cleandatawmap import datacleaner
from systemfiles.reportingfiles.generatereportingfiles import export_reporting_dbs_to_excel
from systemfiles.getcheckdetails.getchecks import Checks, get_checks
from systemfiles.hourlyvisitors.calculatehourlyvisitors import calculate_hourly_sales
import time
from datetime import timedelta
import pandas as pd
 


total_time = timedelta()

byear = 2024
bmonth = 5
bday= 1

eyear = 2024
emonth = 6
eday = 2


mod = 'sql'

if mod == 'sql':
    start_time = time.time()
    data = get_sales_from_db(byear,bmonth,bday,eday,emonth,eyear)
    
    elapsed_time = time.time() - start_time
    total_time += timedelta(seconds=elapsed_time)
    print(f"get sales data from SQL server took {elapsed_time/60} minutes")
else:
    start_time = time.time()
    data = read_database_file()

    elapsed_time = time.time() - start_time
    total_time += timedelta(seconds=elapsed_time)
    print(f"Reading database file took {elapsed_time/60} minutes")


start_time = time.time()
cleanData = datacleaner(data)
elapsed_time = time.time() - start_time 
total_time += timedelta(seconds=elapsed_time)
print(f"Cleaning data took {elapsed_time/60} minutes")

start_time = time.time()    
export_reporting_dbs_to_excel(year=byear,month=bmonth,data=cleanData)
elapsed_time = time.time() - start_time
total_time += timedelta(seconds=elapsed_time)
print(f"Generating reporting files took {elapsed_time/60} minutes")


start_time = time.time()    
calculate_hourly_sales(cleanData,year=byear,month=bmonth)
elapsed_time = time.time() - start_time
total_time += timedelta(seconds=elapsed_time)
print(f"Generating hourly sales took {elapsed_time/60} minutes")


start_time = time.time()  
check_objects = get_checks(clean_data=cleanData)
elapsed_time = time.time() - start_time
total_time += timedelta(seconds=elapsed_time)
print(f"Getting check details took {elapsed_time/60} minutes")


print(f"Total time: {total_time}")




