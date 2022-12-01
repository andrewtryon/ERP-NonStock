from dotenv import load_dotenv
load_dotenv()
import os
import pyodbc
import pandas as pd
import subprocess

if __name__ == '__main__':

    #Connection String
    conn_str = os.environ.get(r"sage_conn_str").replace("UID=;","UID=" + os.environ.get(r"sage_login") + ";").replace("PWD=;","PWD=" + os.environ.get(r"sage_pw") + ";") 
    #Establish sage connection
    print('Connecting to Sage')
    cnxn = pyodbc.connect(conn_str, autocommit=True)  
    
    #SQL Sage data into dataframe
    sql = """
        SELECT 
            CI_Item.ItemCode
        FROM 
            CI_Item CI_Item, 
            IM_ItemWarehouse IM_ItemWarehouse
        WHERE 
            CI_Item.ItemCode = IM_ItemWarehouse.ItemCode AND
            IM_ItemWarehouse.WarehouseCode = '000' AND
            IM_ItemWarehouse.ReorderPointQty = 0 AND
            (CI_Item.UDF_NONSTOCK = 'N' or CI_Item.UDF_NONSTOCK is Null)     
    """

    PutOnNonStock = pd.read_sql(sql,cnxn) 
    print('AA_PUTONNONSTOCK_VIWI5Q')
    if PutOnNonStock.shape[0] > 0:
        filepath = r'\\FOT00WEB\Alt Team\Qarl\Automatic VI Jobs\Maintenance\CSVs\AA_PUTONNONSTOCK_VIWI5Q.csv' 
        PutOnNonStock.to_csv(filepath, index=False, header=False)
        print(PutOnNonStock)
        
        #Auto VI .... uncomment  below to turn on....untested
        p = subprocess.Popen('Auto_PutOnNonStock_VIWI5Q.bat', cwd=r"Y:\Qarl\Automatic VI Jobs\Maintenance", shell=True)
        stdout, stderr = p.communicate()
        p.wait()
        print('Sage VI Complete!')
    else:
        print('No PutOnNonStock')     
        

    sql = """
        SELECT 
            CI_Item.ItemCode
        FROM 
            CI_Item CI_Item, 
            IM_ItemWarehouse IM_ItemWarehouse
        WHERE
            CI_Item.ItemCode = IM_ItemWarehouse.ItemCode AND
            IM_ItemWarehouse.WarehouseCode = '000' AND
            IM_ItemWarehouse.ReorderPointQty > 0 AND
            CI_Item.UDF_NONSTOCK = 'Y'
    """
    print('AA_TAKEOFFNONSTOCK_VIWI5P')
    TakeOffFlag = pd.read_sql(sql,cnxn)
    if TakeOffFlag.shape[0] > 0:
        filepath = r'\\FOT00WEB\Alt Team\Qarl\Automatic VI Jobs\Maintenance\CSVs\AA_TAKEOFFNONSTOCK_VIWI5P.csv' 
        TakeOffFlag.to_csv(filepath, index=False, header=False)
        print(TakeOffFlag)
        
        #Auto VI .... uncomment  below to turn on....untested
        p = subprocess.Popen('Auto_TakeOffNonStock_VIWI5P.bat', cwd=r"Y:\Qarl\Automatic VI Jobs\Maintenance", shell=True)
        stdout, stderr = p.communicate()
        p.wait()
        print('Sage VI Complete!')
    else:
        print('No TakeOffFlag')                  
        