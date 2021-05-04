# -*- coding: utf-8 -*-
# -*- coding: utf-8 -*-
"""
Created on Tue Apr 13 10:55:31 2021

@author: nrehman
"""

import xmlrpc.client
import pandas as pd
import pandas as pd
import glob, os


url = 'https://packsmartinc-pack-smart-production-451613.dev.odoo.com'
db = 'packsmartinc-pack-smart-production-451613'


class Odoo():
    
    def __init__(self):
        self.common = xmlrpc.client.ServerProxy('{}/xmlrpc/2/common'.format(url))
        self.models = xmlrpc.client.ServerProxy('{}/xmlrpc/2/object'.format(url))

    def authenticate(self,username,password):
        self.username = username
        self.password = password
        self.uid = self.common.authenticate(db, self.username, self.password, {})
        print(self.uid)
        return self.uid
    
    def getPurchasePrice(self):
        purch_price = self.models.execute_kw(db, self.uid, self.password, 'purchase.order.line',
                        #product_type is product AND state is purchase OR release
                        #filters out service/consumables
                        'search_read', [[['product_type','=','product'],
                                         '|', ['state','=','purchase'],['state','=','release']]],
                        {'fields':['product_id','price_unit','discount','product_qty','product_uom','state','partner_id']})
        
        purch_price = pd.DataFrame(purch_price)
        return purch_price
    
    def getSalePrice(self):
        sale_price = self.models.execute_kw(db, self.uid, self.password, 'product.pricelist',
                        #product_type is product AND state is purchase OR release
                        #filters out service/consumables
                        'search_read', [[[]]],
                        {'fields':['item_ids']}) #TODO: find variable for "price" in product.pricelist
        
        sale_price = pd.DataFrame(sale_price)
        return sale_price
    

class Spares():
    def __init__(self):
        self.path = input("Enter path for module files: ") #Ex: C:/Users/nrehman/Documents/Modules
        #Read all xls files in user specified folder
        self.all_modules = glob.glob(self.path + "/*.xls")
        
        self.outputfile_name = input('Enter name for output file: ')
        self.outputfile_path = self.path + "/" + self.outputfile_name + ".xlsx"
        
    
    def generate_list(self):
                
        #Append each file contents to dataframe
        self.bom = pd.concat((pd.read_excel(m, header=7).assign(MODULE=os.path.basename(m)) for m in self.all_modules))
        
        #drop all rows containing NAs in Spare Class columns
        self.bom = self.bom.dropna(subset = ['SPARE CLASS'])
        
        #Convert to string
        self.bom['SPARE CLASS'] = self.bom['SPARE CLASS'].astype(str)
        self.bom['MODULE'] = self.bom['MODULE'].astype(str)
        
        #Strip Spaces
        self.bom['SPARE CLASS'] = self.bom['SPARE CLASS'].str.replace(' ','')
        self.bom['MODULE'] = self.bom['MODULE'].str.replace('.xls','')
        
        #Sort df by Spare Class values
        self.bom['SPARE CLASS'] = pd.Categorical(self.bom['SPARE CLASS'], ['1','2','3','9','X'])
        self.bom = self.bom.sort_values('SPARE CLASS')
        
        #Create new df without Space Class 'X' items
        self.bom_spares = self.bom[self.bom['SPARE CLASS'] != 'X']
        
        #Create new df without duplicates
        self.bom_spares_unique = self.bom_spares.drop_duplicates(subset = ['PART NUMBER'])
        
        #Iterate through each unique part in df
        for i in range(self.bom_spares_unique.shape[0]):
            #extract part number
            part = self.bom_spares_unique['PART NUMBER'].iloc[i]
            
            #filter bom_spares by "part" & output MODULE column to a list
            #provides all module occurences of "part"
            modules_list = self.bom_spares[self.bom_spares['PART NUMBER'] == part]['MODULE'].to_list()
            
            #replace MODULE column in bom_spares_unique with modules_list
            self.bom_spares_unique['MODULE'].iloc[i] = modules_list
            
           ##TODO: GET SUM OF COLUMN INSTEAD to reduce O^n##
            
            #filter bom_spares by "part" & output PROJ QTY column to a list
            total_qty = self.bom_spares[self.bom_spares['PART NUMBER'] == part]['PROJ\nQTY.'].to_list()
            
            #Calculate sum of list
            total_qty_sum = 0
            for j in total_qty:
                try:
                    total_qty_sum+=j
                except:
                    "Couldnt add qty"
            self.bom_spares_unique['PROJ\nQTY.'].iloc[i] = total_qty_sum
        
        return self.bom_spares_unique
        
    def export_list(self): 
        #Export df to excel sheet
        with pd.ExcelWriter(self.outputfile_path) as writer:
            self.bom.to_excel(writer, sheet_name='All Parts')
            self.bom_spares.to_excel(writer, sheet_name='Spares')
            self.bom_spares_unique.to_excel(writer, sheet_name='Spares Unqiue')
        
        print("\n Spare Parts List Created, " + self.outputfile_path)
        

class Db(object):
    def __init__(self,odoo,spares):
        self.df_purchprice = odoo.getPurchasePrice()
        self.df_spares = spares.generate_list()
        print(self.df_spares)
        print(self.df_purchprice)
        




odoo = Odoo()
spares = Spares()
username = input("Enter Odoo Username: ")
password = input("Enter password: ")
odoo.authenticate(username, password)
#df1= odoo.getPurchasePrice()
#odoo.getSalePrice()
        

test1 = Db(odoo, spares)    
