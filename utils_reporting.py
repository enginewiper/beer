import pandas as pd
import numpy as np
from datetime import date, timedelta
from pandas.io import excel


# is a transaction date within the previous month? boolean
def is_current_reporting_period(transaction_date):
    last_day_of_prev_month = date.today().replace(day=1) - timedelta(days=1)
    start_day_of_prev_month = date.today().replace(day=1) - timedelta(days=last_day_of_prev_month.day)
    return start_day_of_prev_month <= transaction_date <= last_day_of_prev_month


# return TABC 111 brewing, LLC license number based on comptroller beverage class
def get_111_license(comptroller_beverage_class):
    if comptroller_beverage_class == 'ML':
        # return ale TABC license number
        return 'B1050482'
    else:
        return 'BA1050483'


# return the prefix for the output filename (year_lastmonth_)
def get_output_file_name_prefix():
    today = date.today()
    first = today.replace(day=1)
    last_month = first - timedelta(days=1)
    return last_month.strftime("%Y_%m_")


# return the prefix for the previous month's report name (year_month before last_)
def get_previous_report_name_prefix():
    today = date.today()
    first = today.replace(day=1)
    last_month = first - timedelta(days=32)
    return last_month.strftime("%Y_%m_")


class Transaction:
    def __init__(self,
                 seller_tabc_permit_number,
                 retailer_tabc_permit_number,
                 retailer_tx_taxpayer_number,
                 retailer_name,
                 retailer_street_addr,
                 retailer_city,
                 retailer_state,
                 retailer_zip,
                 comptroller_beverage_class,
                 product_upc,
                 product_brandname,
                 individual_container_size,
                 container_units,
                 number_units,
                 selling_price):
        # comptroller reporting related attributes
        # 1. Seller's TABC Permit or License Numbers
        self.seller_tabc_permit_number = str(seller_tabc_permit_number)
        # 2. Retailer's/Purchaser's TABC Permit or License Number
        self.retailer_tabc_permit_number = str(retailer_tabc_permit_number)
        # 3. Retailer's/Purchaser's Tax Identification Number
        self.retailer_tx_taxpayer_number = str(retailer_tx_taxpayer_number)
        # 4. Retailer's/Purchaser's Name
        self.retailer_name = str(retailer_name)
        # 5. Retailer's Street Address
        self.retailer_street_addr = str(retailer_street_addr)
        # 6. Retailer's City
        self.retailer_city = str(retailer_city)
        # 7. Retailer’s State
        self.retailer_state = str(retailer_state)
        # 8. Retailer's Five Digit Zip Code
        self.retailer_zip = str(retailer_zip)
        # 9. Beverage Class
        self.comptroller_beverage_class = str(comptroller_beverage_class)
        # 10. Universal Product Code
        self.product_upc = str(product_upc)
        # 11. Brand Name
        self.product_brandname = str(product_brandname)
        # 12. Individual Container Size, string concat with units (e.g. 32oz, 5.12G)
        self.individual_container_size = str(individual_container_size)
        # container units
        self.container_units = str(container_units)
        # display container units should be formatted as 32oz, or 5.12G
        if container_units == 'oz':
            self.display_container_units = str(int(individual_container_size)) + container_units
        else:
            self.display_container_units = str(individual_container_size) + container_units
        # 13. Number of Containers
        self.number_units = int(number_units)
        # 14. Selling Price
        self.selling_price = float(selling_price)

    def to_dict(self):
        return {
            # 1. Seller's TABC Permit or License Numbers
            'seller_tabc_permit_number': self.seller_tabc_permit_number,
            'retailer_tabc_permit_number': self.retailer_tabc_permit_number,
            'retailer_tx_taxpayer_number': self.retailer_tx_taxpayer_number,
            'retailer_name': self.retailer_name,
            'retailer_street_addr': self.retailer_street_addr,
            'retailer_city': self.retailer_city,
            'retailer_state': self.retailer_state,
            'retailer_zip': self.retailer_zip,
            'comptroller_beverage_class': self.comptroller_beverage_class,
            'product_upc': self.product_upc,
            'product_brandname': self.product_brandname,
            'individual_container_size': self.individual_container_size,
            'container_units': self.container_units,
            'display_container_units': self.display_container_units,
            'number_units': self.number_units,
            'selling_price': self.selling_price
        }


# TABC
# Malt Liquor
'''
1.   Inventory, Beginning of Month  (Line 6 on Prior Monthly Report)
2.   Ale/Malt Liquor Brewed (Gallons Bottled & Kegged)
3.   Ale/Malt Liquor Imported (Page 2, Schedule A)
4.   Ale/Malt Liquor Returned from TX Wholesalers
5.   Total Ale/Malt Liquor Available (Sum of Lines 1,2,3,4)
6.   Inventory, End of Month
7.   Wholesaler Sales (Page 2, line 2)
8.   Other Exemptions (Page 2, line 3)
9.   Total Exemptions (Sum of Lines 6, 7 & 8)
10.  Ale/Malt Liquor Subject to Taxation (Line 5 minus Line 9)
11. Tax Rate Per Gallon	$0.198
Compliance Totals
12.  Total Ale/Malt Liquor Sold to Retailers (Page 3) - generate summary of
13.  Total Ale/Malt Liquor Sold for On-Premise Consumption
14.  Total Ale/Malt Liquor Sold for Off-Premise Consumption
15.  Total Taxable Sales* (Sum of Lines 12,13, & 14)
*The sum of lines 12, 13, and 14 (Line 15) should be less than or equal to Line 10.

GROSS TAX DUE (Line 10 x 11)
LESS 2% (If payment received by due date)
LESS AUTHORIZED CREDITS (Attach TABC letter)
TAX DUE STATE
'''

# Beer
'''
1.   Inventory, Beginning of Month  (Line 6 on Prior Monthly Report)
2.   Beer Manufactured (Gallons Bottled & Kegged)
3.   Beer Imported (Page 2, Schedule A)
4.   Beer Returned from TX Distributors 		
5.   Total Beer Available (Sum of Lines 1,2,3,4)
6.   Inventory, End of Month
7.   Distributor Sales   (Page 2, line 2)		
8.   Other Exemptions (Page 2, line 3)		
9.   Total Exemptions  (Sum of Lines 6, 7 & 8)	
10.  Beer Subject to Taxation (Line 5 minus Line 9)		
11. Tax Rate Per Gallon	$0.193548	
Required Compliance Totals		
12.  Total Beer Sold to Retailers (Page 3)		
13.  Total Beer Sold for On-Premise Consumption 		
14.  Total Beer Sold for Off-Premise Consumption		
15.  Total Taxable Sales* (Sum of Lines 12,13, & 14)	
*The sum of lines 12, 13, and 14 (Line 15) should be less than or equal to Line 10. 		

GROSS TAX DUE (Line 10 x 11)
LESS 2% (If payment received by due date)   	
LESS AUTHORIZED CREDITS (Attach TABC letter)		
TAX DUE STATE
'''

# Comptroller reporting section

# set paths
invPath = 'Inventory.xlsx'
productsPath = 'Products.xlsx'
retailerCustomersPath = 'RetailerCustomers.xlsx'
transactionsPath = 'Transactions.xlsx'

# load excel sheets as pandas data frames
dfInv = pd.io.excel.read_excel(invPath)
dfProducts = pd.io.excel.read_excel(productsPath)
dfRetailerCustomers = pd.io.excel.read_excel(retailerCustomersPath)
dfTransactions = pd.io.excel.read_excel(transactionsPath)
transactions = []
# figure out which transactions were in the previous month
for row in dfTransactions.itertuples(index=False):
    if is_current_reporting_period(row[dfTransactions.columns.get_loc('Date')]):
        internalCustomerID = row[dfTransactions.columns.get_loc('InternalCustomerID')]
        retailerCustomer = (dfRetailerCustomers.loc[dfRetailerCustomers['InternalCustomerID']
                                                    == internalCustomerID]).to_dict('list')
        productID = row[dfTransactions.columns.get_loc('ProductID')]
        product = (dfProducts.loc[dfProducts['ProductID'] == productID]).to_dict('list')
        # format outputs based on comptroller spec above
        # 1. Seller's TABC Permit or License Numbers
        outputSellerTABCPermitNumber = get_111_license(product['ComptrollerBeverageClass'][0])
        # 2. Retailer's/Purchaser's TABC Permit or License Number
        outputRetailerTABCPermitNumber = retailerCustomer['TABCPermitNumber'][0]
        # 3. Retailer's/Purchaser's Tax Identification Number
        outputRetailerTXTaxpayerNumber = retailerCustomer['TaxID'][0]
        # 4. Retailer's/Purchaser's Name
        outputRetailerName = retailerCustomer['RetailerName'][0]
        # 5. Retailer's Street Address
        outputRetailerStreetAddr = retailerCustomer['RetailerStreetAddr'][0]
        # 6. Retailer's City
        outputRetailerCity = retailerCustomer['RetailerCity'][0]
        # 7. Retailer’s State
        outputRetailerState = retailerCustomer['RetailerState'][0]
        # 8. Retailer's Five Digit Zip Code
        outputRetailerZip = retailerCustomer['RetailerFiveDigitZip'][0]
        # 9. Beverage Class
        outputBeverageClass = product['ComptrollerBeverageClass'][0]
        # 10. Universal Product Code
        outputUPC = product['UPC'][0]
        # 11. Brand Name
        outputBrandName = product['BrandName'][0]
        # 12. Individual Container Size
        outputContainerUnits = row[dfTransactions.columns.get_loc('ContainerUnits')]
        outputIndividualContainerSize = row[dfTransactions.columns.get_loc('ContainerSize')]
        outputDisplayIndividualContainerSize = \
            str(row[dfTransactions.columns.get_loc('ContainerSize')]) + outputContainerUnits
        # 13. Number of Containers
        outputNumberUnits = row[dfTransactions.columns.get_loc('NumUnits')]
        # 14. Selling Price
        outputSellingPrice = row[dfTransactions.columns.get_loc('UnitPrice')]

        transactions.append(Transaction(seller_tabc_permit_number=outputSellerTABCPermitNumber,
                                        retailer_tabc_permit_number=outputRetailerTABCPermitNumber,
                                        retailer_tx_taxpayer_number=outputRetailerTXTaxpayerNumber,
                                        retailer_name=outputRetailerName,
                                        retailer_street_addr=outputRetailerStreetAddr,
                                        retailer_city=outputRetailerCity,
                                        retailer_state=outputRetailerState,
                                        retailer_zip=outputRetailerZip,
                                        comptroller_beverage_class=outputBeverageClass,
                                        product_upc=outputUPC,
                                        product_brandname=outputBrandName,
                                        individual_container_size=outputIndividualContainerSize,
                                        container_units=outputContainerUnits,
                                        number_units=outputNumberUnits,
                                        selling_price=outputSellingPrice)
                            )
dfComptroller = pd.DataFrame([transaction.to_dict() for transaction in transactions])
# remove 111 brewing transactions
dfComptroller = dfComptroller.loc[~dfComptroller['retailer_name'].str.contains('111 Brewing, LLC')]
# summarize comptroller by retailer, then by brand, and aggregate total sales price and number of containers
aggregation_functions = {'seller_tabc_permit_number': 'first',
                         'retailer_tabc_permit_number': 'first',
                         'retailer_tx_taxpayer_number': 'first',
                         'retailer_street_addr': 'first',
                         'retailer_city': 'first',
                         'retailer_state': 'first',
                         'retailer_zip': 'first',
                         'comptroller_beverage_class': 'first',
                         'product_upc': 'first',
                         'individual_container_size': 'first',
                         'container_units': 'first',
                         'selling_price': 'sum',
                         'number_units': 'sum'}
df_new = dfComptroller.groupby(['retailer_name', 'product_brandname']).agg(aggregation_functions)
df_new.columns = ['seller_tabc_permit_number',
                  'retailer_tabc_permit_number',
                  'retailer_tx_taxpayer_number',
                  'retailer_street_addr',
                  'retailer_city',
                  'retailer_state',
                  'retailer_zip',
                  'comptroller_beverage_class',
                  'product_upc',
                  'individual_container_size',
                  'container_units',
                  'total_price',
                  'total_units']
df_new = df_new.reset_index()
df_new['individual_container'] = df_new['individual_container_size'] + df_new['container_units']
# define column ordering
cols = ['seller_tabc_permit_number',
        'retailer_tabc_permit_number',
        'retailer_tx_taxpayer_number',
        'retailer_name',
        'retailer_street_addr',
        'retailer_city',
        'retailer_state',
        'retailer_zip',
        'comptroller_beverage_class',
        'product_upc',
        'product_brandname',
        'individual_container',
        'total_units',
        'total_price'
        ]
# remove temporary columns
df_new = df_new.drop('container_units', 1)
df_new = df_new.drop('individual_container_size', 1)
# reorder columns
df_new = df_new[cols]
# round total price to nearest dollar and format without decimals
df_new = df_new.round({'total_price': 0})
df_new['total_price'] = df_new['total_price'].apply(np.int64)
# output to csv
comptrollerOutputFileName = 'complete_forms/' + get_output_file_name_prefix() + 'comptroller.csv'
df_new.to_csv(comptrollerOutputFileName, index=False, header=False)
# with pd.option_context('display.max_rows', None, 'display.max_columns', None):
#    print(df_new)
