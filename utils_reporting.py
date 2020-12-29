import pandas as pd
from datetime import date, timedelta


# is a transaction date within the previous month? boolean
def is_current_reporting_period(transactiondate):
    last_day_of_prev_month = date.today().replace(day=1) - timedelta(days=1)
    start_day_of_prev_month = date.today().replace(day=1) - timedelta(days=last_day_of_prev_month.day)
    return start_day_of_prev_month <= transactiondate <= last_day_of_prev_month


# return TABC 111 brewing, LLC license number based on comptroller beverage class
def get_111_license(comptroller_beverage_class):
    if comptroller_beverage_class == 'ML':
        # return ale TABC license number
        return 'B1050482'
    else:
        return 'BA1050483'


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

# Comptroller

'''
1. Seller's TABC Permit or License Numbers
2. Retailer's/Purchaser's TABC Permit or License Number
3. Retailer's/Purchaser's Tax Identification Number	
4. Retailer's/Purchaser's Name	
5. Retailer's Street Address	
6. Retailer's City	
7. Retailer’s State	
8. Retailer's Five Digit Zip Code	
9. Beverage Class	
10. Universal Product Code	
11. Brand Name	
12. Individual Container Size	
13. Number of Containers	
14. Selling Price


Data Layout – 1st field/column
 Seller's TABC Permit or License Number
 The following permit and license holders must file:
 Wholesalers (W, X and LX)
 Distributors (BB, BD and BC)
 Wineries (G)
 Package stores holding local distributor’s permits (P W/ LP); and
 Certain brewers and beer manufactures (B and BA)
 Examples:
 BC123456; BD1234567; P123456; & P1234567
 Each line item on the report must identify the specific TABC permit or license
actually making the sale.
 Consolidation of sales under a single TABC permit or license is not allowed.
 A single Comptroller-issued taxpayer number may be associated with multiple TABC
permits or licenses but an individual TABC permit or license may only be associated
with a single taxpayer number.
 This is a mandatory field



Data Layout – 2nd field/column
 Retailer's TABC Permit or License Number
 A retailer is the customer or business that purchased alcoholic
beverages from you, the seller
 Examples: BG123456; BQ1234567; N123456; & N1234567
 Each line item on this report must identify the specific and
correct TABC permit or license to which the sale was made.
 This is a mandatory field



Data Layout – 3rd field/column
 Retailer's Tax Identification Number
 The 11-digit tax identification number assigned by the
Comptroller
 Must begin with a 3 or 1
 Do not include dashes
 Examples:
 311111111111
 177777777777
 This is a mandatory field


Data Layout – 4th field/column
 Retailer's TABC Trade Name:
 This is the TABC trade name as it appears on the retailer’s
permit or license
 Examples:
 Cold Creek Beer And Wine
 Hill country Liquor Store
 Happy Trails Store
 This is a mandatory field


Data Layout – 5th field/column
 Retailer's Street Address
 This is the retailer’s physical address (street number and street
name)
 Examples:
 1800 North Congress Ave
 6915 Main St.
 1628 South Pecan Lane
 This is a mandatory field


Data Layout – 6th field/column
 Retailer's City:
 This is the name of the city where the retailer is located
 Examples:
 Houston
 Dallas
 Odessa
 Bastrop
 This is a mandatory field



Data Layout – 7th field/column
 Retailer's State:
 This is the retailer’s 2-character state code
(TX)
 This is a mandatory field



Data Layout – 8th field/column
 Retailer's Five Digit Zip Code
 Examples:
 78701

 76101

 77702

 79902
 Do NOT use Zip+4 format
 This is a mandatory field



Data Layout – 9th field/column
 Beverage Class
 Examples:
 DS = Distilled Spirits
 W = Wine
 B = Beer
 ML = Malt Liquor and Ale
 Ale is considered malt liquor and should be reported as
“ML.”
 Each line item can only have 1 class of beverage designation.

 This is a mandatory field 



 Universal Product Code
 The manufacturer’s UPC for each line item; this is not a SKU
code or other internal code
 Each line item will have only 1 UPC, usually 12 digits but not to
exceed 18 digits
 If no UPC exists, enter a “0”
 Examples:



Data Layout – 11th field/column
 Brand Name:
 This is the complete and specific brand name of each product
you sold
 Examples:
 Miller
 Budweiser
 Bud Lite
 Shiner Oktoberfest
 Jack Daniels Black
 Johnnie Walker Red
 This is a mandatory field



Data Layout – 12th field/column
 Individual Container Size:
 This is the individual bottle, can or container size
 Multi-unit packages or case packs must reflect the size of the individual unit
 Do not use “keg,” “case” or other generic size description
 Report size as follows:
 Distilled spirits and wine containers sizes less than 1 liter in milliliters:
 750 ml 375 ml, 500 ml
 Distilled spirits containers 1 liter and greater in liters:
 1.0L, 1.75L, 1.5L
 Packaged beer and malt liquor in ounces:
 12oz, 16oz, 7oz
 Draft beer in gallons:
 15.5G, 7.25G
 Imported draft beer or malt liquor may be reported in either gallons or liters
 This is a mandatory field 



 Number of Containers:
 This is the number of individual bottles, cans or containers for
each line item
 Multi-unit packages, such as cases, must be broken down into
the number of individual bottles or cans
 To report a credit, enter a negative number (-12; -7)
 Do NOT include any spaces, decimal points, commas, etc.

 This is a mandatory field



Data Layout – 14th field/column
 Net Selling Price:
 This is the total sales amount rounded to the nearest dollar
charged to the customer of each line item on this report and
should include any applicable discounts
 To report a credit, enter a negative number (-7; -12)
 Do not include dollar signs, spaces, decimal, commas, etc.

 This is a mandatory field
'''

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

# print(df)

# for row in dfInv.iterrows():
#    print(row)

# for row in dfProducts.iterrows():
#    print(row)

# for row in dfRetailerCustomers.iterrows():
#    print(row)

# for row in dfTransactions.iterrows():
# print(row[])
#   print(row)

transactions = []
# figure out which transactions were in the previous month
for row in dfTransactions.itertuples(index=False):
    # print(is_current_reporting_period(row[dfTransactions.columns.get_loc('Date')]))
    if is_current_reporting_period(row[dfTransactions.columns.get_loc('Date')]):
        # include this transaction in the list of transactions to process for reporting
        # print('dump transaction info:')
        # print(row[dfTransactions.columns.get_loc('Date')])
        # print(row[dfTransactions.columns.get_loc('ProductID')])
        # print(row[dfTransactions.columns.get_loc('InternalCustomerID')])
        # print(row[dfTransactions.columns.get_loc('ContainerSize')])
        # print(row[dfTransactions.columns.get_loc('ContainerUnits')])
        # print(row[dfTransactions.columns.get_loc('OffPrem')])
        # print(row[dfTransactions.columns.get_loc('UnitPrice')])
        # print(row[dfTransactions.columns.get_loc('NumUnits')])

        # get information about retailer customer based on internalCustomerID,
        # and convert df to dict
        internalCustomerID = row[dfTransactions.columns.get_loc('InternalCustomerID')]
        retailerCustomer = (dfRetailerCustomers.loc[dfRetailerCustomers['InternalCustomerID']
                                                    == internalCustomerID]).to_dict('list')
        # # print(dfRetailerCustomers)
        # print('dump retailer info:')
        # # print(retailerCustomer)
        # print(retailerCustomer['InternalCustomerID'][0])
        # print(retailerCustomer['TABCPermitNumber'][0])
        # print(retailerCustomer['TaxID'][0])
        # print(retailerCustomer['RetailerName'][0])
        # print(retailerCustomer['RetailerStreetAddr'][0])
        # print(retailerCustomer['RetailerCity'][0])
        # print(retailerCustomer['RetailerState'][0])
        # print(retailerCustomer['RetailerFiveDigitZip'][0])

        # get information about the product related to this transaction
        productID = row[dfTransactions.columns.get_loc('ProductID')]
        product = (dfProducts.loc[dfProducts['ProductID'] == productID]).to_dict('list')

        # print(product)

        # print('dump product info:')
        # print(product['ProductID'][0])
        # print(product['ComptrollerBeverageClass'][0])
        # print(product['UPC'][0])
        # print(product['BrandName'][0])

        # comptroller reporting
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
        outputDisplayIndividualContainerSize = str(row[dfTransactions.columns.get_loc('ContainerSize')]) + \
                                               outputContainerUnits
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
# transactions is a list of all transaction objects for the reporting period.
# sum the total number of containers and extended price for each package type to each retailer.
# for transaction in transactions:
#     print(transaction.to_dict())

dfComptroller = pd.DataFrame([transaction.to_dict() for transaction in transactions])
# dfComptroller = dfComptroller.round({'selling_price': 0})
# dfComptroller = dfComptroller.astype({'selling_price': int})
#filter to remove transactions from 111 brewing
dfComptroller = dfComptroller.loc[~dfComptroller['retailer_name'].str.contains('111 Brewing, LLC')]
# Total = dfComptroller.loc[dfComptroller['product_brandname'].str.contains('Say When Local Motive IPA')]
#dfComptrollergrouped = dfComptroller.groupby('product_brandname')

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
#, 'retailer_name': 'first'}
#df_new = dfComptroller.groupby(dfComptroller['retailer_name', 'product_brandname']).aggregate(aggregation_functions)
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
#df_new = df_new.join(dfComptroller['retailer_tabc_permit_number'])
#df_new.merge(dfComptroller.set_index('retailer_name'), on='retailer_name')

with pd.option_context('display.max_rows', None, 'display.max_columns', None):
#    pass
    print(df_new)
#    print('print total')
#    print(Total)

#    print(dfComptroller)
