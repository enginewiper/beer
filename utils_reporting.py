import pandas as pd
from datetime import date, timedelta


def is_current_reporting_period(transactiondate):
    last_day_of_prev_month = date.today().replace(day=1) - timedelta(days=1)
    start_day_of_prev_month = date.today().replace(day=1) - timedelta(days=last_day_of_prev_month.day)
    return start_day_of_prev_month <= transactiondate <= last_day_of_prev_month


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

invPath = 'Inventory.xlsx'
productsPath = 'Products.xlsx'
retailerCustomersPath = 'RetailerCustomers.xlsx'
transactionsPath = 'Transactions.xlsx'

dfInv = pd.io.excel.read_excel(invPath)
dfProducts = pd.io.excel.read_excel(productsPath)
dfRetailerCustomers = pd.io.excel.read_excel(retailerCustomersPath)
dfTransactions = pd.io.excel.read_excel(transactionsPath)

# print(df)

#for row in dfInv.iterrows():
#    print(row)

#for row in dfProducts.iterrows():
#    print(row)

#for row in dfRetailerCustomers.iterrows():
#    print(row)

#for row in dfTransactions.iterrows():
    #print(row[])
 #   print(row)

for row in dfTransactions.itertuples(index=False):
    #print(is_current_reporting_period(row[dfTransactions.columns.get_loc('Date')]))
    if is_current_reporting_period(row[dfTransactions.columns.get_loc('Date')]):
        #include this transaction in the list of transactions to process for reporting
        pass
    print(row[dfTransactions.columns.get_loc('Date')]))
    print(row[dfTransactions.columns.get_loc('ProductID')])
    print(row[dfTransactions.columns.get_loc('InternalCustomerID')])
    print(row[dfTransactions.columns.get_loc('ContainerSize')])
    print(row[dfTransactions.columns.get_loc('ContainerUnits')])
    print(row[dfTransactions.columns.get_loc('OffPrem')])
