import pandas as pd
import numpy as np
from datetime import date, timedelta, datetime
from pandas.io import excel
from openpyxl import load_workbook
from shutil import copyfile

# Set paths for all files statically.
invPath = 'Inventory.xlsx'
productsPath = 'Products.xlsx'
retailerCustomersPath = 'RetailerCustomers.xlsx'
transactionsPath = 'Transactions.xlsx'


# is a transaction date within the previous month? boolean
def is_current_reporting_period(transaction_date):
    last_day_of_prev_month = date.today().replace(day=1) - timedelta(days=1)
    start_day_of_prev_month = date.today().replace(day=1) - timedelta(days=last_day_of_prev_month.day)
    return start_day_of_prev_month <= transaction_date <= last_day_of_prev_month


def get_previous_report_name_prefix():
    today = date.today()
    first = today.replace(day=1)
    last_month = first - timedelta(days=32)
    return last_month.strftime("%Y_%m_")


reports_rootdir = 'C:/Users/svenf/Desktop/beer/reportsdir/'
# reports_rootdir = 'C:/!/projects/PycharmProjects/beer/reportsdir/'


previous_tabc_235_filename = reports_rootdir + get_previous_report_name_prefix() + 'c-235.xlsx'
previous_tabc_236_filename = reports_rootdir + get_previous_report_name_prefix() + 'c-236.xlsx'


def new235():  # make a blank copy of a 235 form for the current month and year
    # Puts a new 235 form in complete_forms to be filled out by fill235 method
    current_month = str(datetime.now().month)
    current_year = str(datetime.now().year)
    new_filename = current_year + "_" + current_month + "_" + "c-235.xlsx"
    copyfile("form_templates/235.xlsx", new_filename)
    return new_filename


def new236():  # make a blank copy of a 236 form for the current month and year
    # Puts a new 236 form in complete_forms to be filled out by fill236 method
    current_month = str(datetime.now().month)
    current_year = str(datetime.now().year)
    new_filename = current_year + "_" + current_month + "_" + "c-235.xlsx"
    copyfile("form_templates/236.xlsx", new_filename)
    return new_filename


class Transaction:
    def __init__(self,
                 product_id,
                 individual_container_size,
                 container_units,
                 number_units,
                 is_off_premise,
                 is_retailer_sale):
        self.product_id = str(product_id)  # Product ID
        self.individual_container_size = str(individual_container_size)  # Individual Container Size
        self.container_units = str(container_units)  # oz or gallons?
        self.number_units = int(number_units)  # number of containers sold
        if container_units == 'oz':  # If the containers are in oz, convert to gallons, times x amount of containers
            self.total_gallons = number_units * (individual_container_size / 128)
        else:  # If the containers are in gallons
            self.total_gallons = number_units * individual_container_size

        self.is_off_premise = str(is_off_premise)
        self.is_retailer_sale = is_retailer_sale  # 0 or 1
        self.product_list_index = 0  # Index of the product in the product list array. TBD


class Product:
    def __init__(self, brandName, productID, isLiquor):  # Product name, product ID, is it liquor or not (0 or 1)?
        self.alphabetical_order = 0  # Used to find which row of the tax form that the product data will be written to.
        self.brandName = str(brandName)
        self.productID = str(productID)
        self.isLiquor = isLiquor  # TODO: change this to a boolean type
        self.ONP_total_gallons_sold = 0  # ON PREMISE SALES total gallons (will need to convert units to gallons before inputting)
        self.OFFP_total_gallons_sold = 0  # OFF PREMISE SALES
        self.RTL_sixth_barrel_sold = 0  # RETAILER SALES (add different container types below this line)
        self.RTL_twentytwo_oz_bottle_sold = 0
        self.RTL_total_gallons_sold = 0


def generate_transaction_list():  # Generate a list of Transaction objects with data needed for both forms.
    transactions = []
    dfTransactions = pd.io.excel.read_excel(transactionsPath)
    # dfProducts = pd.io.excel.read_excel(productsPath)

    for row in dfTransactions.itertuples(index=False):  # iterate through all rows in dataframe
        if is_current_reporting_period(row[dfTransactions.columns.get_loc('Date')]):  # Getting data into local vars
            productID = row[dfTransactions.columns.get_loc('ProductID')]  # Getting product ID into local var
            outputIndividualContainerSize = row[
                dfTransactions.columns.get_loc('ContainerSize')]  # Getting container size #
            outputContainerUnits = row[
                dfTransactions.columns.get_loc('ContainerUnits')]  # Getting units (oz or gallons?)
            outputNumberUnits = row[dfTransactions.columns.get_loc('NumUnits')]  # Getting number of containers sold
            isOffPremise = row[dfTransactions.columns.get_loc('OffPrem')]  # is it off premise or on premise
            if isOffPremise == "TRUE":
                isOffPremise = 1
            elif isOffPremise == "FALSE":
                isOffPremise = 0
            internalCustomerID = row[dfTransactions.columns.get_loc('InternalCustomerID')]  # internal customer id
            cxid = int(internalCustomerID)  # Converting internal customer ID to an int.
            is_rtl_sale = 0  # boolean, 0 or 1
            if cxid != 1:  # If internal customer ID is not 1, it's a retailer sale.
                is_rtl_sale = 1
            elif cxid == 1:
                is_rtl_sale = 0
            transactions.append(Transaction(product_id=productID,
                                            individual_container_size=outputIndividualContainerSize,
                                            container_units=outputContainerUnits,
                                            number_units=outputNumberUnits,
                                            is_off_premise=isOffPremise,
                                            is_retailer_sale=is_rtl_sale)
                                )
    return transactions


def test_transaction_list(transaction_list_name):
    test_list = transaction_list_name
    for x in range(0, len(test_list)):
        print("PRODUCT ID:" + test_list[x].product_id)
        print(str(test_list[x].number_units) + " of " + str(test_list[x].individual_container_size) + test_list[x].container_units)
        print("Off premise: " + test_list[x].is_off_premise)
        if test_list[x].is_retailer_sale == 1:
            print ("(Retailer sale)")
        print("\n")


# puts product data into an array of objects
def generate_product_list(product_book):  # Put file path / filename for Products.xlsx as the parameter.
    book = load_workbook(product_book, data_only=True)
    product_list = []  # array that objects will go in
    number_of_products = -1  # Starts at -1 so it won't count the top cell.
    x_column = book.active['A']  # for this part it doesn't actually matter which column
    for x in x_column:
        if x.value is None:
            break
        number_of_products = number_of_products + 1

    for x in range(0, number_of_products):  # Make (number_of_products) products
        d = book.active[("D" + str(x+2))].value  # find brand name for new object
        a = book.active[("A" + str(x+2))].value  # find productID for new object
        b = book.active[("B" + str(x+2))].value  # find isLiquor for new object
        is_liquor = 0
        if b == "B":
            is_liquor = 0
        elif b == "ML":
            is_liquor = 1
        product_list.append(Product(d, a, is_liquor))

    return product_list


def test_product_list(pbook):
    test_list = generate_product_list(pbook)
    for x in range(0, len(test_list)):
        print("BRAND NAME:")
        print(test_list[x].brandName)
        print(" PRODUCT ID:")
        print(test_list[x].productID)
        print(" IS LIQUOR:")
        print(test_list[x].isLiquor)
        print("\n")


class PreviousForm:  # Class that holds all the data from previous 235 or 236 form.
    def __init__(self, filename):  # Takes in previous form filename as parameter.
        self.workbook = load_workbook(filename, data_only=True)  # can be appended to object and use openpyxl functions
        self.inv_end_of_month = (self.workbook['Summary Page'])['B19'].value  # Line 1
        self.beverage_brewed_ytd = (self.workbook['Summary Page'])['C15'].value  # Line 2 YTD
        self.beverage_imported_ytd = (self.workbook['Summary Page'])['C16'].value  # Line 3 YTD
        self.beverage_taxable_ytd = (self.workbook['Summary Page'])['C23'].value  # Line 10 YTD
        self.tax_rate_per_gallon = (self.workbook['Summary Page'])['B24'].value  # Line 11
        self.sold_to_retailers_ytd = (self.workbook['Summary Page'])['C26'].value  # Line 12 YTD
        self.on_premise_consumption_ytd = (self.workbook['Summary Page'])['C27'].value  # Line 13 YTD
        self.off_premise_consumption_ytd = (self.workbook['Summary Page'])['C28'].value  # Line 14 YTD


class Current235:  # Class that holds data to be put in the new 235 form.
    def __init__(self):
        # new235 will return the filename after creating a new form
        self.filename = new235()
        # Summary Page Variables
        self.line1 = 0  # Inventory, Beginning of month
        self.line2 = 0  # Beer Manufactured
        self.line3 = 0  # Beer Imported
        self.line4 = 0  # Beer Returned from TX Distributors
        self.line5 = 0  # Total Beer Available
        self.line6 = 0  # Inventory, End of month
        self.line7 = 0  # Distributor sales
        self.line8 = 0  # Other Exemptions
        self.line9 = 0  # Total Exemptions
        self.line10 = 0  # Beer subject to taxation
        self.line12 = 0  # Total beer sold to retailers
        self.line13 = 0  # Total beer sold for on-premise consumption
        self.line14 = 0  # Total beer sold for off-premise consumption
        self.line15 = 0  # Total taxable sales
        self.tax_rate_per_gallon = (self.filename['Summary Page'])['B24'].value
        self.grosstax = 0  # Gross Tax Due
        self.less_2percent = 0
        self.less_authorized_credits = 0
        self.tax_due_state = 0


class Current236:  # Class that holds data to be put in the new 235 form.
    def __init__(self):
        # new235 will return the filename after creating a new form
        self.filename = new236()
        # Summary Page Variables
        self.line1 = 0  # Inventory, Beginning of month
        self.line2 = 0  # Ale/Malt Liquor Manufactured
        self.line3 = 0  # Ale/Malt Liquor Imported
        self.line4 = 0  # Ale/Malt Liquor Returned from TX Distributors
        self.line5 = 0  # Total Ale/Malt Liquor Available
        self.line6 = 0  # Inventory, End of month
        self.line7 = 0  # Distributor sales
        self.line8 = 0  # Other Exemptions
        self.line9 = 0  # Total Exemptions
        self.line10 = 0  # Ale/Malt Liquor subject to taxation
        self.line12 = 0  # Total Ale/Malt Liquor sold to retailers
        self.line13 = 0  # Total Ale/Malt Liquor sold for on-premise consumption
        self.line14 = 0  # Total Ale/Malt Liquor sold for off-premise consumption
        self.line15 = 0  # Total taxable sales
        self.tax_rate_per_gallon = (self.filename['Summary Page'])['B24'].value
        self.grosstax = 0  # Gross Tax Due
        self.less_2percent = 0
        self.less_authorized_credits = 0
        self.tax_due_state = 0


array_transaction_data = generate_transaction_list()
array_products = generate_product_list("Products.xlsx")
prev_235_form = PreviousForm(previous_tabc_235_filename)  # Creating object using filename as parameter.

test_transaction_list(array_transaction_data)
'''  work on this stuff later




















# 1.   Inventory, Beginning of Month  (Line 6 on Prior Monthly Report)
inv_beginning_of_month = prev_235_form.inv_end_of_month
# 2.   Ale/Malt Liquor Brewed (Gallons Bottled & Kegged)
beverage_brewed_monthly = 0.00
# TODO calculate the amount in gallons added to inventory.xlsx during reporting period, output current_sheet b15
beverage_brewed_ytd = 0.00
# TODO sum beverage_brewed_monthly with C15 from previous ytd, output current_sheet c15


# 3.   Ale/Malt Liquor Imported (Page 2, Schedule A)
# always zero for us
beverage_imported_monthly = 0.00
beverage_imported_ytd = 0.00
# TODO output current_sheet b16 and c16

# 4.   Ale/Malt Liquor Returned from TX Wholesalers
# always zero for us
beverage_returned_from_wholesalers_monthly = 0.00
# TODO output current_sheet b17

# 5.   Total Ale/Malt Liquor Available (Sum of Lines 1,2,3,4)
beverage_total_available_monthly = inv_beginning_of_month + \
                                   beverage_brewed_monthly + \
                                   beverage_imported_monthly + \
                                   beverage_returned_from_wholesalers_monthly
# TODO output to current_sheet b18

# 6.   Inventory, End of Month
inv_end_of_month = 0.00
# TODO calculate end of month inventory from inventory.xlsx
# TODO output to current_sheet b19

# 7.   Wholesaler Sales (Page 2, line 2)
# always zero for us
wholesaler_sales_monthly = 0.00
wholesaler_sales_ytd = 0.00
# TODO output to current_sheet b20 and c20

# 8.   Other Exemptions (Page 2, line 3)
# we will calculate these manually as one-offs if there ever are any
other_exemptions_monthly = 0.00
other_exemptions_ytd = 0.00
# TODO output to current_sheet b21 and c21

# 9.   Total Exemptions (Sum of Lines 6, 7 & 8)
beverage_total_exemptions_monthly = inv_end_of_month + wholesaler_sales_monthly + other_exemptions_monthly
# TODO output to current_sheet b22

# 10.  Ale/Malt Liquor Subject to Taxation (Line 5 minus Line 9)
beverage_subject_to_taxation_monthly = beverage_total_available_monthly - beverage_total_exemptions_monthly
beverage_subject_to_taxation_ytd = (prev_235_form.beverage_taxable_ytd + beverage_subject_to_taxation_monthly)

# 11. Tax Rate Per Gallon
tax_rate_per_gal = prev_235_form.tax_rate_per_gallon

# Compliance Totals
# 12.  Total Ale/Malt Liquor Sold to Retailers (Page 3)
# TODO generate Brand Summary Sold sheet inputs from transactions.xlsx
# calculate total values monthly and ytd
# output to current_sheet B26 and C26

# 13.  Total Ale/Malt Liquor Sold for On-Premise Consumption
# TODO generate total ale/malt liquor sold on-prem from transactions.xlsx
# calculate total values monthly and ytd
# output to current_sheet B27 and C27

# 14.  Total Ale/Malt Liquor Sold for Off-Premise Consumption
# TODO generate total ale/malt liquor sold off-prem from transactions.xlsx
# calculate total values monthly and ytd
# output to current_sheet B28 and C28

# 15.  Total Taxable Sales* (Sum of Lines 12,13, & 14)

# TODO calculate total taxable sales monthly
# output to current_sheet B29

gross_tax_due = beverage_subject_to_taxation_monthly * tax_rate_per_gal
# TODO output to current_sheet B32

discount_tax_due = gross_tax_due - (gross_tax_due * .02)
# TODO output to current_sheet B33

less_authorized_credits = discount_tax_due
# TODO output to current_sheet B34

tax_due_state = discount_tax_due
# TODO output to current_sheet B35


# test_product_list(productsPath)
'''