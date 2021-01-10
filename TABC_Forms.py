import pandas as pd
from datetime import date, timedelta, datetime
from pandas.io import excel
from openpyxl import load_workbook
from shutil import copyfile

# Set paths for all files statically.
invPath = 'Inventory.xlsx'
productsPath = 'Products.xlsx'
retailerCustomersPath = 'RetailerCustomers.xlsx'
transactionsPath = 'Transactions.xlsx'


def get_previous_report_name_prefix():
    today = date.today()
    first = today.replace(day=1)
    last_month = first - timedelta(days=32)
    return last_month.strftime("%Y_%m_")


reports_rootdir = 'C:/Users/svenf/Desktop/beer/reportsdir/'
# reports_rootdir = 'C:/!/projects/PycharmProjects/beer/reportsdir/'


previous_tabc_235_filename = reports_rootdir + get_previous_report_name_prefix() + 'c-235.xlsx'
previous_tabc_236_filename = reports_rootdir + get_previous_report_name_prefix() + 'c-236.xlsx'


# is a transaction date within the previous month? boolean
def is_current_reporting_period(transaction_date):
    last_day_of_prev_month = date.today().replace(day=1) - timedelta(days=1)
    start_day_of_prev_month = date.today().replace(day=1) - timedelta(days=last_day_of_prev_month.day)
    return start_day_of_prev_month <= transaction_date <= last_day_of_prev_month


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


class Product:
    def __init__(self, brand_name, product_id, is_liquor):  # Product name, product ID, is it liquor or not (0 or 1)?
        self.alphabetical_order = 0  # Used to find which row of the tax form that the product data will be written to.
        self.brandName = str(brand_name)
        self.productID = str(product_id)
        self.isLiquor = is_liquor  # TODO: change this to a boolean type
        self.ONP_total_gallons_sold = 0  # ON PREMISE SALES total gallons
        self.OFFP_total_gallons_sold = 0  # OFF PREMISE SALES
        self.RTL_sixth_barrel_sold = 0  # RETAILER SALES (add different container types below this line)
        self.RTL_twentytwo_oz_bottle_sold = 0
        self.RTL_total_gallons_sold = 0


def generate_transaction_list():  # Generate a list of Transaction objects with data needed for both forms.
    transactions = []
    df_transactions = pd.io.excel.read_excel(transactionsPath)
    # dfProducts = pd.io.excel.read_excel(productsPath)

    for row in df_transactions.itertuples(index=False):  # iterate through all rows in dataframe
        if is_current_reporting_period(row[df_transactions.columns.get_loc('Date')]):  # Getting data into local vars
            output_product_id = row[df_transactions.columns.get_loc('ProductID')]  # Getting product ID into local var
            output_individual_container_size = row[
                df_transactions.columns.get_loc('ContainerSize')]  # Getting container size #
            output_container_units = row[
                df_transactions.columns.get_loc('ContainerUnits')]  # Getting units (oz or gallons?)
            output_number_units = row[df_transactions.columns.get_loc('NumUnits')]  # Getting number of containers sold
            is_off_premise = row[df_transactions.columns.get_loc('OffPrem')]  # is it off premise or on premise
            if is_off_premise == "TRUE":
                is_off_premise = 1
            elif is_off_premise == "FALSE":
                is_off_premise = 0
            internal_customer_id = row[df_transactions.columns.get_loc('InternalCustomerID')]  # internal customer id
            cxid = int(internal_customer_id)  # Converting internal customer ID to an int.
            is_rtl_sale = 0  # boolean, 0 or 1
            if cxid != 1:  # If internal customer ID is not 1, it's a retailer sale.
                is_rtl_sale = 1
            elif cxid == 1:
                is_rtl_sale = 0
            transactions.append(Transaction(product_id=output_product_id,
                                            individual_container_size=output_individual_container_size,
                                            container_units=output_container_units,
                                            number_units=output_number_units,
                                            is_off_premise=is_off_premise,
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
            print("(Retailer sale)")
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


def test_product_list(given_array):
    test_list = given_array
    for x in range(0, len(test_list)):
        print("BRAND NAME:")
        print(test_list[x].brandName)
        print(" PRODUCT ID:")
        print(test_list[x].productID)
        print(" IS LIQUOR:")
        print(test_list[x].isLiquor)
        print("Total Gallons Sold to Retailers:" + str(test_list[x].RTL_total_gallons_sold))
        print("Sixth Barrels sold to Retailers:" + str(test_list[x].RTL_sixth_barrel_sold))
        print("22oz Bottles sold to Retailers:" + str(test_list[x].RTL_twentytwo_oz_bottle_sold))
        print("Total gallons sold On premise consumption:" + str(test_list[x].ONP_total_gallons_sold))
        print("Total gallons sold Off premise consumption:" + str(test_list[x].OFFP_total_gallons_sold))
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


array_transaction_data = generate_transaction_list()  # List of transactions in the current reporting period.
array_products = generate_product_list("Products.xlsx")  # List of products
product_id_dict = {p.productID: p for p in array_products}  # Creates a dictionary of product IDs that point to objects.
prev_235_form = PreviousForm(previous_tabc_235_filename)  # Creating object using filename as parameter.
prev_236_form = PreviousForm(previous_tabc_236_filename)
current_235_form = Current235()
current_236_form = Current236()

for x in range(0, len(array_transaction_data)):  # For the number of transactions in the array
    # Intermediary variables for total amounts sold in transaction
    total_gallons = array_transaction_data[x].total_gallons  # total gallons for this transaction
    total_22oz_bottles = 0
    total_sixth_barrels = 0
    # If the transaction is 22oz bottles
    if array_transaction_data[x].individual_container_size == "22" and array_transaction_data[x].container_units == "oz":
        total_22oz_bottles = array_transaction_data[x].number_units

    # If the transaction is sixth barrels
    if array_transaction_data[x].individual_container_size == "5.12" and array_transaction_data[x].container_units == "G":
        total_sixth_barrels = array_transaction_data[x].number_units

    # If the transaction is a retailer sale
    if array_transaction_data[x].is_retailer_sale == 1:
        # Increase product object data by amounts sold
        product_id_dict[array_transaction_data[x].product_id].RTL_sixth_barrel_sold += total_sixth_barrels
        product_id_dict[array_transaction_data[x].product_id].RTL_twentytwo_oz_bottle_sold += total_22oz_bottles
        product_id_dict[array_transaction_data[x].product_id].RTL_total_gallons_sold += total_gallons

    # If the transaction is not a retailer sale
    if array_transaction_data[x].is_retailer_sale == 0:
        if array_transaction_data[x].is_off_premise == "True":
            product_id_dict[array_transaction_data[x].product_id].OFFP_total_gallons_sold += total_gallons

        if array_transaction_data[x].is_off_premise == "False":
            product_id_dict[array_transaction_data[x].product_id].ONP_total_gallons_sold += total_gallons


# Write all data to current 235 form object

# Write all data to current 236 form object

# Output data to 235 form page 1

# Output data to 235 form page 3

# Output data to 236 form page 1

# Output data to 236 form page 3
