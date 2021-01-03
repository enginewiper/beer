import pandas as pd
import numpy as np
from datetime import date, timedelta, datetime
from pandas.io import excel
from openpyxl import load_workbook
from shutil import copyfile


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


class Product:
    def __init__(self, d, a, b):  # Product name, product ID, is it liquor or not (0 or 1)?
        self.xl_brand_row = 0  # Used to find which row of the tax form that the product data will be written to.
        self.brandName = d
        self.productID = a
        self.isLiquor = b
        self.half_barrel_sold_ONP = 0  # ON PREMISE SALES
        self.fourth_barrel_sold_ONP = 0
        self.sixth_barrel_sold_ONP = 0
        self.twentyfour_twelve_sold_ONP = 0
        self.twentyfour_sixteen_sold_ONP = 0
        self.twelve_thirtytwo_sold_ONP = 0
        self.sixtyfour_oz_sold_ONP = 0
        self.thirtytwo_oz_sold_ONP = 0
        self.sixteen_oz_sold_ONP = 0
        self.half_barrel_sold_OFFP = 0  # OFF PREMISE SALES
        self.fourth_barrel_sold_OFFP = 0
        self.sixth_barrel_sold_OFFP = 0
        self.twentyfour_twelve_sold_OFFP = 0
        self.twentyfour_sixteen_sold_OFFP = 0
        self.twelve_thirtytwo_sold_OFFP = 0
        self.sixtyfour_oz_sold_OFFP = 0
        self.thirtytwo_oz_sold_OFFP = 0
        self.sixteen_oz_sold_OFFP = 0
        self.total_sold_in_gallons = 0


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
    for x in range(0, 3):
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


prev_235_form = PreviousForm(previous_tabc_235_filename)  # Creating object using filename as parameter.

# print(previous_tabc_235_workbook.sheetnames)
# >>> ['Summary Page', 'Schedules', 'Brand Summary Sold', 'Supplemental Schedule']
print((prev_235_form.workbook['Summary Page'])['B18'].value)


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


test_product_list("Products.xlsx")  # IT WORKS!
