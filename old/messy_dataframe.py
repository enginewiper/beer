# load excel sheets as pandas data frames
dfInv = pd.io.excel.read_excel(invPath)  # load Inventory
dfProducts = pd.io.excel.read_excel(productsPath)
dfRetailerCustomers = pd.io.excel.read_excel(retailerCustomersPath)
dfTransactions = pd.io.excel.read_excel(transactionsPath)
transactions = []

# figure out which transactions were in the previous month
for row in dfTransactions.itertuples(index=False):  # for each row in transactions.xlsx
    if is_current_reporting_period(
            row[dfTransactions.columns.get_loc('Date')]):  # if the row is in the current reporting period
        internalCustomerID = row[
            dfTransactions.columns.get_loc('InternalCustomerID')]  # get the internal customer ID
        retailerCustomer = (dfRetailerCustomers.loc[dfRetailerCustomers['InternalCustomerID']
                                                    == internalCustomerID]).to_dict('list')
        productID = row[dfTransactions.columns.get_loc('ProductID')]  # get the product ID
        product = (dfProducts.loc[dfProducts['ProductID'] == productID]).to_dict(
            'list')  # get product from products list
        outputRetailerName = retailerCustomer['RetailerName'][0]
        # 10. Universal Product Code
        outputUPC = product['UPC'][0]
        # 11. Brand Name
        outputBrandName = product['BrandName'][0]
        # 12. Individual Container Size
        outputContainerUnits = row[dfTransactions.columns.get_loc('ContainerUnits')]
        outputIndividualContainerSize = row[dfTransactions.columns.get_loc('ContainerSize')]

        outputNumberUnits = row[dfTransactions.columns.get_loc('NumUnits')]

        outputOffPremise = row[dfTransactions.columns.get_loc('OffPrem')]

        outputRetailerSale = 1

        if ((outputRetailerName) == ('111 Brewing, LLC')):
            outputRetailerSale = 0

        transactions.append(Transaction(product_upc=outputUPC,
                                        product_brandname=outputBrandName,
                                        individual_container_size=outputIndividualContainerSize,
                                        container_units=outputContainerUnits,
                                        number_units=outputNumberUnits,
                                        is_off_premise=outputOffPremise,
                                        is_retailer_sale=outputRetailerSale)
                            )
dfAllSales = pd.DataFrame([transaction.to_dict() for transaction in transactions])

with pd.option_context('display.max_rows', None, 'display.max_columns', None):
    print(dfAllSales)
