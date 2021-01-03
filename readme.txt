Updates 1/3/2021
----------------
- Fixed 235/236 form templates
- Added Status column to Inventory.xlsx
- Merging 235/236 form scripts into TABC_Forms.py
- Added function generate_product_list
- Added Current235 class

Current235 class objects:
Create a new 235 form to be filled out upon construction
Store data to be written to the 235 form
Scrape data from Transactions.xlsx, format the data to be written
Write all data to the form

generate_product_list:
creates an array of Product objects for each product in products.xlsx
each Product object stores data for on-premise sales, off-premise sales
as well as brandname, productID, Beer or malt liquor

Each product in the array has a variable xl_brand_row
xl_brand_row will be used later to put brand names in an alphabetical ordered list,
and write the data for that brandname to the right row