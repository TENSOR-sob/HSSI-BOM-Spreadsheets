# HSSI-BOM-Spreadsheets
Generates High Steel BOM Spreadsheets
Requires detailed job.
Must have job/BILL/PRODUCT/product.csv which is generated by running highbillcsv in the BILL directory
Must have JobStandards.xlsm (or JobStandards-temp.xlsm) and Products.xlsx (or Products-temp.xlsm) in the BILL/PRODUCT directory.  These files are automatically copied from the HSSI system when the job is started.
1st run HighJobStds.py to generate the JobStandards.xlsm file.  2nd run HighJobProduct.py to generate the product BOM spreadsheets.