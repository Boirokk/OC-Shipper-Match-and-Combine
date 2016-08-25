# Created by Boirokk 2016-08-19
# This progam searches for matching OC and Shipper excel files. Creates a Copy of the OC and copies the shipper excel to a
# worksheet (tab on bottom) of OC and adds a note worksheet (tab on bottom) to oc
# This program does not copy the formatting from the shipper packer. (Please select all and use the painter to transfer the formating from
# an existing shipper packer.)
import openpyxl
from xlrd import open_workbook
from os.path import join
from xlutils.copy import copy
import os, re, shutil
# import re is for Regular Expressions


# CONSTANTS
OC_FILE_PATH = r"C:\Users\Chad\Desktop\OC SP Merge\OC" # Path where oc excel files exist
SHIPPER_FILE_PATH = r"C:\Users\Chad\Desktop\OC SP Merge\SP" # Path where shipper excel files exist
SAVE_FOLDER = r"C:\Users\Chad\Desktop\New folder" # Path to save new combined files
MATCH_FOUND_LIST = []



    
#search for oc#'s in shipper packer folder
for shipfile in os.listdir(SHIPPER_FILE_PATH):
    try:
        ship_num = re.findall(r'^\d{5}\w|^\d{5}', shipfile) # Use regular expression to find 5 digits in the front of the string
        ship_num = ''.join(ship_num) # Convert list to string
        match_shipfile = SHIPPER_FILE_PATH + os.sep + shipfile # Get file path and save to variable

        #search for oc#'s in OC folder
        for ocfile in os.listdir(OC_FILE_PATH):
            oc_num = re.findall(r'^OC\d{5}\w|^OC\d{5}', ocfile) # Use regular expression to find 5 digits in the front of the string
            oc_num = ''.join(oc_num) # Convert list to string
            oc_num = oc_num[2:] # Remove the OC characters from string
            match_ocfile = OC_FILE_PATH + os.sep + ocfile # Get file path and save to variable
            
            print(ship_num) # Print shipper results to screen for visual
            print(oc_num) # Print OC results to screen for visual
            
            # If oc and shipper numbers match do this
            if oc_num == ship_num:

                if '.xls' in match_ocfile:
                    rb = open_workbook(match_ocfile, formatting_info=True, on_demand=True) # Open OC excel with formatting
                    rb2 = open_workbook(match_shipfile) # Open the shipper excel
                sheet = rb2.sheet_by_index(0) # Get the first sheet in the shipper excel
                wb = copy(rb) # Copy OC excel
                sheet_shipper_packer = wb.add_sheet('Yachts') # Create a new worksheet (tab on bottom) named Yachts
                sheet_notes = wb.add_sheet('Notes') # Create a new worksheet (tab on bottom) named Notes
                
                
                # Copy content of shipper excel work sheet
                for row in range(sheet.nrows):
                    #  Get all the rows in the sheet (each rows is a list)
                    rowList = sheet.row_values(row)
                    for col in range(sheet.ncols):
                        #  Get all the values in each list
                        oneValue = rowList[col]
                        #  Copy the values to target worksheet

                        sheet_shipper_packer.write(row,col,oneValue) # Write the copied content to new worksheet in new combined excel

                # Create headers on the Notes worksheet in new combined excel
                sheet_notes.write(0,0,'Date')
                sheet_notes.write(0,1,'Name')
                sheet_notes.write(0,2,'Description')
                sheet_notes.write(0,3,'Dept.')
                
                # Resize cells in the Notes worksheet in new combined excel
                sheet_notes.col(2).width = 8000
                
                # Save new combined excel file
                wb.save(SAVE_FOLDER + os.sep + ocfile)

                
                MATCH_FOUND_LIST.append(match_ocfile)
                MATCH_FOUND_LIST.append(match_shipfile)
                
    except:
        continue


for i in MATCH_FOUND_LIST:
    try:
        os.remove(i)
    except:
        continue

done = input('All Matches found! Press enter to exit...')
    



    
            


        
