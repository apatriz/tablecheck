#!/usr/bin/env python

import arcpy,os, csv

EXCEL_FILE = r"C:\Users\patrizio\Projects\Monroe_Signs\test\Lookup_Table"

"""
1. read Excel file into dictionary
2. use arcpy update cursor to iterate through each table record (use only fields
in Excel table to increase cursor speed)
3. if any row value is None, get standard field value from 'Descrip' key and field name.
4. update row and move to next

"""

def load_excel(filename):
    res = {}
    wb = load_workbook(filename=filename,use_iterators=True)
    ws = wb.active
    header_range = "A1:S1"
    headers = []

    #get headers
    for row in ws.iter_rows(range_string=header_range):
        for cell in row:
            headers.append(cell.value)
    return headers
    
    
    
    
    
    
if __name__ == "__main__":
    print load_excel(EXCEL_FILE)
