#!/usr/bin/env python

import arcpy,os
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

EXCEL_FILE = r"C:\Users\patrizio\Projects\Monroe_Signs\test\Lookup_Table.xlsx"
FIELD_NAME_RANGE = "A1:S1"
ID_COLUMN = "A"

DATABASE_TABLE = r"C:\Users\patrizio\Projects\Monroe_Signs\test\Monroe_Signs.gdb\Signs"
FIELD_MAP = {"SignType":"Descrip",
             "DimHeight":"DimHeight",
             "DimWidth":"DimWidth",
             "LegendColor1":"LegendColor1",
             "LegendColor2":"LegendColor2",
             "LegendColor3":"LegendColor3",
             "SheetColor1":"SheetingColor1",
             "SheetColor2":"SheetingColor2",
             "RegPkRestr1":"RegPkRestrType1",
             "RegPkRestr2":"RegPkRestrType2",
             "RegPkTimeLim1":"RegPkTimeLimit1",
             "RegPkTimeLim2":"RegPkTimeLimit2",
             "RegPkArrow1":"RegPkArrow1",
             "RegPkTimeYear1":"RegPkTimeYear1",
             "RegPkTimeYear2":"RegPkTimeYear2",
             "RegPkVehExcep1":"RegPkVehExceptions1",
             "RegPkVehExcep2":"RegPkVehExceptions1",
             }
             

"""
1. read Excel file into dictionary - DONE
2. use arcpy update cursor to iterate through each table record (use only fields
in Excel table to increase cursor speed)
3. if any row value is None, get standard field value from 'Descrip' key and field name.
4. update row and move to next

"""

#decorator to convert indexes to 0-based
def convertindex(func):
    def minus_one(index_string):
        return func(index_string) - 1
    return minus_one
get_index = convertindex(column_index_from_string)

    

def get_headers(filename,field_range):
    """
    Input:
    filename -> path to excel file
    field_range -> range of cell coordinates for getting headers, e.g. 'A1:S1'

    Output: list
    """
    wb = load_workbook(filename=filename,read_only=True)
    ws = wb.active
    headers = []

    for row in ws.iter_rows(range_string=field_range):
        for cell in row:
            headers.append(cell.value)
    return headers


def load_dict(filename,headers,key_id):
    wb = load_workbook(filename=filename,read_only=True)
    ws = wb.active
    res = {}
    for row in ws.iter_rows(row_offset=1):
        index = 0
        entry = {}     
        for cell in row:
            try:
                val = cell.value.strip()
            except AttributeError:
                val = cell.value
            entry[headers[index]] = val
            index += 1
        res[row[key_id].value] = entry        
    return res

def check_arc_table(filename,id_field_name,field_map,lookup_table):
    fields = [field.name for field in arcpy.ListFields(filename) if field.name in field_map]
    print fields
    with arcpy.da.UpdateCursor(filename,fields) as cursor:
        for row in cursor:
            index = 0
            new_row = []
            for cell in row:
                if fields[index] == id_field_name and cell:
                    lookup_id = cell
                    print lookup_id
                elif fields[index] == id_field_name and not cell:
                    new_row = row
                    break
                if not cell and lookup_id in lookup_table:  
                    lookup_value = field_map[cursor.fields[index]]
                    print "Lookup_value:{0}".format(lookup_value)
                    if lookup_table[lookup_id][lookup_value]:
                        cell = lookup_table[lookup_id][lookup_value]
                        print "Adding value {0} to cell".format(cell)
                new_row.append(cell)
                index += 1
            new_row = tuple(new_row)
            print new_row
            cursor.updateRow(new_row)
                
                
                
        
              
    
if __name__ == "__main__":
    headers = get_headers(EXCEL_FILE,FIELD_NAME_RANGE)
    key_id = get_index(ID_COLUMN)
    print headers
    lookup_table = load_dict(EXCEL_FILE,headers,key_id)
##    print lookup_table["D4-3"]

    check_arc_table(DATABASE_TABLE,"SignType",FIELD_MAP,lookup_table)
    
    
