#!/usr/bin/env python

"""
This script can be used to update a database table in Arc. The script
will fill in the database records with missing values based on the lookup table (
generated from and Excel .xlsx file).

"""

import arcpy,os
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string,get_column_letter

###BEGIN PARAMETER INPUT#############################

# Set the Excel file to use as a lookup table
EXCEL_FILE = r"C:\Users\patrizio\Projects\Monroe_Signs\test\data_v2\Lookup_Table.xlsx"

# Set the Excel table range to extract the table header names.
# The header names should be equivalent to the field names of the database table (but do not need to have the same name)
FIELD_NAME_RANGE = "A1:S1"

# Set the lookup table column which contains the shared id (foreign key)
ID_COLUMN = "A"

# Set the database table that will be updated
DATABASE_TABLE = r"C:\Users\patrizio\Projects\Monroe_Signs\test\data_v2\Monroe_Signs.gdb\Signs"

#Set the database foreign key field name corresponding to the lookup table id
FOREIGN_KEY = "SignType"

# Set the field map.
# Maps the field name in the database table to the header in the lookup table. This
# is used in case the field naming convention/order differs between both tables.
FIELD_MAP = {"SignType":"Code",
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

##############END PARAMETER INPUT############################


#decorator to convert indexes to 0-based
def convertindex(func):
    def minus_one(index_string):
        return func(index_string) - 1
    return minus_one


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


def load_dict(filename,field_range,id_column):
    wb = load_workbook(filename=filename,read_only=True)
    ws = wb.active
    res = {}
    get_index = convertindex(column_index_from_string)
    headers = get_headers(filename,field_range)
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
        res[row[get_index(id_column)].value] = entry
    return res


def check_arc_table(filename,foreign_key,field_map,lookup_table):
    fields = [field.name for field in arcpy.ListFields(filename) if field.name in field_map]
    with arcpy.da.UpdateCursor(filename,fields) as cursor:
        for row in cursor:
            index = 0
            new_row = []
            for cell in row:
                if fields[index] == foreign_key and cell:
                    lookup_id = cell
                elif fields[index] == foreign_key and not cell:
                    new_row = row
                    break
                if not cell and lookup_id in lookup_table:
                    lookup_value = field_map[cursor.fields[index]]
                    if lookup_table[lookup_id][lookup_value]:
                        cell = lookup_table[lookup_id][lookup_value]
                new_row.append(cell)
                index += 1
            new_row = tuple(new_row)
            cursor.updateRow(new_row)
    print "Finsihed updating table {0}".format(os.path.abspath(filename))
    return os.path.abspath(filename)


def test_tablecheck(test_entries):
    
    lookup_table = load_dict(EXCEL_FILE,FIELD_NAME_RANGE,ID_COLUMN)
    
    for i in test_entries:
        assert lookup_table[i] == test_entries[i] 
    print "Test complete."


if __name__ == "__main__":

##    print test_tablecheck({"D4-3":{u'Code': u'D4-3', u'SheetingColor2': None,
##                                        u'RegPkVehExceptions1': None, u'DimHeight': 18L,
##                                        u'SheetingColor1': u'White', u'LegendColor1': u'Green',
##                                        u'LegendColor2': None, u'LegendColor3': None, u'RegPkRestrType2': None,
##                                        u'RegPkRestrType1': None, u'RegPkTimeLimit2': None, u'Collect': u'LiDAR',
##                                        u'RegPkTimeLimit1': None, u'RegPkVehExceptions2': None, u'RegPkArrow1': None,
##                                        u'Descrip': u'Bike Parking: D4-3', u'RegPkTimeYear1': None, u'RegPkTimeYear2': None, u'DimWidth': 12L},
##                          "W1-2aL":{u'Code': u'W1-2aL', u'SheetingColor2': None, u'RegPkVehExceptions1': None,
##                                     u'DimHeight': 30L, u'SheetingColor1': u'Yellow', u'LegendColor1': u'Black',
##                                     u'LegendColor2': None, u'LegendColor3': None, u'RegPkRestrType2': None,
##                                     u'RegPkRestrType1': None, u'RegPkTimeLimit2': None, u'Collect': u'LiDAR',
##                                     u'RegPkTimeLimit1': None, u'RegPkVehExceptions2': None,
##                                     u'RegPkArrow1': None, u'Descrip': u'Curve Left: W1-2aL',
##                                     u'RegPkTimeYear1': None, u'RegPkTimeYear2': None, u'DimWidth': 30L}})
    lookup_table = load_dict(EXCEL_FILE,FIELD_NAME_RANGE,ID_COLUMN)
    check_arc_table(DATABASE_TABLE,FOREIGN_KEY,FIELD_MAP,lookup_table)


