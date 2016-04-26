# ERP Contact Import Tool
# This utility fixes the contacts csv file for importing into our ERP (Odoo)
# Specifically this utility is for the step of replacing the company names with their export_ids after the companies have been imported

# Copyright (c) 2016 David Todd (alopexc0de) https://c0defox.es

# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:

# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.

# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

import csv
import os
import re

# Change these paths to reflect your needs
config = {
    'partner_path': 'C:\\Users\\dtodd\\Desktop\\res.partner (2).csv',                                                   # This is the csv file containing two columns, id and name
    'import_path': 'C:\\Users\\dtodd\\Desktop\\To Import\\MN Companies-50 FTE-HQ one (contacts) Campaign.csv',          # This is the csv file containing whatever needs to be imported
    'temp_path': 'C:\\Users\\dtodd\\Desktop\\temp',                                                                     # This is a csv file that will be deleted on program completion
    'column_to_change': 'parent_id/id'                                                                                  # This is the name of the column that you want to change
}

# Find the index number of a column in a row and return that
def find_index(row=None, col=None):
    # if ((row == None) | (col == None)) | ((type(row) != type(str)) | (type(col) != type(str))):
    #     raise ValueError('find_index requires both row and col to be set to strings')
    #     exit(1)

    index = None
    for i,j in enumerate(row):
        if j == col:
            index = i
            break
    return index

os.rename(config['import_path'], config['import_path']+'.orig') # Store the original import_path in a backup file

with open(config['partner_path'], 'r') as partner_csv:
    reader = csv.reader(partner_csv)

    first_row = next(reader)                                                                        
    if ('id' not in first_row) | ('name' not in first_row):
        raise IndexError('Either \'id\' or \'name\' is not defined in your partner_path csv file')
        exit(1)

    id = find_index(first_row, 'id')
    name = find_index(first_row, 'name')

    if (id == None) | (name == None):
        raise ValueError('Either id or name could not be indexed with find_index')
        exit(1)

    partner_ids = {rows[id]:rows[name] for rows in reader} # Build a dict containing the company names as the keys and the __export__IDs as values
    reader = None

    # Open both the import_csv and temp_csv files
    with open(config['import_path']+'.orig', 'r') as import_csv, open(config['temp_path'], 'w') as temp_csv:
        reader = csv.reader(import_csv)
        writer = csv.writer(temp_csv, lineterminator='\n')

        first_row = next(reader) 
        if config['column_to_change'] not in first_row:
            raise IndexError(config['column_to_change']+' is not defined in your import_path csv file')
            exit(1)

        col = find_index(first_row, config['column_to_change'])

        if col != None: 
            writer.writerow(first_row)
            for lines in reader:
                # Taken from http://stackoverflow.com/a/2400577/5727514 and modified to support what I need 
                pattern = re.compile('|'.join(partner_ids.keys()))
                lines[col] = pattern.sub(lambda x: partner_ids[x.group()], lines[col])

                writer.writerow(lines)
        else:
            raise ValueError('col could not be indexed with find_index')
            exit(1)

# Close our files to be nice to the OS
    temp_csv.close()
    import_csv.close()
partner_csv.close()

# On program completion, move the temp file to the original file
os.rename(config['temp_path'], config['import_path'])