"""
Module with functions that can:
    load_file(fileName, sheet_num=0) - load files into memory
    extract_files(df, manager) - break multi-period data files 
        into individual period files
    move_files(manager, top_dir) - move created files into appropriate
        file directory

"""


import pandas as pd
import os
import re

def load_file(fileName, sheet_num=0):
    '''
    Takes an Excel spreadsheet (loads first sheet by default)
    returns a Pandas dataframe
    '''
    try:
        data = pd.ExcelFile(fileName)
        sheets = data.sheet_names
        return pd.read_excel(fileName, sheets[sheet_num])
    except IOError as e:
        print "I/O error({0}): {1}".format(e.errno, e.strerror)
    except:
        print "Unexpected error:", sys.ex_info()[0]

def delete_file(fileName):
    '''
    Takes a file name
    Removes file from location if exists
    '''
    try:
        os.remove(fileName)
    except WindowsError:
        pass

def find_date_col(df):
    '''
    Takes a df file
    Returns name of date column
    '''
    columns = df.columns
    date_cols = [col for col in columns if df[col].dtype == '<M8[ns]']
    if len(date_cols) > 1:
        raise ValueError('More than one DateTime column.')
    else:
        return date_cols[0]

def build_dates(df, date_col):
    '''
    Takes a dataframe and a column name containing the date
    Returns a Pandas DatetimeIndex of all the unique dates in file
    '''
    return pd.DatetimeIndex(df[date_col].unique())

def extract_files(df, manager):
    '''
    Takes a Pandas dataframe and a column name containing the date
    Returns None; splits files by unique date in date column,
                  saves them in the current directory
    '''
    date_col = find_date_col(df)
    dates = build_dates(df, date_col)
    manager = manager.title()
    for date in dates:
        print 'Processing %s' % date
        temp_df = df[df[date_col] == date]
        if 10 > date.month:
            temp_name = '%s %s0%s.xlsx' % (manager, date.year, date.month)
        else:
            temp_name = '%s %s%s.xlsx' % (manager, date.year, date.month)
        delete_file(temp_name)
        temp_df.to_excel(temp_name, index=False)

def is_valid(manager, fileName):
    '''
    Takes manager name and file name
    Returns Boolean if file name matches extract_files format
    '''
    return fileName[0:len(manager)] == manager.title() and \
            fileName[-5:] == '.xlsx'

def extract_file_date(fileName):
    '''
    Takes a file name
    Returns year, month strings
    '''
    pattern = re.compile("[\d]{5,6}")
    date = re.search(pattern, fileName).group()
    return date[:4], date[-2:]
    
def move_files(manager, top_dir):
    '''
    Takes a manager name and the directory containing the RawData folder
    Returns None; moves all xlsx files matching extract_files naming
                  convention and moves them into the appropriate folder
    '''
    manager = manager.title()
    file_list = os.listdir('.')
    for name in file_list:
        if is_valid(manager, name):
            yr, mth = extract_file_date(name)
            new_loc = '%s/RawData/%s/%s/%s/%s%sSales.xlsx' % \
                    (top_dir, yr, mth, manager, yr, mth)
            delete_file(new_loc)
            os.rename(name, new_loc)
            print '%s moved to %s' % (name, new_loc)
    print 'Processing complete'
