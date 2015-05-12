"""
Module with functions that can:
    load_file(fileName, header=0) - load files into memory
    extract_files(sheet_array, manager) - break multi-period data files
        into individual period files. Handles multiple periods in one
        sheet or multiple sheets. Detects whether sales or assets file.
    move_files(manager, top_dir) - move created files into appropriate
        file directory

Example usage (assumes pwd is the parent directory of RawData):
    $ ipython
    $ import raw
    $ sheet_array = raw.load_file('ex_file.xlsx', 1)  # header starts on 2nd row
    $ raw.extract_files(sheet_array, 'mgr_dir_name', 1, 0)  # startrow, startcol
    $ raw.move_files('mgr_dir_name', '.')
"""

import pandas as pd
import os
import sys
import re
import Tkinter, tkFileDialog


root = Tkinter.Tk()
root.withdraw()

# TODO: update docstrings to clearly label args
# TODO: refactor into class?


class Sheet(object):
    def __init__(self, name, data):
        self.name, self.data = name, data

    def get_name(self):
        return self.name

    def get_data(self):
        return self.data


def load_file(header=0):
    """
    Takes an Excel spreadsheet
    Returns an array of Sheet objects containing worksheet name and data
    """
    fileName = tkFileDialog.askopenfilename(
        filetypes=[('Excel', ('.xlsx', '.xls'))])
    try:
        data = pd.ExcelFile(fileName)
        sheets = data.sheet_names
        if is_single_sheet(sheets):
            measure = raw_input('Is this a sales or assets file? (if both ' +
                                'measures are present, indicate sales)\n>> ')
            return [Sheet(measure, data.parse(sheets[0], header))]
        return [Sheet(name, data.parse(name, header)) for name in sheets]
    except IOError as e:
        print "I/O error({0}): {1}".format(e.errno, e.strerror)
    except:
        print "Unexpected error:", sys.exc_info()[0]


def is_single_sheet(sheet_names):
    """
    Takes a list of Excel worksheet names
    Returns True if only a single sheet or all sheets have default names
    """
    single_sheet = len(sheet_names) == 1
    if len(sheet_names) > 1:
        excel_default_sheets = sheet_names[0] == 'Sheet1' and sheet_names[1] == 'Sheet2' \
            and sheet_names[2] == 'Sheet3'
    return single_sheet or excel_default_sheets


def delete_file(fileName):
    """
    Takes a file name
    Removes file from location if exists
    """
    try:
        os.remove(fileName)
    except WindowsError:
        pass


def find_date_col(df):
    """
    Takes a df file
    Returns name of date column
    """
    columns = df.columns
    date_cols = [col for col in columns if df[col].dtype == '<M8[ns]']
    if len(date_cols) > 1:
        raise ValueError('More than one DateTime column.')
    elif len(date_cols) == 0:
        raise ValueError('No date columns in dataframe.')
    else:
        return date_cols[0]


def build_dates(df, date_col):
    """
    Takes a dataframe and a column name containing Datetime data
    Returns a Pandas DatetimeIndex of all the unique dates in file
    """
    return pd.DatetimeIndex(df[date_col].unique())


def extract_files(sheet_array, manager, startrow=0, startcol=0):
    """
    Takes a Sheet object array and a manager directory name
    Returns None; applies process_sheet to each Sheet separately
    """
    for sheet in sheet_array:
        process_sheet(sheet, manager, startrow, startcol)
    print 'Processing complete'


def is_assets(col_name):
    """
    Takes a string col_name
    Returns true if asset is contained in col_name
    """
    pattern = re.compile('.*?[aA][sS][sS][eE][tT].*?')
    return re.search(pattern, col_name) is not None


def is_sales(col_name):
    """
    Takes a string col_name
    Returns true if sale is contained in col_name
    """
    pattern = re.compile('.*?[sS][aA][lL][eE].*?')
    return re.search(pattern, col_name) is not None


def process_sheet(sheet, manager, startrow, startcol):
    """
    Takes a Sheet object and a manager directory name
    Returns None; splits files by unique date in date column,
                  saves them in the current directory
    """
    df, name = sheet.get_data(), sheet.get_name()
    date_col = find_date_col(df)
    dates = build_dates(df, date_col)
    manager = manager
    for date in dates:
        print 'Processing %s' % date
        temp_df = df[df[date_col] == date]
        save_period_files(temp_df, name, date, manager, startrow, startcol)


def save_period_files(df, name, date, manager, startrow, startcol):
    """
    Takes a Pandas dataframe, date & manager name
    Returns None; saves a file '<manager> yyyymm.xlsx' for given date
    Note: deletes file of name name before saving!
    """
    if is_assets(name): measure = 'Asset'
    elif is_sales(name): measure = 'Sales'
    else:
        raise BaseException("Sheet '%s' is neither Sales nor Assets" % name)
    if 10 > date.month:
        temp_name = '%s %s %s0%s.xlsx' % (manager, measure, date.year, date.month)
    else:
        temp_name = '%s %s %s%s.xlsx' % (manager, measure, date.year, date.month)
    delete_file(temp_name)
    df.to_excel(temp_name, index=False, startrow=startrow, startcol=startcol)


def is_valid(manager, fileName):
    """
    Takes manager name and file name
    Returns Boolean if file name (roughly) matches extract_files format
    """
    return fileName[0:len(manager)] == manager and \
        fileName[-5:] == '.xlsx'


def extract_file_date(fileName):
    """
    Takes a file name
    Returns year, month strings
    """
    pattern = re.compile("[\d]{5,6}")
    date = re.search(pattern, fileName).group()
    return date[:4], date[-2:]


def extract_file_type(name):
    """
    Takes a file name
    Returns a string ('Assets' or 'Sales') if string in file name
    """
    if 'Asset' in name: return 'Asset'
    elif 'Sales' in name: return 'Sales'
    else: raise BaseException("Filename '%s' is neither Sales nor Assets")


def move_files(manager, top_dir='.'):
    """
    Takes a manager name and the directory name of the directory
    containing the RawData folder
    Returns None; moves all xlsx files matching extract_files naming
                  convention and moves them into the appropriate folder
    """
    file_list = os.listdir('.')
    for name in file_list:
        if is_valid(manager, name):
            yr, mth = extract_file_date(name)
            measure = extract_file_type(name)
            new_loc = '%s/RawData/%s/%s/%s/' % (top_dir, yr, mth, manager)
            new_name = '%s%s%s.xlsx' % (yr, mth, measure)
            if not os.path.exists(new_loc):
                os.makedirs(new_loc)
                print '\nCreated directory:\n%s' % new_loc
            delete_file(new_loc + new_name)
            os.rename(name, new_loc + new_name)
            print "\n'%s' moved to:\n%s" % (name, new_loc)
    print '\nProcessing complete'


def check_args(user_in):
    """
    Checks user input for exit commands
    """
    pattern = re.compile('(quit|exit|q)', re.IGNORECASE)
    if re.search(pattern, user_in):
        sys.exit()


def main():

    header_in = raw_input('Please indicate which row the headers ' +
                          'are in:\n>> ')
    check_args(header_in)

    # load file
    try:
        header_in = int(header_in)
    except ValueError:
        print 'This is not a valid row number.'
        sys.exit()

    data = load_file(header_in - 1)

    manager_name = raw_input('Please indicate the participant name ' +
                             '(NOTE: it should match the directory ' +
                             'name exactly!)\n>> ')
    check_args(manager_name)

    # extract files
    extract_files(data, manager_name)

    run_move_files = raw_input('Please check that the extracted files ' +
                               'are named and formatted correctly.\n\n' +
                               'Press any key to select the folder ' +
                               'containing the RawData folder.')
    check_args(run_move_files)

    top_path = tkFileDialog.askdirectory()
    if 'RawData' not in os.listdir(top_path):
        create_raw = raw_input('RawData is not in ' + top_path +
                               '. Do you wish to create it? (Y/N)\n>> ')
        if create_raw == 'N':
            sys.exit()

    # move files
    move_files(manager_name, top_path)

    raw_input('\nPress any key to exit.')


if __name__ == '__main__':

    main()
