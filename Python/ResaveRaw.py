__author__ = 'Paul Garaud'
__credits__ = 'First version created 10/27/2014'

import os
import glob
import argparse as ap
import Tkinter, tkFileDialog


def resave(fpath, year, month, manager):
    year, month, manager = [i or '*' for i in (year, month, manager)]
    files = glob.glob('{fpath}/{year}/{month}/{manager}/*'.format(
        fpath=fpath, year=year, month=month, manager=manager))
    print '\nUpdated last modified time for:'
    if len(files) == 0:
        print 'No files modified.'
    for f in files:
        os.utime(f, None)
        print f


def main():
    parser = ap.ArgumentParser(
        description=
        """
        Utility script to 'resave' (ie, update last modified date) on all
        selected files (by any combination of manager, month, and/or year).
        """
        )
    # add & define command line args
#    parser.add_argument('path',
#                        help='Parent directory for all raw files.')
    parser.add_argument('--year', help='Resave all files for a given year.')
    parser.add_argument('--month', help='Resave all files for a given month.')
    parser.add_argument('--manager',
                        help='Resave all files for a given manager.')

    args = parser.parse_args()

    root = Tkinter.Tk()
    root.withdraw()
    filepath = tkFileDialog.askdirectory(title='Please select the RawData' +
                                         ' directory.')

    resave(filepath, args.year, args.month, args.manager)


if __name__ == '__main__':

    main()
