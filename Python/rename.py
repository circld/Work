"""
Utility to rename all subfolders (of arbitrary depth) matching particular
name to a new directory name. Note that the directory contents are not affected.
"""
__author__ = 'Paul Garaud'
__date__ = '4/14/2015'


import os
import argparse as ap
from pprint import pprint


pwd = os.getcwd()


def rename_sub_dirs(d, old_name, new_name):

    for i in os.listdir(d):
        idir = os.path.join(d, i)
        if os.path.isdir(idir):
            if i == old_name:
                renamed = os.path.join(d, new_name)
                os.rename(idir, renamed)
                yield renamed
            else:
                for j in rename_sub_dirs(
                        os.path.join(d, idir), old_name, new_name
                ):
                    yield j


if __name__ == '__main__':

    argparser = ap.ArgumentParser(
        prog='rename.py',
        description='Utility to rename all subfolders (of arbitrary depth) '
        'matching particular name to a new directory name. Note that the '
        'directory contents are not affected.',
        usage='$ python rename.py old_dir_name new_dir_name',
    )

    argparser.add_argument(
        'old_dir_name', type=str, help='Subdirectory name to change.'
    )
    argparser.add_argument(
        'new_dir_name', type=str, help='New name with which to replace old_dir_name.'
    )
    args = argparser.parse_args()

    x = rename_sub_dirs(pwd, args.old_dir_name, args.new_dir_name)

    # print all changes to stdout
    print('Directories successfully renamed:')
    pprint(list(x))
