from os.path import basename, getmtime, join
from os import remove, system
from glob import glob, iglob
from time import time
import argparse as ap


def get_latest(directory, backup_name):
    """
    Takes name of database and returns the most recently modified backup.
    """
    files = glob("{path}*".format(
        path=join(directory, backup_name)
    ))
    modified = [getmtime(i) for i in files]
    try:
        return basename(files[modified.index(max(modified))])
    except ValueError:
        raise ValueError(
            "No filenames found. Please check that the directory is correct."
        )


def robocopy(source, destination, filenames, options):
    commands = list()

    # mount network drive
    commands.append("pushd {directory}".format(directory=source))

    # call robocopy
    commands.append("robocopy . {destination} {files} {options}".format(
        destination=destination, files=" ".join(filenames),
        options=options
    ))

    # unmount network drive
    commands.append("popd")

    # run command
    system("&& ".join(commands))  # join separate commands with && (cmd.exe)


def create_parser():

    parser = ap.ArgumentParser(
        usage =
        """
        $ python update_bak.py <source> <destination> <files> -l
        """,
        description =
        """
        update_bak.py robocopies the latest copies of specifed"
        database backups.
        """
    )
    parser.add_argument(
        'source', type=str, help='Source directory files currently reside in.'
    )
    parser.add_argument(
        'destination', type=str,
        help='Directory to which files should be copied.'
    )
    parser.add_argument(
        'files', nargs='*', help='Databases for which to copy latest backup.'
    )
    parser.add_argument(
        '-l', '--list', action='store_const', const='/L',
        help='Run ROBOCOPY.exe in list mode (/L).'
    )
    return parser


def remove_old_copies(destination, files):
    dest_files = iglob(join(destination, '*'))
    for df in dest_files:
        for f in files:
            if f in df and (time() - getmtime(df) > 24 * 60 * 60):
                remove(df)
                print('{0} has been deleted.'.format(df))


def main():
    # handle arguments
    parser = create_parser()
    args = parser.parse_args()

    # copy bak files
    files = [get_latest(args.source, db) for db in args.files]
    robocopy(args.source, args.destination, files, args.list)

    # remove older bak files from destination directory
    if not args.list:
        remove_old_copies(args.destination, args.files)


if __name__ == "__main__":

    main()
