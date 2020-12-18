#**************************************************************************************************#
#                                                                                                  #
# ppt_set_title                                                                                    #
# Sets the title of each PowerPoint presentation equal to the filename.                            #
#                                                                                                  #
# Usage: ppt_set_title [file]...                                                                   #
#                                                                                                  #
# Authors: S.G.M. Neves                                                                            #
#                                                                                                  #
#**************************************************************************************************#

# Import modules
from comtypes.client import CreateObject
import glob
import os
import sys

def main():

    # Get the names of the PPTX files
    if len(sys.argv) == 1:

        filepaths = glob.glob('*.pptx')
    else:
        filepaths = sys.argv[1:]

    # Open PowerPoint application
    pp = CreateObject("Powerpoint.Application")

    for filepath in filepaths:

        # Check if the file exists
        if (not os.path.isfile(filepath)):
            print('The following file was not found:\n' + filepath)
            sys.exit(1)

        # Get the absolute path
        filepath = os.path.abspath(filepath)

        # Open PowerPoint file
        prs = pp.Presentations.Open(filepath)

        # Set the title
        prs.BuiltinDocumentProperties['Title'] = os.path.splitext(os.path.basename(filepath))[0]

        # Save and close PowerPoint file
        prs.Save()
        prs.Close()

    # Close PowerPoint application
    pp.Quit()

if __name__ == '__main__':
    main()
