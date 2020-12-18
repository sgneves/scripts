#**************************************************************************************************#
#                                                                                                  #
# ppt_to_pdf                                                                                       #
# Saves the PowerPoint presentations as PDF files.                                                 #
#                                                                                                  #
# Usage: ppt_to_pdf [file]...                                                                      #
#                                                                                                  #
# Authors: S.G.M. Neves                                                                            #
#                                                                                                  #
#**************************************************************************************************#

# Import modules
import comtypes.client
import glob
import os
import sys

def main():

    # Get the names of the PPTX files
    if len(sys.argv) == 1:

        filepaths = glob.glob('*.pptx')
    else:
        filepaths = sys.argv[1:]

    # Get current working directory
    directory = os.getcwd()

    # Open PowerPoint application
    pp = comtypes.client.CreateObject("Powerpoint.Application")

    for filepath in filepaths:

        # Check if the filepath contains the directory
        if os.path.dirname(filepath) == '':
            filepath = directory + os.path.sep + filepath

        # Check if the file exists
        if (not os.path.isfile(filepath)):
            print('The following file was not found:\n' + filepath)
            sys.exit(1)

        # Open PowerPoint file
        prs = pp.Presentations.Open(filepath)

        # Save as PDF
        prs.SaveAs(filepath.replace('.pptx', '.pdf'), 32)

        # Close PowerPoint file
        prs.Close()

    # Close PowerPoint application
    pp.Quit()

if __name__ == '__main__':
    main()
