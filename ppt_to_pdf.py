#**************************************************************************************************#
#                                                                                                  #
# ppt_to_pdf                                                                                       #
# Saves each PowerPoint presentation as a PDF file.                                                #
#                                                                                                  #
# Usage: ppt_to_pdf [file]...                                                                      #
#                                                                                                  #
# Authors: S.G.M. Neves                                                                            #
#                                                                                                  #
#**************************************************************************************************#

# Import modules
from comtypes.client import Constants, CreateObject
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

    # Get the constant that specifies the PDF file format
    pdf_format = Constants(pp).ppSaveAsPDF

    for filepath in filepaths:

        # Check if the file exists
        if (not os.path.isfile(filepath)):
            print('The following file was not found:\n' + filepath)
            sys.exit(1)

        # Get the absolute path
        filepath = os.path.abspath(filepath)

        # Open PowerPoint file
        prs = pp.Presentations.Open(filepath)

        # Save as PDF
        prs.SaveAs(filepath.replace('.pptx', '.pdf'), pdf_format)

        # Close PowerPoint file
        prs.Close()

    # Close PowerPoint application
    pp.Quit()

if __name__ == '__main__':
    main()
