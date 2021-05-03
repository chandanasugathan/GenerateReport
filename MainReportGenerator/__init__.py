# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import os.path

import MarksheetParser
import ReportGenerator

def Pass_Params(input_filename, output_folder):
    ob = MarksheetParser.MarksheetParser(str(input_filename))
    list_print = ob.generate_report_data()
    print("Output Directory :")
    print(output_folder)
    rp = ReportGenerator.ReportGenerator("resources\Template.docx", os.path.join(output_folder,"Generated_Report.docx"),list_print)
    rp.generate_report()
# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    Pass_Params("..\\resources\input.xlsx", os.path.curdir)

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
