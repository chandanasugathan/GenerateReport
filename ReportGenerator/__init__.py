from __future__ import print_function
from mailmerge import MailMerge


class ReportGenerator:
    template = "template.docx"
    table_entries = []
    list_print = []
    output = "Generated_Report.docx"

    def __init__(self):
        self.template = "template.docx"
        table_row = {'sr_no': '1', 'subject': 'EM-II', 'mark': '17', 'attendance': '90'}
        self.table_entries.append(table_row)
        self.table_entries.append(table_row)
        self.table_entries.append(table_row)
        self.table_entries.append(table_row)
        custom_1 = {
            'month1': 'January',
            'month2': 'February',
            'month3': 'February',
            'year1': '2018',
            'year2': '2018',
            'name': 'PATKAR MALLISHA S.',
            'branch': 'CST',
            'division': 'C',
            'roll': '1',
            'sem': 'I',
            'sr_no': self.table_entries}
        self.list_print.append(custom_1)
        self.list_print.append(custom_1)

    def __init__(self, template, output, report_data):
        print("Template File: ")
        print(template)
        self.list_print = report_data
        self.template = template
        self.output = output

    def generate_report(self):
        print("Printing to Word File...")
        document = MailMerge(self.template)
        print(document.get_merge_fields())
        document.merge_templates(self.list_print, separator='page_break')
        document.write(self.output)
        print("Report Generated in {0} file".format(self.output))
