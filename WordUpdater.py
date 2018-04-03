# -*- coding: utf-8 -*-

import win32com.client
import re
import pprint


def pretty_print(o):
    pp = pprint.PrettyPrinter(indent=4)
    pp.pprint(o)


class Word:

    def __init__(self, path):
        self.path = path
        self.app = win32com.client.gencache.EnsureDispatch('Word.Application')
        self.app.Visible = True
        self.app.DisplayAlerts = False
        self.app.Documents.Open(self.path)
        self.doc = self.app.ActiveDocument

    def updateTable(self, bookmark, data, heading_rows=1):
        word_range = self.doc.Bookmarks(bookmark).Range
        table = word_range.Tables(1)
        rows_count = table.Rows.Count
        if not rows_count >= len(data) + heading_rows:
            table.Select()
            self.app.Selection.InsertRowsBelow(NumRows=len(data) + heading_rows - rows_count)
        i = heading_rows
        for entry in data: #sorted(data, key=lambda x: (x[0], x[1])):
            i += 1
            for n in range(len(entry)):
                table.Cell(i, n+1).Range.Text = entry[n]

    def updateIDs(self, bookmark, prefix):
        rex = re.compile('[A-Z]+', re.IGNORECASE)
        word_range = self.doc.Bookmarks(bookmark).Range
        table = word_range.Tables(1)
        rows_count = table.Rows.Count
        count = 0
        for rid in range(1, rows_count + 1):
            m = rex.search(table.Cell(rid, 1).Range.Text)
            if m:
                pass
            else:
                count += 1
                table.Cell(rid, 1).Range.Text = f"{prefix}-{count:02}"


def make_data():
    data = """
|21/06/2017|10:00am|Log into GUI and test logon and operation|Chatura Fernando
|21/06/2017|10:00am|Log into grapevine/SSH and test logon and operation|Chatura Fernando
|21/06/2017|10:00am|Log into CIMC and test logon and operation|Chatura Fernando
|21/06/2017|10:00am|Perform discovery of network devices with the appliance in situ|Chatura Fernando
|21/06/2017|10:00am|Perform a device inventory using the tool's capabilities|Chatura Fernando
|21/06/2017|10:00am|Perform a host inventory using the tool's capabilities|Chatura Fernando
|21/06/2017|10:00am|Generate a network map using the tool's capabilities|Chatura Fernando
|21/06/2017|10:00am|Generate a path trace usinng the tool's capabilities|Chatura Fernando
|21/06/2017|10:00am|Sign off commissioning tests and perform handover to Network Tower|Chatura Fernando
""".splitlines()
    data = [x.split('|') for x in data if len(x) > 0]
    '''
    new_data = []
    main_counter = 0
    for tier_site, switch, cur_model in data:
        if len(switch) <= 1:
            main_counter += 1
            sub_counter = 0
        else:
            sub_counter +=1
        new_data.append([f"{main_counter}.{sub_counter}", tier_site, switch, cur_model])
    '''
    return data

def main(bookmark, data=[], heading_rows=1):
    my_path = r'C:\Users\ric\Documents\I0431900 (106187) APIC-EM Full Deployment - Implementation Plan.docx'
    wd = Word(my_path)
    wd.updateTable(bookmark, data, heading_rows)
    wd.updateIDs(bookmark, prefix="TSK-")

def mock(data, **kwargs):
    pretty_print(data)

if __name__ == "__main__":
    data = make_data()
    mock(bookmark='bk3', data=data, heading_rows=1)
    main(bookmark='bk3', data=data, heading_rows=1)
