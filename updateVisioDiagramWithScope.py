import os
import win32com.client
import re
import pprint
import yaml


def pretty_print(o):
    pp = pprint.PrettyPrinter(indent=4)
    pp.pprint(o)

ROOT = r'C:\Users\rdapaz\Desktop'


vsd = win32com.client.Dispatch('Visio.Application')
vsd.Visible = True
doc = vsd.Documents.Open(os.path.join(ROOT, 'Fibre topology and rack requirements (BTS).vsdm'))
doc_page = doc.Pages("Electrical")

with open(r'C:\Users\rdapaz\Documents\doctools\proposed_cable_scope.yaml', 'r') as f:
    data = yaml.load(f)

pretty_print(data)


for _id, info in data.items():
    for idx in range(1, doc_page.Shapes.Count+1):
        shp = doc_page.Shapes(idx)
        if shp.Name == _id:
            shp.Text = f"{info}"