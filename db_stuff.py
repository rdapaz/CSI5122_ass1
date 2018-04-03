import pprint
import win32com.client
from decimal import Decimal
import datetime
import psycopg2


def pretty_printer(o):
	pp = pprint.PrettyPrinter(indent=4)
	pp.pprint(o)


def value_colName(iVal):
    retVal = None
    if iVal <= 26:
        retVal = chr(64+iVal)
    else:
        m = int(iVal/26)
        n = iVal - m*26
        if n==0:
            m = m-1
            n = 26
        retVal = f'{value_colName(m)}{value_colName(n)}' 
    return retVal


fields = """
_name|text
status|text
description|text
_references|text
phase|text
votes|text
comments|text
""".splitlines()


def int_or_same(int_p):
	try:
		return int(int_p)
	except:
		return int_p


fields = [x.split('|') for x in fields if len(x) > 0]

column_names = [value_colName(x+1) for x in range(len(fields))]

data_fields = {k: v for k, v in zip(column_names, fields)}
pretty_printer(data_fields)

arr = []
for col_name, rest in data_fields.items():
	field_name, data_type = rest
	arr.append(dict(col_name=col_name, field_name=field_name, data_type=data_type))

xlApp = win32com.client.gencache.EnsureDispatch('Excel.Application')
xlApp.Visible = True


path = r'C:\Users\rdapaz\Desktop\allitems.xlsx'
wk = xlApp.Workbooks.Open(path)
sh = wk.Worksheets('allitems')
print(sh.Name)

EOF = sh.Range('A1000000').End(-4162).Row
print(EOF)

vals = []
for row in range(4, EOF+1):
	for p in arr:
		if p['data_type'] == 'text':
			exec(f"{p['field_name']} = str(sh.Range('{p['col_name']}{row}').Value) if sh.Range('{p['col_name']}{row}').Value else None")
		elif p['data_type'] == 'decimal':
			exec(f"{p['field_name']} = Decimal(sh.Range('{p['col_name']}{row}').Value) if sh.Range('{p['col_name']}{row}').Value else 0.0")
		elif p['data_type'] in ('int', 'long'):
			exec(f"{p['field_name']} = int_or_same(sh.Range('{p['col_name']}{row}').Value) if sh.Range('{p['col_name']}{row}').Value else 0")
		elif p['data_type'] == 'date':
			exec(f"{p['field_name']} = str(sh.Range('{p['col_name']}{row}').Value) if sh.Range('{p['col_name']}{row}').Value else '1970-01-01'")
			exec(f"{p['field_name']} = {p['field_name']}[:10]")
			exec(f"{p['field_name']} = datetime.datetime.strptime({p['field_name']}, '%Y-%m-%d')")
	exec("vals.append({})".format([eval(p['field_name']) for p in arr]))

pretty_printer(vals)

conn = psycopg2.connect("dbname='vulnerable' user=postgres")
if True:
    cur = conn.cursor()

    sql = """
    	CREATE TABLE IF NOT EXISTS \"public\".\"cve\" (
    		id serial primary key,
            _name text,
            status text,
            description text,
            _references text,
            phase text,
            votes text,
            comments text
            )
	"""

    cur.execute(sql)
    
    sql =  """ INSERT INTO \"public\".\"cve\" (
                _name, status, description, _references, phase, votes, comments
                ) VALUES 
           		(%s, %s, %s, %s, %s, %s, %s) 
            """
    cur.executemany(sql, vals)
    conn.commit()
    conn.close()