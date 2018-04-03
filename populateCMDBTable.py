import pprint
import win32com.client
from decimal import Decimal
import datetime
import psycopg2
import json


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
search_code|text
host_name|text
_id|text
ip_address|text
impact|text
model|text
name|text
serial_number|text
source_code_location|text
start_date|text
status|text
admin_access_required|text
admin_url|text
alternate_software|text
application_portfolio|text
application_software_status|text
application_type|text
approval_group|text
asset_tag|text
assigned|text
assigned_to|text
assignment_group|text
attributes|text
business_owner|text
cpu|text
can_print|text
category|text
checked_in|text
checked_out|text
comments|text
commission_date|text
company|text
correlation_id|text
cost_center|text
customer_vendor_agreements|text
dns_domain|text
dsl|text
date_disposed|text
decommission_reason|text
default_gateway|text
department|text
description|text
disaster_recovery|text
discovery_source|text
due|text
due_in|text
esp_validated|text
end_of_lease|text
environment|text
fault_count|text
first_discovered|text
fully_qualified_domain_name|text
gl_account|text
goods_received_date|text
hdd|text
import_data|text
installed|text
invoice_number|text
justification|text
key_business_function|text
lease_review_date|text
lease_schedule_number|text
lease_contract|text
licence_type|text
license_expiry|text
license_quantity|text
life_cycle_status|text
location|text
mac_address|text
maintenance_expiry|text
maintenance_provider|text
maintenance_provider_lookup|text
make|text
managed_ci|text
managed_by|text
manufactured_by|text
manufacturer|text
memory|text
metered_software|text
model_id|text
monitor|text
most_recent_discovery|text
name_2|text
operating_system|text
operational_status|text
order_received|text
ordered|text
ownership|text
purchase_order|text
purchase_cost|text
purchase_date|text
purchased|text
ru_position|text
related_service_documentation|text
sccm_package_name|text
status_not_used|text
subtype|text
subcategory|text
substatus|text
supplier|text
supplier_lookup|text
support_workgroup|text
supported_by|text
tier_type|text
_type|text
user_authentication_required|text
user_url|text
vendor|text
ver_ctrl_repository|text
version_no|text
warranty_expiry|text
warranty_expiration|text
sys_id|text
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


path = r'E:\Projects\Western Power\FY17 Network Refresh\WP_Network_Inventory.xls'
wk = xlApp.Workbooks.Open(path)

vals = []
for shIdx in range(1, wk.Worksheets.Count+1):
    sh = wk.Worksheets(shIdx)
    if sh.Name == 'Page 1':

        EOF = sh.Range('A65536').End(-4162).Row

        for row in range(2, EOF+1):
            for p in arr:
                print(p['field_name'], p['col_name'], sep='|')
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

def json_serial(obj):
    """JSON serializer for objects not serializable by default json code"""
    if isinstance(obj, (datetime.datetime, datetime.date)):
        return obj.isoformat()
    raise TypeError ("Type %s not serializable" % type(obj))

with open(r'C:\Users\rdapaz\Documents\python_assorted\cmdb.json', 'w') as fout:
    json.dump(vals, fout, indent=True, default=json_serial)

conn = psycopg2.connect("dbname='Analyse Expenses' user=postgres")
# try:
if True:
    cur = conn.cursor()

    sql = """
        CREATE TABLE IF NOT EXISTS \"public\".\"CMDB\" (
        bom_id serial primary key,
        search_code text, 
        host_name text, 
        _id text, 
        ip_address text, 
        impact text, 
        model text, 
        name text, 
        serial_number text, 
        source_code_location text, 
        start_date text, 
        status text, 
        admin_access_required text, 
        admin_url text, 
        alternate_software text, 
        application_portfolio text, 
        application_software_status text, 
        application_type text, 
        approval_group text, 
        asset_tag text, 
        assigned text, 
        assigned_to text, 
        assignment_group text, 
        attributes text, 
        business_owner text, 
        cpu text, 
        can_print text, 
        category text, 
        checked_in text, 
        checked_out text, 
        comments text, 
        commission_date text, 
        company text, 
        correlation_id text, 
        cost_center text, 
        customer_vendor_agreements text, 
        dns_domain text, 
        dsl text, 
        date_disposed text, 
        decommission_reason text, 
        default_gateway text, 
        department text, 
        description text, 
        disaster_recovery text, 
        discovery_source text, 
        due text, 
        due_in text, 
        esp_validated text, 
        end_of_lease text, 
        environment text, 
        fault_count text, 
        first_discovered text, 
        fully_qualified_domain_name text, 
        gl_account text, 
        goods_received_date text, 
        hdd text, 
        import_data text, 
        installed text, 
        invoice_number text, 
        justification text, 
        key_business_function text, 
        lease_review_date text, 
        lease_schedule_number text, 
        lease_contract text, 
        licence_type text, 
        license_expiry text, 
        license_quantity text, 
        life_cycle_status text, 
        location text, 
        mac_address text, 
        maintenance_expiry text, 
        maintenance_provider text, 
        maintenance_provider_lookup text, 
        make text, 
        managed_ci text, 
        managed_by text, 
        manufactured_by text, 
        manufacturer text, 
        memory text, 
        metered_software text, 
        model_id text, 
        monitor text, 
        most_recent_discovery text, 
        name_2 text, 
        operating_system text, 
        operational_status text, 
        order_received text, 
        ordered text, 
        ownership text, 
        purchase_order text, 
        purchase_cost text, 
        purchase_date text, 
        purchased text, 
        ru_position text, 
        related_service_documentation text, 
        sccm_package_name text, 
        status_not_used text, 
        subtype text, 
        subcategory text, 
        substatus text, 
        supplier text, 
        supplier_lookup text, 
        support_workgroup text, 
        supported_by text, 
        tier_type text, 
        _type text, 
        user_authentication_required text, 
        user_url text, 
        vendor text, 
        ver_ctrl_repository text, 
        version_no text, 
        warranty_expiry text, 
        warranty_expiration text, 
        sys_id text 
        )
    """

    cur.execute(sql)
    
    sql =  """ INSERT INTO \"public\".\"CMDB\" ( search_code, 
                host_name, _id, ip_address, impact, model, name, serial_number, source_code_location, start_date, status, admin_access_required, 
                admin_url, alternate_software, application_portfolio, application_software_status, application_type, approval_group, asset_tag, 
                assigned, assigned_to, assignment_group, attributes, business_owner, cpu, can_print, category, checked_in, checked_out, 
                comments, commission_date, company, correlation_id, cost_center, customer_vendor_agreements, dns_domain, dsl, date_disposed, 
                decommission_reason, default_gateway, department, description, disaster_recovery, discovery_source, due, due_in, esp_validated, 
                end_of_lease, environment, fault_count, first_discovered, fully_qualified_domain_name, gl_account, goods_received_date, hdd, 
                import_data, installed, invoice_number, justification, key_business_function, lease_review_date, lease_schedule_number, 
                lease_contract, licence_type, license_expiry, license_quantity, life_cycle_status, location, mac_address, maintenance_expiry, 
                maintenance_provider, maintenance_provider_lookup, make, managed_ci, managed_by, manufactured_by, manufacturer, memory, 
                metered_software, model_id, monitor, most_recent_discovery, name_2, operating_system, operational_status, order_received, 
                ordered, ownership, purchase_order, purchase_cost, purchase_date, purchased, ru_position, related_service_documentation, 
                sccm_package_name, status_not_used, subtype, subcategory, substatus, supplier, supplier_lookup, support_workgroup, 
                supported_by, tier_type, _type, user_authentication_required, user_url, vendor, ver_ctrl_repository, version_no, 
                warranty_expiry, warranty_expiration, sys_id ) VALUES 
                ( %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , 
                    %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , 
                    %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , 
                    %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , 
                    %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s , %s )
            """
    cur.executemany(sql, vals)
    conn.commit()
    conn.close()
'''
except (Exception, psycopg2.DatabaseError) as error:
    print(error)
finally:
    if conn is not None:
        conn.close()
'''