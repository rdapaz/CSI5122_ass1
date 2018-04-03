import re
import pprint

def pretty_printer(o):
    pp = pprint.PrettyPrinter(indent=4)
    pp.pprint(o)


data = """
Name
Status
Description
References
Phase
Votes
Comments
""".splitlines()

def replace_spaces(sText):
    rex = re.compile('\s+')
    sText = re.sub(r'\s+\Z', '', sText)
    sText = re.sub(r'\A\s+', '', sText)
    sText = rex.sub(repl='_', string=sText)
    sText = sText.replace('(', '').replace(')', '').replace('-', '').replace('/','_').replace('__','_').replace('#', 'no').replace('.','')
    sText = re.sub(r'[_]+', '_', sText)
    sText = sText.lower()
    return sText

data = [replace_spaces(x) for x in data if len(x) > 0]

for item in data:
    print(item)