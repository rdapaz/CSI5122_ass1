"""
Uses and tested under Python 3.6
"""

from collections import OrderedDict
import yaml
import yaml.constructor
import win32com.client
import datetime


class OrderedDictYAMLLoader(yaml.Loader):
    """
    A YAML loader that loads mappings into ordered dictionaries.
    """

    def __init__(self, *args, **kwargs):
        yaml.Loader.__init__(self, *args, **kwargs)

        self.add_constructor(u'tag:yaml.org,2002:map', type(self).construct_yaml_map)
        self.add_constructor(u'tag:yaml.org,2002:omap', type(self).construct_yaml_map)

    def construct_yaml_map(self, node):
        data = OrderedDict()
        yield data
        value = self.construct_mapping(node)
        data.update(value)

    def construct_mapping(self, node, deep=False):
        if isinstance(node, yaml.MappingNode):
            self.flatten_mapping(node)
        else:
            raise yaml.constructor.ConstructorError(None, None,
                'expected a mapping node, but found %s' % node.id, node.start_mark)

        mapping = OrderedDict()
        for key_node, value_node in node.value:
            key = self.construct_object(key_node, deep=deep)
            try:
                hash(key)
            except TypeError:
                raise yaml.constructor.ConstructorError('while constructing a mapping',
                    node.start_mark, 'found unacceptable key (%s)' % exc, key_node.start_mark)
            value = self.construct_object(value_node, deep=deep)
            mapping[key] = value
        return mapping


def isDateLike(string):
    from dateutil.parser import parse
    try: 
        parse(string)
        return True
    except ValueError:
        return False


class MicrosoftProject:
    def __init__(self, doc):
        self.__file = doc
        self.__app = win32com.client.Dispatch('MSProject.Application')
        self.__app.Visible = True
        self.__doc = self.__app.FileOpen(self.__file)
        self.__proj = self.__app.ActiveProject
        self.__task_ids = []

    def add_new_auto_task(self, task_name, nesting, durn=None, resources=None):
        tsk = self.__proj.Tasks.Add(Name=task_name)
        nesting += 1
        tsk.Manual = False
        if durn:
            tsk.Duration = durn
        if resources:
            tsk.ResourceNames = resources
        tsk.Text1 = nesting
        if not (len(self.__task_ids) == 0 or self.__task_ids[-1][1] != nesting):
            tsk.Predecessors = self.__task_ids[-1][0]
        
        while int(tsk.OutlineLevel) < int(tsk.Text1):
            tsk.OutlineIndent()
        while int(tsk.OutlineLevel) > int(tsk.Text1):
            tsk.OutlineOutdent()
        
        self.__task_ids.append([tsk.ID, nesting])

    def add_new_manual_task(self, task_name, nesting, stt, fin):
        tsk = self.__proj.Tasks.Add(Name=task_name)
        nesting += 1
        tsk.Manual = True 
        tsk.Text1 = nesting
        tsk.Start = stt
        tsk.Finish = fin

        while int(tsk.OutlineLevel) < int(tsk.Text1):
            tsk.OutlineIndent()
        while int(tsk.OutlineLevel) > int(tsk.Text1):
            tsk.OutlineOutdent()

    def add_new_summary_task(self, task_name, nesting):
        tsk = self.__proj.Tasks.Add(Name=task_name)
        nesting += 1
        tsk.Manual = False 
        print(f"Nesting={nesting}")
        tsk.Text1 = nesting
        print(tsk.Text1)

        while int(tsk.OutlineLevel) < int(tsk.Text1):
            tsk.OutlineIndent()
        while int(tsk.OutlineLevel) > int(tsk.Text1):
            tsk.OutlineOutdent()
              
    def yaml_to_gantt(self, obj, nesting=0):
        for task, rest in obj.items():
            if type(rest) == OrderedDict:
                self.add_new_summary_task(task_name=task, 
                                  nesting=nesting
                                  )
                self.yaml_to_gantt(rest, nesting+1)
            else:
                val = rest.split('|', 2)[0]
                # print(val, len(val), sep="|")
                if not isDateLike(val):  
                  durn, resources = rest.split('|', 2)
                  self.add_new_auto_task(task_name=task, 
                                    durn=durn,
                                    resources=resources,
                                    nesting=nesting
                                   )
                else:
                  stt, fin = rest.split('|', 2)
                  # print(stt, fin, sep="|")
                  # stt = datetime.datetime.strptime(stt, '%d/%m/%Y')
                  # fin = datetime.datetime.strptime(stt, '%d/%m/%Y')
                  self.add_new_manual_task(task_name=task, 
                                    stt=stt,
                                    fin=fin,
                                    nesting=nesting
                                   )

def main(o):
    doc = r"C:\Users\rdapaz\Desktop\SIP Close Out of Legacy Projects - Copy.mpp"
    pj = MicrosoftProject(doc)
    pj.yaml_to_gantt(o)
    
"""
def test(o, nesting=0):
    SPACES = '   '
    for task, rest in o.items():
        if type(rest) == OrderedDict:
            print('{}{}'.format(SPACES*nesting, task))
            test(rest, nesting+1)
        else:
            print(": ".join(['{}{}'.format(SPACES*nesting, task), rest]))
"""

if __name__ == "__main__":
    with open(r'E:\work\I0441177 - Network EoL Refresh FY1617\tasks.yaml', 'r') as f:
        data = yaml.load(f, Loader=OrderedDictYAMLLoader)
    # obj = yaml.load(data, Loader=OrderedDictYAMLLoader)
    # test(data)
    main(data)