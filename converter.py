import pandas
from lxml import etree
import math
import datetime
def xlsx_to_xml_simple_mappings():
    return {
    'userid':'userid',
    'password':'GrpID',
    'first_name':'FirstName',
    'last_name':'LastName',
    'addr1':'Addr1',
    'addr2':'Addr2',
    'addr3':'Addr3',
    'city':'City',
    'state':'State',
    'birth_dt':'DOB',
    'sex':'Gender',
    }

def schema_location():
    return 'filename.xsd'

def validate_schema(xml_contents):
    with open('filename.xsd', 'r') as f:
        schema_string = f.read()
    schema_root = etree.XML(schema_string)
    schema = etree.XMLSchema(schema_root)
    parser = etree.XMLParser(schema=schema)
    root = etree.XML(xml_contents, parser)

def convert_to_xml(input:str):
    xls = pandas.ExcelFile(input,convert_float = False,dtypes={'zip4':str,'zip5':str,'conf_mailto_name':str})
    flat_array = xls.parse(xls.sheet_names[0]).transpose().to_dict()
    flat_xlsx_array = []
    for k,record in flat_array.items():
        record = format_xlsx_record(record)
        flat_xlsx_array.append(xlsx_record_to_xml_record(record))
    unflattened_array = unflatten_array(flat_xlsx_array,{'name':'Member_Details','properties':['LastName','FirstName','Gender','DOB']})

    xsi = 'http://www.w3.org/2001/XMLSchema-instance'
    root = etree.Element("filename", nsmap={'xsi': xsi})
    root.attrib[etree.QName(xsi, 'noNamespaceSchemaLocation')] = schema_location()

    for key,subscriber in unflattened_array.items():
        subscriber_node = etree.SubElement(root,'filename_Letter')
        #TODO: Add the rest of the attributes
        for k,v in subscriber.items():
            if k != 'Member_Details':
                subscriber_detail_node = etree.SubElement(subscriber_node,k)
                subscriber_detail_node.text = v



        members_node = etree.SubElement(subscriber_node,'Member_Details')
        for member in subscriber['Member_Details']:
            member_node = etree.SubElement(members_node,'mbr')
            for k,v in member.items():
                member_detail_node = etree.SubElement(member_node,k)
                member_detail_node.text = v
    contents = etree.tostring(root,encoding="utf-8",pretty_print=True,xml_declaration=True)
    validate_schema(contents)
    return contents


''' Given a flattened array (an array that was designed with object composition in mind but flattened to fit a 2d table/array)
will return a dict with the converted objects (as defined by flattened object)
In addition, duplicate items will be removed'''
def unflatten_array(flat_array,flattened_object):
    unflattened_array = {}
    for element in flat_array:
        unflattened_object = {}
        for k in flattened_object['properties']:
            unflattened_object[k] = element[k]

        #the index of each item will be a hash of the full element without the object in order to create a smaller set of unique items.
        columns = element.items()
        for k in flattened_object['properties']:
            del element[k]
        id = hash(frozenset(element.items()))

        if id not in unflattened_array:
            unflattened_array[id] = element
            unflattened_array[id][flattened_object['name']] = [unflattened_object]
        else:
            unflattened_array[id][flattened_object['name']].append(unflattened_object)
    return unflattened_array


def read_excel(input:str):
    xls = pandas.ExcelFile(input)
    return xls.parse(xls.sheet_names[0])

def format_xlsx_record(xlsx_values):
    empty_keys = []
    for k,v in xlsx_values.items():
        if v == '' or (isinstance(v,float) and math.isnan(v)) or v is None:
            v = ''
        else:
            if isinstance(v,pandas._libs.tslibs.timestamps.Timestamp):
                v = v.strftime('%d-%m-%Y')

            if isinstance(v,float):
                v = int(v)
            if isinstance(v,int):
                v = str(v)
                if 'zip5' == k:
                    v = v.zfill(5)
                elif 'zip4' == k:
                    v = v.zfill(4)
                elif 'sub_ssn' == k:
                    v = v.zfill(9)
                elif 'sub_id' == k:
                    v = v.zfill(9)
                elif 'mbr_ssn' == k:
                    v = v.zfill(9)

        if not isinstance(v,str):
            raise ('Unexpected type ' + str(type(v)) + 'in record ' + str(xlsx_values))
            #xlsx_values[k] = str(v)
        xlsx_values[k] = v

    for ek in empty_keys:
        del xlsx_values[ek]

    return xlsx_values

def xlsx_record_to_xml_record(xlsx_values):
    """
    Populate values in xml at record level.

    Args:
        xlsx_values (dict): Excel file values as dictionary.
        xml_values (dict): XML file values as dictionary.

    Returns:
        Dict: Xml values.
    """
    xml_values = {}
    for xlsx_key, xml_key in xlsx_to_xml_simple_mappings().items():
        xml_values[xml_key] = xlsx_values[xlsx_key]

    xml_values['Zip'] = get_zip(xlsx_values)
    xml_values["LastName"] = get__last_name(xlsx_values)
    return xml_values

def get_zip(xlsx_values):
    if  xlsx_values['zip4'] != '':
        xlsx_values['zip4'] = ' ' + xlsx_values['zip4']

    return str(xlsx_values['zip5']) + str(xlsx_values['zip4'])


def get__last_name(xlsx_values):
    """
    Get "LastName" value for xml.

    If conf_ind = 'Y', use conf_mailto_name.
        'LastName'
    Returns:
        String: LastName value.
    """
    if xlsx_values['conf_ind'] =='Y':
        return xlsx_values['conf_mailto_name']

    return xlsx_values['last_name']

