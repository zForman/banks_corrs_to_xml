import xlrd
from lxml import etree as et

xls_file_path = 'Active_EFTS_Banks.xlsx'
xls_file = 'EFTS_Banks.xls'


# Load data from excel file
def load_xls(xls_file):
    rb = xlrd.open_workbook(xls_file, formatting_info=True)
    xl_sheet = rb.sheet_by_index(0)

    set_rows = []

    # Pars all lines of a file
    for row_idx in range(xl_sheet.nrows):
        bic = str(xl_sheet.cell(row_idx, 0).value)
        bank = str(xl_sheet.cell(row_idx, 1).value)
        #     owner = str(xl_sheet.cell(row_idx, 2).value)
        #     switf_reg = str(xl_sheet.cell(row_idx, 3).value)
        #     participant_type_code = str(xl_sheet.cell(row_idx, 4).value)
        result = {
            'bic': bic,
            'institution': bank,
            # 'owner': owner,
            # 'switf_reg': switf_reg,
            # 'participant_type_code': participant_type_code
        }
        # print(result)
        set_rows.append(result)

    return set_rows


# load_xls(xls_file)

def create_xml(input_xls_file, output_xml_file_name='corrs.xml'):
    # Create XML tree in output_xml_file_name.xml
    root = et.Element('PIEMsg')
    data = et.SubElement(root, 'Data')
    set_rows = load_xls(input_xls_file)

    for row_item in set_rows:
        # print(row_item)
        field_names = et.SubElement(data, 'FieldNames')
        params = et.SubElement(data, 'Params')
        row = et.SubElement(data, 'Row')
        data_sub = et.SubElement(row, 'Data')
        bic = et.SubElement(row, 'BIC')
        name = et.SubElement(row, 'NAME')
        ext_code = et.SubElement(row, 'EXT_CODE')
        clearing_center_name = et.SubElement(row, 'CLEARING_CENTER_NAME')
        clearing_center_code = et.SubElement(row, 'CLEARING_CENTER_CODE')
        owner = et.SubElement(row, 'OWNER')
        sub_code = et.SubElement(row, 'SUB_CODE')
        ident_num = et.SubElement(row, 'IDENT_NUM')
        swift_reg = et.SubElement(row, 'SWIFT_REG')
        participant_type_code = et.SubElement(row, 'TYPE')

        bic.text = row_item['bic']
        name.text = row_item['institution']
#         owner.text = row_item['owner']
#         swift_reg.text = row_item['switf_reg']
#         participant_type_code.text = row_item['participant_type_code']
        my_tree = et.ElementTree(root)
        my_tree.write(output_xml_file_name, pretty_print=True, encoding="utf-8")


create_xml(xls_file)
