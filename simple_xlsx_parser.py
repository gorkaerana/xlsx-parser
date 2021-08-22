
import re
from pprint import pprint
import xml.etree.ElementTree as ET
from zipfile import ZipFile


__all__ = ['parse_xlsx']


PREFIXES = {
    'xmlns': ('{http://schemas.openxmlformats.org/'
              'spreadsheetml/2006/main}'),
    'r': ('{http://schemas.openxmlformats.org/'
          'officeDocument/2006/relationships}')
    }


def find_files_containing_string(filepath, string):
    """
    I couldn't quite inmediately figure out how to grep within the
    files contained in a zipped directory using the UNIX command
    line. So I wrote this method.
    """
    matches_per_file = {}
    with ZipFile(filepath, 'r') as zip_file:
        for file_name in zip_file.namelist():
            with zip_file.open(file_name) as fp:
                content = fp.read()
                matches = re.findall(str.encode(string), content)
                matches_per_file[file_name] = matches
    return matches


def find_sheet_id_given_name(filepath, sheet_name):
    """
    In Excel files the data for each sheet is contained in a file
    named 'xl/worksheets/sheet{sheet_id}.xml', where the mapping
    between the 'sheet_name' displayed in Excel and the sheet_id
    is stored in 'xl/workbook.xml'. More specifically, under tag
    'sheets', there is a tag called 'sheet' (one for each of them);
    where attribute 'name' contains sheet_name and attribute
    'name' contains the sheet_id.
    """
    with ZipFile(filepath, 'r') as zip_file:
        with zip_file.open('xl/workbook.xml') as workbook_file:
            sheets = ET.parse(workbook_file)\
                       .find(f"{PREFIXES['xmlns']}sheets")
            for sheet in sheets:
                attribs = sheet.attrib
                if attribs.get('name') == sheet_name:
                    return attribs.get(f"{PREFIXES['r']}id", [])[3:]
    raise ValueError(
        f"Excel file '{filepath}' does not contain sheet named"
        f" '{sheet_name}'."
    )


def extract_table_from_xml_file(zip_filepath, xml_filepath):
    """
    """
    table = []
    with ZipFile(zip_filepath, 'r') as zip_file:
        with zip_file.open(xml_filepath) as xml_file:
            # The data is contained under tag 'sheetData' and the
            # values of the cells of each row under tag 'row'
            sheet_data = ET.parse(xml_file)\
                           .getroot()\
                           .find(f"{PREFIXES['xmlns']}sheetData")\
                           .findall(f"{PREFIXES['xmlns']}row")
            # Fetch cell values from tag 'v' and attribute 'text'
            for row in sheet_data:
                row_values = []
                for cell in row.iter():
                    value = cell.find(f"{PREFIXES['xmlns']}v")
                    if value is not None:
                        row_values.append(value.text)
                table.append(row_values)
    return table

def parse_xlsx(filepath, sheet_name=None):
    """
    Parses the tabular data contained in sheet 'sheet_name' of .xlsx
    file saved in 'filepath'.
    Arguments:
    filepath
    sheet_name

    Output:
    table
    """
    # The sheet_name -> sheet_id mapping ought to be extracted from
    # 'xl/workbook.xml'
    if sheet_name is None:
        sheet_id = '1'
    else:
        sheet_id = find_sheet_id_given_name(filepath, sheet_name)
    xml_filepath = f"xl/worksheets/sheet{sheet_id}.xml"
    # The table data is to be extracted from xml_filepath
    return extract_table_from_xml_file(filepath, xml_filepath)


if __name__ == '__main__':
    pprint(parse_xlsx('test_excel.zip', 'superloen'))
