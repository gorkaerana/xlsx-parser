
from pprint import pprint
import xml.etree.ElementTree as ET
from zipfile import ZipFile


TAG_PREFIX = ("{http://schemas.openxmlformats.org/"
              "spreadsheetml/2006/main}")

def parse_xlsx(filepath, sheet_name):
    """
    Parses the tabular data contained in sheet 'sheet_name' of .xlsx
    file saved in 'filepath'.
    Arguments:
    filepath
    sheet_name

    Output:
    table
    """
    table = []
    with ZipFile(filepath, 'r') as zip_file:
        xml_filepath = f"xl/worksheets/{sheet_name}.xml"
        with zip_file.open(xml_filepath) as xml_file:
            sheet_data = ET.parse(xml_file)\
                           .getroot()\
                           .find(f'{TAG_PREFIX}sheetData')
            for row in sheet_data.findall(f'{TAG_PREFIX}row'):
                row_values = [
                    v.text
                    for cell in row.iter()
                    if (v := cell.find(f'{TAG_PREFIX}v')) is not None
                ]
                table.append(row_values)
    return table


if __name__ == '__main__':
    pprint(parse_xlsx('test_excel.zip', 'sheet1'))
