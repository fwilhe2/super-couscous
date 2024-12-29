import argparse
import json
import os
import xml.etree.ElementTree as ET
from xml.dom import minidom

# This script can create a Flat Open Document Spreadsheet (fods) file which can be opened and edited with LibreOffice Calc.
# The flat variant of ODS is nice because it is plain text which makes it easier to view/edit compared to the zipped versions.
# But also it's much nicer than importing csv files because fods files can have style and data type information.
# Currencies, dates, times, percentages etc can be properly represented without needing manual editing.
# See https://wiki.documentfoundation.org/Libreoffice_and_subversion for further information.

def create_fods_structure():
    # Not sure which of those namespaces are really needed. If some are omitted, libreoffice won't open the document.
    # There is probably a way to find out which are needed and reduce this set.
    root = ET.Element('office:document', {
        'xmlns:office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
        'xmlns:table': 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
        'xmlns:text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
        'xmlns:style': 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
        'xmlns:fo': 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
        'xmlns:svg': 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0',
        'xmlns:chart': 'urn:oasis:names:tc:opendocument:xmlns:chart:1.0',
        'xmlns:dr3d': 'urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0',
        'xmlns:math': 'http://www.w3.org/1998/Math/MathML',
        'xmlns:form': 'urn:oasis:names:tc:opendocument:xmlns:form:1.0',
        'xmlns:script': 'urn:oasis:names:tc:opendocument:xmlns:script:1.0',
        'xmlns:config': 'urn:oasis:names:tc:opendocument:xmlns:config:1.0',
        'xmlns:xlink': 'http://www.w3.org/1999/xlink',
        'xmlns:dc': 'http://purl.org/dc/elements/1.1/',
        'xmlns:meta': 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0',
        'xmlns:number': 'urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0',
        'xmlns:of': 'urn:oasis:names:tc:opendocument:xmlns:of:1.2',
        'xmlns:xforms': 'http://www.w3.org/2002/xforms',
        'xmlns:xsd': 'http://www.w3.org/2001/XMLSchema',
        'xmlns:xsi': 'http://www.w3.org/2001/XMLSchema-instance',
        'xmlns:grddl': 'http://www.w3.org/2003/g/data-view#',
        'xmlns:xhtml': 'http://www.w3.org/1999/xhtml',
        'xmlns:presentation': 'urn:oasis:names:tc:opendocument:xmlns:presentation:1.0',
        'xmlns:css3t': 'http://www.w3.org/TR/css3-text/',
        'xmlns:formx': 'urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0',
        'xmlns:oooc': 'http://openoffice.org/2004/calc',
        'xmlns:ooow': 'http://openoffice.org/2004/writer',
        'xmlns:rpt': 'http://openoffice.org/2005/report',
        'xmlns:draw': 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0',
        'xmlns:ooo': 'http://openoffice.org/2004/office',
        'xmlns:calcext': 'urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0',
        'xmlns:tableooo': 'http://openoffice.org/2009/table',
        'xmlns:drawooo': 'http://openoffice.org/2010/draw',
        'xmlns:loext': 'urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0',
        'xmlns:dom': 'http://www.w3.org/2001/xml-events',
        'xmlns:field': 'urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0',
        'office:version': '1.3',
        'office:mimetype': 'application/vnd.oasis.opendocument.spreadsheet'
    })

    automatic_styles = ET.SubElement(root, 'office:automatic-styles')

    add_default_styles_for_datatypes(automatic_styles)

    body = ET.SubElement(root, 'office:body')
    spreadsheet = ET.SubElement(body, 'office:spreadsheet')

    return root, spreadsheet

def add_default_styles_for_datatypes(automatic_styles):
    # I wish it was not needed to embed this into the document, but it looks like without those styles LibreOffice just does not know how to format those types properly.

    # FLOAT_STYLE
    float_style = ET.SubElement(automatic_styles, 'number:number-style', {'style:name': '___FLOAT_STYLE', 'style:volatile': 'true'})
    ET.SubElement(float_style, 'number:number', {'number:decimal-places': '2', 'number:min-decimal-places': '2', 'number:min-integer-digits': '1', 'number:grouping': 'true'})

    negative_number_float_style = ET.SubElement(automatic_styles, 'number:number-style', {'style:name': '__FLOAT_STYLE'})
    ET.SubElement(negative_number_float_style, 'style:text-properties', {'fo:color': '#ff0000'})
    ET.SubElement(negative_number_float_style, 'number:text').text = '-'
    ET.SubElement(negative_number_float_style, 'number:number', {'number:decimal-places': '2', 'number:min-decimal-places': '2', 'number:min-integer-digits': '1', 'number:grouping': 'true'})
    ET.SubElement(negative_number_float_style, 'style:map', {'style:condition': 'value()>=0', 'style:apply-style-name': '___FLOAT_STYLE'})

    ET.SubElement(automatic_styles, 'style:style', {'style:name': 'FLOAT_STYLE', 'style:family': 'table-cell', 'style:parent-style-name': 'Default', 'style:data-style-name': '__FLOAT_STYLE'})

    # DATE_STYLE
    date_style = ET.SubElement(automatic_styles, 'number:date-style', {'style:name': '__DATE_STYLE'})
    ET.SubElement(date_style, 'number:year', {'number:style': 'long'})
    ET.SubElement(date_style, 'number:text').text = '-'
    ET.SubElement(date_style, 'number:month', {'number:style': 'long'})
    ET.SubElement(date_style, 'number:text').text = '-'
    ET.SubElement(date_style, 'number:day', {'number:style': 'long'})

    ET.SubElement(automatic_styles, 'style:style', {'style:name': 'DATE_STYLE', 'style:family': 'table-cell', 'style:parent-style-name': 'Default', 'style:data-style-name': '__DATE_STYLE'})

    # TIME_STYLE
    time_style = ET.SubElement(automatic_styles, 'number:time-style', {'style:name': '__TIME_STYLE'})
    ET.SubElement(time_style, 'number:hours', {'number:style': 'long'})
    ET.SubElement(time_style, 'number:text').text = ':'
    ET.SubElement(time_style, 'number:minutes', {'number:style': 'long'})
    ET.SubElement(time_style, 'number:text').text = ':'
    ET.SubElement(time_style, 'number:seconds', {'number:style': 'long'})

    ET.SubElement(automatic_styles, 'style:style', {'style:name': 'TIME_STYLE', 'style:family': 'table-cell', 'style:parent-style-name': 'Default', 'style:data-style-name': '__TIME_STYLE'})

    # CURRENCY_STYLE
    currency_style = ET.SubElement(automatic_styles, 'number:currency-style', {'style:name': '___EUR_STYLE', 'style:volatile': 'true', 'number:language': 'en', 'number:country': 'DE'})
    ET.SubElement(currency_style, 'number:number', {'number:decimal-places': '2', 'number:min-decimal-places': '2', 'number:min-integer-digits': '1', 'number:grouping': 'true'})
    ET.SubElement(currency_style, 'number:text')
    ET.SubElement(currency_style, 'number:currency-symbol', {'number:language': 'de', 'number:country': 'DE'}).text = '€'

    negative_number_currency_style = ET.SubElement(automatic_styles, 'number:currency-style', {'style:name': '__EUR_STYLE', 'number:language': 'en', 'number:country': 'DE'})
    ET.SubElement(negative_number_currency_style, 'style:text-properties', {'fo:color': '#ff0000'})
    ET.SubElement(negative_number_currency_style, 'number:text').text = '-'
    ET.SubElement(negative_number_currency_style, 'number:number', {'number:decimal-places': '2', 'number:min-decimal-places': '2', 'number:min-integer-digits': '1', 'number:grouping': 'true'})
    ET.SubElement(negative_number_currency_style, 'number:text')
    ET.SubElement(negative_number_currency_style, 'number:currency-symbol', {'number:language': 'de', 'number:country': 'DE'}).text = '€'
    ET.SubElement(negative_number_currency_style, 'style:map', {'style:condition': 'value()>=0', 'style:apply-style-name': '___EUR_STYLE'})

    ET.SubElement(automatic_styles, 'style:style', {'style:name': 'EUR_STYLE', 'style:family': 'table-cell', 'style:parent-style-name': 'Default', 'style:data-style-name': '__EUR_STYLE'})

    # PERCENTAGE_STYLE
    percentage_style = ET.SubElement(automatic_styles, 'number:percentage-style', {'style:name': '__PERCENTAGE_STYLE'})
    ET.SubElement(percentage_style, 'number:number', {'number:decimal-places': '2', 'number:min-decimal-places': '2', 'number:min-integer-digits': '1'})
    ET.SubElement(percentage_style, 'number:text').text = '%'

    ET.SubElement(automatic_styles, 'style:style', {'style:name': 'PERCENTAGE_STYLE', 'style:family': 'table-cell', 'style:parent-style-name': 'Default', 'style:data-style-name': '__PERCENTAGE_STYLE'})

def add_content(spreadsheet_data, spreadsheet_element):
    table = ET.SubElement(spreadsheet_element, 'table:table', {'table:name': 'Sheet1'})

    for row in spreadsheet_data:
        table_row = ET.SubElement(table, 'table:table-row')
        for cell in row:
            if isinstance(cell, dict):
                value_type = cell.get("valueType", "string")
                value = cell.get("value", "")
                attributes = {
                    'office:value-type': value_type,
                    'calcext:value-type': value_type,
                }
                if value_type == 'float':
                    attributes['office:value'] = value
                    attributes['table:style-name'] = "FLOAT_STYLE"
                elif value_type == 'date':
                    attributes['office:date-value'] = value
                    attributes['table:style-name'] = "DATE_STYLE"
                elif value_type == 'time':
                    time = cell.get("value")
                    # assume hh:mm:ss format for now
                    components = time.split(":")
                    attributes['office:time-value'] = f"PT{components[0]}H{components[1]}M{components[2]}S"
                    attributes['table:style-name'] = "TIME_STYLE"
                elif value_type == 'currency':
                    attributes['office:value'] = value
                    attributes['office:currency'] = "EUR"  # Hardcoded for now
                    attributes['table:style-name'] = "EUR_STYLE"
                elif value_type == 'percentage':
                    attributes['office:value'] = value
                    attributes['table:style-name'] = "PERCENTAGE_STYLE"
                else:
                    attributes['office:value'] = value

                table_cell = ET.SubElement(table_row, 'table:table-cell', attributes)
                text_p = ET.SubElement(table_cell, 'text:p')
                text_p.text = value
            else:
                table_cell = ET.SubElement(table_row, 'table:table-cell', {
                    'office:value-type': 'string'
                })
                text_p = ET.SubElement(table_cell, 'text:p')
                text_p.text = cell

def save_fods_file(root, fods_file):
    fods_content = ET.tostring(root, encoding='utf-8', method='xml')

    pretty_fods_content = pretty_print(fods_content)

    with open(fods_file, 'w', encoding='utf-8') as f:
        f.write(pretty_fods_content)

def pretty_print(xml):
    return minidom.parseString(xml).toprettyxml(indent="  ")


def write_fods(spreadsheet, fods_file):
    root, spreadsheet_element = create_fods_structure()
    add_content(spreadsheet, spreadsheet_element)
    save_fods_file(root, fods_file)


def parse_json_file(input_file):
    if not os.path.isfile(input_file):
        raise(Exception(f"File not found: {input_file}"))

    with open(input_file, 'r') as file:
        try:
            data = json.load(file)
            return data
        except json.JSONDecodeError as err:
            raise(Exception(f"Error decoding JSON: {err}"))


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Flat ODS writer.')
    parser.add_argument('input_file', type=str, help='Path to the JSON file containing the spreadsheet data')
    parser.add_argument('output_file', type=str, help='Filename of the output file, should end in ".fods"')

    args = parser.parse_args()

    spreadsheet = parse_json_file(args.input_file)
    write_fods(spreadsheet, args.output_file)
