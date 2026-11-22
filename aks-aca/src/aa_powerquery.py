#!/usr/bin/env python3
"""Load one Power Query query into a table of an .xlsx written by openpyxl, so Excel opens it with a real
connection (Data > Queries & Connections) and Data > Refresh All reads the source files again.

The query block (the DataMashup part) starts from the one Excel itself wrote for Microsoft's connected-workbooks
library (MIT license, https://github.com/microsoft/connected-workbooks), kept in
assets/powerquery/datamashup-template.b64: only its M code and the query name change, as that library does. The
connection, the query table and the table attributes follow the same library's template.

Usage: load_queries(path, [(sheet_title, table_name, query_name, query_m, data_columns), ...]) after openpyxl saved
the workbook. Each table must exist and hold exactly data_columns, the columns its query returns, in the same order
(Excel drops a query table that also has formula columns: keep formulas in another table).
"""

import base64
import io
import os
import re
import struct
import uuid
import zipfile
from xml.sax.saxutils import escape

import aa_common as aa

TEMPLATE = os.path.join(aa.TOPIC_DIR, "assets", "powerquery", "datamashup-template.b64")
MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
XR16 = ('xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr16" '
        'xmlns:xr16="http://schemas.microsoft.com/office/spreadsheetml/2017/revision16"')


def guid():
    return "{" + str(uuid.uuid4()).upper() + "}"


def data_mashup(queries):
    """The template's DataMashup with Formulas/Section1.m holding the given (name, M code) queries, and one metadata
    item per query, copied from the template's ([MS-QDEFF]: version, package parts, permissions, metadata, permission
    bindings; only the package parts and the metadata XML change)."""
    with open(TEMPLATE, encoding="ascii") as handle:
        blob = base64.b64decode(handle.read().strip())
    parts_length = struct.unpack("<I", blob[4:8])[0]
    parts = zipfile.ZipFile(io.BytesIO(blob[8:8 + parts_length]))
    section = "section Section1;\n\n" + "\n\n".join(f'shared #"{name}" = {m};' for name, m in queries)
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as package:
        for item in parts.infolist():
            data = section.encode("utf-8") if item.filename == "Formulas/Section1.m" else parts.read(item.filename)
            package.writestr(item, data, compress_type=zipfile.ZIP_DEFLATED)
    new_parts = buffer.getvalue()
    rest = blob[8 + parts_length:]
    permissions_length = struct.unpack("<I", rest[:4])[0]
    permissions, rest = rest[:4 + permissions_length], rest[4 + permissions_length:]
    metadata_length = struct.unpack("<I", rest[:4])[0]
    metadata, bindings = rest[4:4 + metadata_length], rest[4 + metadata_length:]
    xml_length = struct.unpack("<I", metadata[4:8])[0]
    xml = metadata[8:8 + xml_length].decode("utf-8").replace('"ResultType" Value="sText"', '"ResultType" Value="sTable"')
    # The template's items for its one query (Section1/Query1 and its Source step), repeated for each query.
    items = re.findall(r"<Item><ItemLocation><ItemType>Formula</ItemType><ItemPath>Section1/Query1(?:/[^<]*)?</ItemPath>"
                       r".*?</Item>", xml, flags=re.DOTALL)
    block = "".join(items)
    xml = xml.replace(block, "".join(block.replace("Section1/Query1", f"Section1/{name}") for name, _ in queries))
    xml_bytes = xml.encode("utf-8")
    metadata = metadata[:4] + struct.pack("<I", len(xml_bytes)) + xml_bytes + metadata[8 + xml_length:]
    out = (blob[:4] + struct.pack("<I", len(new_parts)) + new_parts + permissions + struct.pack("<I", len(metadata))
           + metadata + bindings)
    return base64.b64encode(out).decode("ascii")


def load_query(path, sheet_title, table_name, query_name, query_m, data_columns):
    load_queries(path, [(sheet_title, table_name, query_name, query_m, data_columns)])


def load_queries(path, queries):
    """queries: (sheet title, table name, query name, M code, data columns) for each query loaded to a table."""
    with zipfile.ZipFile(path) as source:
        files = {name: source.read(name) for name in source.namelist()}
    workbook = files["xl/workbook.xml"].decode("utf-8")
    sheets = [title for title, _ in re.findall(r'<sheet\b[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', workbook)]
    rels = files["xl/_rels/workbook.xml.rels"].decode("utf-8")
    types = files["[Content_Types].xml"].decode("utf-8")
    connections, names = [], []
    for index, (sheet_title, table_name, query_name, _, data_columns) in enumerate(queries, start=1):
        table_part = next(name for name in files if re.fullmatch(r"xl/tables/table\d+\.xml", name)
                          and f'displayName="{table_name}"' in files[name].decode("utf-8"))
        table = files[table_part].decode("utf-8")
        columns = re.findall(r'<tableColumn\b[^>]*\bid="(\d+)"[^>]*\bname="([^"]*)"[^>]*?(/>|>.*?</tableColumn>)',
                             table, flags=re.DOTALL)
        if [name for _, name, _ in columns][:len(data_columns)] != list(data_columns):
            raise ValueError(f"Table {table_name} does not start with the columns of query {query_name}")
        cells = re.search(r'\bref="([A-Z]+)(\d+):([A-Z]+)(\d+)"', table).groups()
        # The table becomes the query's table: every column gets a query field; formula columns are not data bound.
        table = table.replace("<table ", '<table tableType="queryTable" totalsRowShown="0" ', 1)
        if "<autoFilter" not in table:
            table = re.sub(r"(<tableColumns\b)", f'<autoFilter ref="{cells[0]}{cells[1]}:{cells[2]}{cells[3]}"/>\\1',
                           table, count=1)
        for column_id, _, _ in columns:
            table = re.sub(rf'(<tableColumn\b[^>]*?\bid="{column_id}")', rf'\1 uniqueName="{column_id}" '
                           rf'queryTableFieldId="{column_id}"', table, count=1)
        files[table_part] = table.encode("utf-8")
        fields = "".join(f'<queryTableField id="{column_id}" name="{escape(name)}" tableColumnId="{column_id}"/>'
                         if position < len(data_columns) else
                         f'<queryTableField id="{column_id}" dataBound="0" tableColumnId="{column_id}"/>'
                         for position, (column_id, name, _) in enumerate(columns))
        unbound = len(columns) - len(data_columns)
        query_table = f"xl/queryTables/queryTable{index}.xml"
        files[query_table] = (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<queryTable xmlns="{MAIN}" {XR16} '
            f'name="ExternalData_{index}" connectionId="{index}" xr16:uid="{guid()}" autoFormatId="16" '
            'applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" '
            f'applyAlignmentFormats="0" applyWidthHeightFormats="0" refreshOnLoad="0"><queryTableRefresh '
            f'nextId="{len(columns) + 1}"' + (f' unboundColumnsRight="{unbound}"' if unbound else "")
            + f'><queryTableFields count="{len(columns)}">{fields}</queryTableFields></queryTableRefresh></queryTable>'
        ).encode("utf-8")
        files[table_part.replace("xl/tables/", "xl/tables/_rels/") + ".rels"] = (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="{PKG_REL}"><Relationship '
            f'Id="rId1" Type="{REL}/queryTable" Target="../queryTables/queryTable{index}.xml"/></Relationships>'
        ).encode("utf-8")
        types = types.replace("</Types>", f'<Override PartName="/{query_table}" ContentType="application/vnd.'
                                          'openxmlformats-officedocument.spreadsheetml.queryTable+xml"/></Types>')
        connections.append(
            f'<connection id="{index}" xr16:uid="{guid()}" keepAlive="1" name="Query - {query_name}" description='
            f'"Connection to the \'{query_name}\' query in the workbook." type="5" refreshedVersion="7" background="1" '
            f'saveData="1"><dbPr connection="Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=&quot;'
            f'{query_name}&quot;;" command="SELECT * FROM [{query_name}]" refreshOnLoad="0"/></connection>')
        names.append(f'<definedName name="ExternalData_{index}" localSheetId="{sheets.index(sheet_title)}" hidden="1">'
                     f"{escape(sheet_title)}!${cells[0]}${cells[1]}:${cells[2]}${cells[3]}</definedName>")
    files["xl/connections.xml"] = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<connections '
                                   f'xmlns="{MAIN}" {XR16}>{"".join(connections)}</connections>').encode("utf-8")
    mashup = data_mashup([(query_name, query_m) for _, _, query_name, query_m, _ in queries])
    files["customXml/item1.xml"] = ('\ufeff<?xml version="1.0" encoding="utf-16"?><DataMashup xmlns="http://schemas.'
                                    f'microsoft.com/DataMashup">{mashup}</DataMashup>').encode("utf-16-le")
    files["customXml/itemProps1.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n<ds:datastoreItem ds:itemID="' + guid() + '" '
        'xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs/></ds:datastoreItem>'
    ).encode("utf-8")
    files["customXml/_rels/item1.xml.rels"] = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="{PKG_REL}"><Relationship '
        f'Id="rId1" Type="{REL}/customXmlProps" Target="itemProps1.xml"/></Relationships>').encode("utf-8")
    workbook = workbook.replace("<definedNames/>", "<definedNames></definedNames>")
    workbook = (workbook.replace("</definedNames>", "".join(names) + "</definedNames>", 1) if "</definedNames>" in workbook
                else workbook.replace("</sheets>", f"</sheets><definedNames>{''.join(names)}</definedNames>"))
    files["xl/workbook.xml"] = workbook.encode("utf-8")
    rels = rels.replace("</Relationships>", f'<Relationship Id="rIdConnections" Type="{REL}/connections" '
                                            'Target="connections.xml"/><Relationship Id="rIdMashup" '
                                            f'Type="{REL}/customXml" Target="../customXml/item1.xml"/></Relationships>')
    files["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")
    types = types.replace("</Types>", (
        '<Override PartName="/xl/connections.xml" ContentType="application/vnd.openxmlformats-officedocument.'
        'spreadsheetml.connections+xml"/><Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.'
        'openxmlformats-officedocument.customXmlProperties+xml"/></Types>'))
    if 'Extension="xml"' not in types:
        types = types.replace("</Types>", '<Default Extension="xml" ContentType="application/xml"/></Types>')
    files["[Content_Types].xml"] = types.encode("utf-8")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as target:
        target.writestr("[Content_Types].xml", files.pop("[Content_Types].xml"))
        for item_name, data in files.items():
            target.writestr(item_name, data)
