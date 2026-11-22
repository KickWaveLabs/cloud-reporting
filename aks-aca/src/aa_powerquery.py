#!/usr/bin/env python3
"""Load one Power Query query into a table of an .xlsx written by openpyxl, so Excel opens it with a real
connection (Data > Queries & Connections) and Data > Refresh All reads the source files again.

The query block (the DataMashup part) starts from the one Excel itself wrote for Microsoft's connected-workbooks
library (MIT license, https://github.com/microsoft/connected-workbooks), kept in
assets/powerquery/datamashup-template.b64: only its M code and the query name change, as that library does. The
connection, the query table and the table attributes follow the same library's template.

Usage: load_query(path, sheet_title, table_name, query_name, query_m, data_columns) after openpyxl saved the workbook.
The table must exist; its first columns must be data_columns, the columns the query returns, in the same order.
Columns after them are formula columns, which Excel keeps and fills on every refresh.
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


def data_mashup(query_name, query_m):
    """The template's DataMashup with Formulas/Section1.m replaced and its query renamed ([MS-QDEFF]: version,
    package parts, permissions, metadata, permission bindings; only the first and the metadata XML change)."""
    with open(TEMPLATE, encoding="ascii") as handle:
        blob = base64.b64decode(handle.read().strip())
    parts_length = struct.unpack("<I", blob[4:8])[0]
    parts = zipfile.ZipFile(io.BytesIO(blob[8:8 + parts_length]))
    section = f'section Section1;\n\nshared #"{query_name}" = {query_m};'
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
    xml = metadata[8:8 + xml_length].decode("utf-8")
    xml = xml.replace("Section1/Query1", f"Section1/{query_name}").replace('"ResultType" Value="sText"',
                                                                         '"ResultType" Value="sTable"')
    xml_bytes = xml.encode("utf-8")
    metadata = metadata[:4] + struct.pack("<I", len(xml_bytes)) + xml_bytes + metadata[8 + xml_length:]
    out = (blob[:4] + struct.pack("<I", len(new_parts)) + new_parts + permissions + struct.pack("<I", len(metadata))
           + metadata + bindings)
    return base64.b64encode(out).decode("ascii")


def load_query(path, sheet_title, table_name, query_name, query_m, data_columns):
    with zipfile.ZipFile(path) as source:
        files = {name: source.read(name) for name in source.namelist()}
    workbook = files["xl/workbook.xml"].decode("utf-8")
    sheets = re.findall(r'<sheet\b[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', workbook)
    rels = files["xl/_rels/workbook.xml.rels"].decode("utf-8")
    targets = dict(re.findall(r'Id="([^"]+)"[^>]*Target="/?(?:xl/)?([^"]+)"', rels))
    targets.update({rel_id: target for target, rel_id in re.findall(r'Target="/?(?:xl/)?([^"]+)"[^>]*Id="([^"]+)"', rels)})
    position = [title for title, _ in sheets].index(sheet_title)
    sheet_part = "xl/" + targets[sheets[position][1]]
    table_part = next(name for name in files if re.fullmatch(r"xl/tables/table\d+\.xml", name)
                      and f'displayName="{table_name}"' in files[name].decode("utf-8"))
    table = files[table_part].decode("utf-8")
    columns = re.findall(r'<tableColumn\b[^>]*\bid="(\d+)"[^>]*\bname="([^"]*)"[^>]*?(/>|>.*?</tableColumn>)', table,
                         flags=re.DOTALL)
    names = [name for _, name, _ in columns]
    if names[:len(data_columns)] != list(data_columns):
        raise ValueError(f"Table {table_name} does not start with the query columns: {names[:len(data_columns)]}")
    cells = re.search(r'\bref="([A-Z]+)(\d+):([A-Z]+)(\d+)"', table).groups()

    # The table becomes the query's table: every column gets a query field; formula columns are not data bound.
    table = table.replace("<table ", '<table tableType="queryTable" totalsRowShown="0" ', 1)
    if "<autoFilter" not in table:
        table = re.sub(r"(<tableColumns\b)", f'<autoFilter ref="{cells[0]}{cells[1]}:{cells[2]}{cells[3]}"/>\\1', table,
                       count=1)
    for column_id, _, _ in columns:
        table = re.sub(rf'(<tableColumn\b[^>]*?\bid="{column_id}")', rf'\1 uniqueName="{column_id}" '
                       rf'queryTableFieldId="{column_id}"', table, count=1)
    files[table_part] = table.encode("utf-8")
    fields = "".join(f'<queryTableField id="{column_id}" name="{escape(name)}" tableColumnId="{column_id}"/>'
                     if index < len(data_columns) else
                     f'<queryTableField id="{column_id}" dataBound="0" tableColumnId="{column_id}"/>'
                     for index, (column_id, name, _) in enumerate(columns))
    unbound = len(columns) - len(data_columns)
    files["xl/queryTables/queryTable1.xml"] = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<queryTable xmlns="{MAIN}" {XR16} '
        f'name="ExternalData_1" connectionId="1" xr16:uid="{guid()}" autoFormatId="16" applyNumberFormats="0" '
        'applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" '
        f'applyWidthHeightFormats="0" refreshOnLoad="0"><queryTableRefresh nextId="{len(columns) + 1}"'
        + (f' unboundColumnsRight="{unbound}"' if unbound else "")
        + f'><queryTableFields count="{len(columns)}">{fields}</queryTableFields></queryTableRefresh></queryTable>'
    ).encode("utf-8")
    table_rels = table_part.replace("xl/tables/", "xl/tables/_rels/") + ".rels"
    files[table_rels] = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="{PKG_REL}">'
                         f'<Relationship Id="rId1" Type="{REL}/queryTable" Target="../queryTables/queryTable1.xml"/>'
                         "</Relationships>").encode("utf-8")
    files["xl/connections.xml"] = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<connections xmlns="{MAIN}" {XR16}><connection '
        f'id="1" xr16:uid="{guid()}" keepAlive="1" name="Query - {query_name}" description="Connection to the '
        f"'{query_name}' query in the workbook.\" type=\"5\" refreshedVersion=\"7\" background=\"1\" saveData=\"1\">"
        f'<dbPr connection="Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=&quot;{query_name}'
        f'&quot;;" command="SELECT * FROM [{query_name}]" refreshOnLoad="0"/></connection></connections>'
    ).encode("utf-8")
    files["customXml/item1.xml"] = ('﻿<?xml version="1.0" encoding="utf-16"?><DataMashup xmlns="http://schemas.'
                                    f'microsoft.com/DataMashup">{data_mashup(query_name, query_m)}</DataMashup>'
                                    ).encode("utf-16-le")
    files["customXml/itemProps1.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n<ds:datastoreItem ds:itemID="' + guid() + '" '
        'xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs/></ds:datastoreItem>'
    ).encode("utf-8")
    files["customXml/_rels/item1.xml.rels"] = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="{PKG_REL}"><Relationship '
        f'Id="rId1" Type="{REL}/customXmlProps" Target="itemProps1.xml"/></Relationships>').encode("utf-8")

    name = (f'<definedName name="ExternalData_1" localSheetId="{position}" hidden="1">'
            f"{escape(sheet_title)}!${cells[0]}${cells[1]}:${cells[2]}${cells[3]}</definedName>")
    workbook = workbook.replace("<definedNames/>", "<definedNames></definedNames>")
    workbook = (workbook.replace("</definedNames>", name + "</definedNames>", 1) if "</definedNames>" in workbook
                else workbook.replace("</sheets>", f"</sheets><definedNames>{name}</definedNames>"))
    files["xl/workbook.xml"] = workbook.encode("utf-8")
    rels = rels.replace("</Relationships>", f'<Relationship Id="rIdConnections" Type="{REL}/connections" '
                                            'Target="connections.xml"/><Relationship Id="rIdMashup" '
                                            f'Type="{REL}/customXml" Target="../customXml/item1.xml"/></Relationships>')
    files["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")
    types = files["[Content_Types].xml"].decode("utf-8")
    types = types.replace("</Types>", (
        '<Override PartName="/xl/connections.xml" ContentType="application/vnd.openxmlformats-officedocument.'
        'spreadsheetml.connections+xml"/><Override PartName="/xl/queryTables/queryTable1.xml" ContentType="application/'
        'vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml"/><Override PartName="/customXml/itemProps1.xml"'
        ' ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/></Types>'))
    if 'Extension="xml"' not in types:
        types = types.replace("</Types>", '<Default Extension="xml" ContentType="application/xml"/></Types>')
    files["[Content_Types].xml"] = types.encode("utf-8")
    assert sheet_part in files
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as target:
        target.writestr("[Content_Types].xml", files.pop("[Content_Types].xml"))
        for item_name, data in files.items():
            target.writestr(item_name, data)
