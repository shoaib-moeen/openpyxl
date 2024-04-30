# Copyright (c) 2010-2024 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

from ..ole import AnchorMarker


@pytest.fixture
def ObjectAnchor():
    from ..ole import ObjectAnchor
    return ObjectAnchor


class TestObjectAnchor:

    def test_ctor(self, ObjectAnchor):
        _from = AnchorMarker()
        to = AnchorMarker()
        anchor = ObjectAnchor(_from=_from, to=to)
        xml = tostring(anchor.to_tree())
        expected = """
        <anchor moveWithCells="0" sizeWithCells="0"
        xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">
          <from>
            <xdr:col>0</xdr:col>
            <xdr:colOff>0</xdr:colOff>
            <xdr:row>0</xdr:row>
            <xdr:rowOff>0</xdr:rowOff>
          </from>
          <to>
            <xdr:col>0</xdr:col>
            <xdr:colOff>0</xdr:colOff>
            <xdr:row>0</xdr:row>
            <xdr:rowOff>0</xdr:rowOff>
          </to>
        </anchor>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, ObjectAnchor):
        src = """
        <anchor moveWithCells="0" sizeWithCells="0">
          <from>
            <col>0</col>
            <colOff>0</colOff>
            <row>0</row>
            <rowOff>0</rowOff>
          </from>
          <to>
            <col>0</col>
            <colOff>0</colOff>
            <row>0</row>
            <rowOff>0</rowOff>
          </to>
        </anchor>
        """
        node = fromstring(src)
        _from = AnchorMarker()
        to = AnchorMarker()
        a1 = ObjectAnchor(_from=_from, to=to)
        a2 = ObjectAnchor.from_tree(node)
        assert a1 == a2
