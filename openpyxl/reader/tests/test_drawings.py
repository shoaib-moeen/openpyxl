# Copyright (c) 2010-2024 openpyxl

from io import BytesIO
from zipfile import ZipFile


def test_read_charts(datadir):
    datadir.chdir()

    archive = ZipFile("sample.xlsx")
    path = "xl/drawings/drawing1.xml"

    from ..drawings import find_images
    charts = find_images(archive, path)[0]
    assert len(charts) == 6


def test_read_drawing(datadir):
    datadir.chdir()

    archive = ZipFile("sample_with_images.xlsx")
    path = "xl/drawings/drawing1.xml"

    from ..drawings import find_images
    images = find_images(archive, path)[1]
    assert len(images) == 3


def test_unsupport_drawing(datadir):
    datadir.chdir()
    out = BytesIO()
    archive = ZipFile(out, mode="w")
    archive.write("unsupported_drawing.xml", "drawing1.xml")

    from ..drawings import find_images
    charts, images, shapes = find_images(archive, "drawing1.xml")
    assert charts == images == shapes == []


def test_unsupported_image_format(datadir):
    datadir.chdir()

    archive = ZipFile("sample_with_unsupported_image_format.xlsx", "r")
    path = "xl/drawings/drawing1.xml"

    from ..drawings import find_images
    images = find_images(archive, path)
    assert images == ([], [], [])


def test_hyperlink(datadir):
    datadir.chdir()

    archive = ZipFile("drawing_with_hyperlink.xlsx", "r")
    path = "xl/drawings/drawing1.xml"

    from ..drawings import find_images
    shapes = find_images(archive, path)[-1]

    assert shapes[0].nvSpPr.cNvPr.hlinkClick.target == "http://www.example.org"
