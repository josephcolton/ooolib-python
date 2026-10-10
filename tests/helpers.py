#################################################################################
# ooolib-python - Python module for creating Open Document Format documents.    #
# Copyright (C) 2006-2026  Joseph Colton                                        #
#                                                                               #
# You can contact me by email at josephcolton@gmail.com                         #
#################################################################################
"""Shared helpers for the ooolib-python tests.

Run every test from the repository root with:
    python3 -m unittest discover tests
"""

# Import python modules
import os
import shutil
import struct
import sys
import tempfile
import unittest
import zipfile
import xml.dom.minidom
import zlib

# Make the library importable
TESTS_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_DIR = os.path.dirname(TESTS_DIR)
CODE_DIR = os.path.join(REPO_DIR, "code")
EXAMPLES_DIR = os.path.join(REPO_DIR, "examples")
if CODE_DIR not in sys.path:
    sys.path.insert(0, CODE_DIR)

import ooolib

MIMETYPES = {
    ".ods": "application/vnd.oasis.opendocument.spreadsheet",
    ".odt": "application/vnd.oasis.opendocument.text",
    ".odp": "application/vnd.oasis.opendocument.presentation",
    ".odg": "application/vnd.oasis.opendocument.graphics",
}

def makePNG(width, height, dpi=None):
    """Return the bytes of a small solid PNG image, with an optional resolution."""
    def chunk(chunkType, data):
        return (struct.pack(">I", len(data)) + chunkType + data +
                struct.pack(">I", zlib.crc32(chunkType + data) & 0xffffffff))
    rows = (b"\x00" + b"\x44\x72\xc4" * width) * height
    png = b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
    if dpi is not None:
        pixelsPerMeter = int(round(dpi / 0.0254))
        png += chunk(b"pHYs", struct.pack(">IIB", pixelsPerMeter, pixelsPerMeter, 1))
    return png + chunk(b"IDAT", zlib.compress(rows)) + chunk(b"IEND", b"")

SVG = (b'<svg xmlns="http://www.w3.org/2000/svg" width="2in" height="1in" viewBox="0 0 200 100">'
       b'<rect width="200" height="100" fill="#ed7d31"/></svg>')

class DocumentTestCase(unittest.TestCase):
    """Test case with a temporary directory and ODF file checks."""
    def setUp(self):
        self.tempDir = tempfile.mkdtemp(prefix="ooolib-test-")

    def tearDown(self):
        shutil.rmtree(self.tempDir, ignore_errors=True)

    def path(self, filename):
        return os.path.join(self.tempDir, filename)

    def export(self, doc, filename):
        """Export doc into the temporary directory, check the file, and return its path."""
        filename = self.path(filename)
        doc.export(filename)
        self.assertValidODF(filename)
        return filename

    def assertValidODF(self, filename):
        """The file must be a zip with an uncompressed mimetype first and well formed XML files."""
        with zipfile.ZipFile(filename) as archive:
            first = archive.infolist()[0]
            self.assertEqual(first.filename, "mimetype")
            self.assertEqual(first.compress_type, zipfile.ZIP_STORED)
            extension = os.path.splitext(filename)[1]
            if extension in MIMETYPES:
                self.assertEqual(archive.read("mimetype").decode("UTF-8"), MIMETYPES[extension])
            for name in archive.namelist():
                if name.endswith(".xml"):
                    try:
                        xml.dom.minidom.parseString(archive.read(name))
                    except Exception as error:
                        self.fail("%s: %s is not well formed XML: %s" % (filename, name, error))

    def readZipString(self, filename, name):
        with zipfile.ZipFile(filename) as archive:
            return archive.read(name).decode("UTF-8")
