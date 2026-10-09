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
import sys
import tempfile
import unittest
import zipfile
import xml.dom.minidom

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
