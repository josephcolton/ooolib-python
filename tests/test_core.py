#################################################################################
# ooolib-python - Python module for creating Open Document Format documents.    #
# Copyright (C) 2006-2026  Joseph Colton                                        #
#                                                                               #
# You can contact me by email at josephcolton@gmail.com                         #
#################################################################################
"""Tests for the XML, unit, and file helpers."""

# Import python modules
import unittest

# Import test helpers (also puts code/ on the path)
import helpers
import ooolibFile
import ooolibUnits
import ooolibXML

class XMLEscapingTest(unittest.TestCase):
    def test_escapeText(self):
        self.assertEqual(ooolibXML.escapeText("Tom & <Jerry>"), "Tom &amp; &lt;Jerry&gt;")

    def test_escapeAttribute(self):
        self.assertEqual(ooolibXML.escapeAttribute('R&D "Q3" \'x\''),
                         "R&amp;D &quot;Q3&quot; &apos;x&apos;")

    def test_attributes_are_escaped(self):
        element = ooolibXML.Element("draw:page")
        element.setAttribute("draw:name", 'A & "B"')
        self.assertEqual(element.toString(), '<draw:page draw:name="A &amp; &quot;B&quot;"/>')

    def test_paragraph_spaces(self):
        self.assertEqual(ooolibXML.escapeParagraphText("a b"), "a b")
        self.assertEqual(ooolibXML.escapeParagraphText("  two  spaces"),
                         '<text:s text:c="2"/>two <text:s/>spaces')
        self.assertEqual(ooolibXML.escapeParagraphText("end "), "end<text:s/>")

    def test_paragraph_tabs_and_breaks(self):
        self.assertEqual(ooolibXML.escapeParagraphText("a\tb\nc"),
                         "a<text:tab/>b<text:line-break/>c")
        self.assertEqual(ooolibXML.escapeParagraphText("a\r\nb"), "a<text:line-break/>b")

    def test_clearChildren(self):
        element = ooolibXML.Element("parent")
        element.addChild(ooolibXML.Element("child"))
        element.clearChildren()
        self.assertEqual(element.toString(), "<parent/>")


class UnitsTest(unittest.TestCase):
    def test_convertLength(self):
        self.assertAlmostEqual(ooolibUnits.convertLength("1in", "cm"), 2.54)
        self.assertAlmostEqual(ooolibUnits.convertLength("25.4mm", "in"), 1.0)
        self.assertAlmostEqual(ooolibUnits.convertLength("72pt", "in"), 1.0)
        self.assertAlmostEqual(ooolibUnits.convertLength("3", "mm"), 76.2)

    def test_convertLength_invalid(self):
        self.assertIsNone(ooolibUnits.convertLength("abc"))
        self.assertIsNone(ooolibUnits.convertLength("50%"))
        self.assertIsNone(ooolibUnits.convertLength(None))

    def test_formatLength(self):
        self.assertEqual(ooolibUnits.formatLength(2.5, "in"), "2.5in")
        self.assertEqual(ooolibUnits.formatLength(3, "mm"), "3mm")
        self.assertEqual(ooolibUnits.formatLength(1 / 3.0, "in"), "0.3333in")
        self.assertEqual(ooolibUnits.formatLength("30mm", "in"), "30mm")
        self.assertIsNone(ooolibUnits.formatLength(None, "in"))
        self.assertRaises(TypeError, ooolibUnits.formatLength, True, "in")


class FileTest(helpers.DocumentTestCase):
    def test_save_and_read(self):
        filename = self.path("test.zip")
        saveFile = ooolibFile.SaveFile(filename)
        saveFile.insertFileString("mimetype", "text/plain")
        saveFile.insertFileString("hello.txt", "Hello é")
        saveFile.close()
        readFile = ooolibFile.ReadFile(filename)
        self.assertEqual(readFile.getFileList(), ["mimetype", "hello.txt"])
        self.assertEqual(readFile.readFileString("hello.txt"), "Hello é")
        self.assertIsNone(readFile.readFileString("missing.txt"))
        self.assertTrue(readFile.hasFile("mimetype"))
        readFile.close()


if __name__ == "__main__":
    unittest.main()
