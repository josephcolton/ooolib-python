#################################################################################
# ooolib-python - Python module for creating Open Document Format documents.    #
# Copyright (C) 2006-2026  Joseph Colton                                        #
#                                                                               #
# You can contact me by email at josephcolton@gmail.com                         #
#################################################################################
"""Tests for the XML, unit, and file helpers."""

# Import python modules
import struct
import unittest

# Import test helpers (also puts code/ on the path)
import helpers
import ooolibFile
import ooolibImage
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


class ImageInfoTest(unittest.TestCase):
    def test_png(self):
        info = ooolibImage.readImageInfo(helpers.makePNG(30, 20))
        self.assertEqual((info["format"], info["mediaType"], info["extension"]), ("png", "image/png", "png"))
        self.assertEqual((info["pixelWidth"], info["pixelHeight"], info["dpiX"]), (30, 20, 96))

    def test_png_resolution(self):
        info = ooolibImage.readImageInfo(helpers.makePNG(300, 150, dpi=300))
        self.assertAlmostEqual(info["dpiX"], 300, places=1)
        width, height = ooolibImage.getImageSize(info, "in")
        self.assertAlmostEqual(width, 1.0, places=3)
        self.assertAlmostEqual(height, 0.5, places=3)

    def test_jpeg(self):
        # Headers only: start of image, JFIF at 72 dpi, and a start of frame for 640 x 480
        jfif = b"JFIF\x00\x01\x01\x01" + struct.pack(">HH", 72, 72) + b"\x00\x00"
        frame = b"\x08" + struct.pack(">HH", 480, 640) + b"\x03" + b"\x00" * 9
        data = (b"\xff\xd8" + b"\xff\xe0" + struct.pack(">H", len(jfif) + 2) + jfif +
                b"\xff\xc0" + struct.pack(">H", len(frame) + 2) + frame)
        info = ooolibImage.readImageInfo(data)
        self.assertEqual((info["format"], info["pixelWidth"], info["pixelHeight"], info["dpiX"]), ("jpeg", 640, 480, 72))
        self.assertEqual(info["extension"], "jpg")

    def test_gif_and_bmp(self):
        gif = ooolibImage.readImageInfo(b"GIF89a" + struct.pack("<HH", 40, 30) + b"\x00" * 10)
        self.assertEqual((gif["format"], gif["pixelWidth"], gif["pixelHeight"]), ("gif", 40, 30))
        bmpHeader = b"BM" + b"\x00" * 16 + struct.pack("<ii", 50, -25) + b"\x00" * 12 + struct.pack("<ii", 3780, 3780)
        bmp = ooolibImage.readImageInfo(bmpHeader)
        self.assertEqual((bmp["format"], bmp["pixelWidth"], bmp["pixelHeight"]), ("bmp", 50, 25))
        self.assertAlmostEqual(bmp["dpiX"], 96, places=0)

    def test_svg(self):
        info = ooolibImage.readImageInfo(helpers.SVG)
        self.assertEqual(ooolibImage.getImageSize(info, "in"), (2.0, 1.0))
        viewBoxOnly = b'<?xml version="1.0"?>\n<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 192 96"/>'
        self.assertEqual(ooolibImage.getImageSize(ooolibImage.readImageInfo(viewBoxOnly), "in"), (2.0, 1.0))
        widthOnly = b'<svg xmlns="http://www.w3.org/2000/svg" width="40mm" viewBox="0 0 200 100"/>'
        width, height = ooolibImage.getImageSize(ooolibImage.readImageInfo(widthOnly), "mm")
        self.assertAlmostEqual(width, 40)
        self.assertAlmostEqual(height, 20)

    def test_unsupported(self):
        self.assertRaises(ValueError, ooolibImage.readImageInfo, b"not an image")


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
