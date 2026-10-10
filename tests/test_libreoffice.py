#################################################################################
# ooolib-python - Python module for creating Open Document Format documents.    #
# Copyright (C) 2006-2026  Joseph Colton                                        #
#                                                                               #
# You can contact me by email at josephcolton@gmail.com                         #
#################################################################################
"""Checks documents with LibreOffice itself. Skipped when soffice is not installed,
or when the OOOLIB_SKIP_LIBREOFFICE environment variable is set (these tests are slow)."""

# Import python modules
import os
import re
import shutil
import subprocess
import unittest
import zipfile

# Import test helpers (also puts code/ on the path)
import helpers
import ooolib

SOFFICE = shutil.which("soffice") or shutil.which("libreoffice")

@unittest.skipIf(SOFFICE is None, "LibreOffice (soffice) is not installed")
@unittest.skipIf(os.environ.get("OOOLIB_SKIP_LIBREOFFICE"), "OOOLIB_SKIP_LIBREOFFICE is set")
class LibreOfficeTest(helpers.DocumentTestCase):
    def convert(self, filename, outputFormat):
        """Convert filename with LibreOffice and return the new file's path."""
        outputDir = self.path("converted")
        profile = "file://" + self.path("profile")
        subprocess.run([SOFFICE, "-env:UserInstallation=" + profile, "--headless",
                        "--convert-to", outputFormat, "--outdir", outputDir, filename],
                       capture_output=True, timeout=180)
        base = os.path.splitext(os.path.basename(filename))[0]
        converted = os.path.join(outputDir, base + "." + outputFormat)
        self.assertTrue(os.path.exists(converted), "LibreOffice could not convert %s" % filename)
        return converted

    def breakConnectorEnds(self, filename):
        """Set every connector end point to 0 so a passing test proves LibreOffice
        routed the connectors from the shapes they are attached to."""
        with zipfile.ZipFile(filename) as archive:
            files = [(info.filename, archive.read(info.filename)) for info in archive.infolist()]
        with zipfile.ZipFile(filename, "w") as archive:
            for name, data in files:
                if name == "content.xml":
                    data = re.sub(rb'(svg:[xy][12])="[^"]*"', rb'\1="0in"', data)
                compression = zipfile.ZIP_STORED if name == "mimetype" else zipfile.ZIP_DEFLATED
                archive.writestr(name, data, compression)

    def test_libreoffice_opens_every_type(self):
        documents = []
        calc = ooolib.Calc()
        calc.content.activeTable.setCellText(1, 1, "Hello & <goodbye>")
        documents.append((calc, "calc.ods"))
        writer = ooolib.Writer()
        writer.content.addParagraph("  Spaces\tand tabs")
        documents.append((writer, "writer.odt"))
        impress = ooolib.Impress()
        impress.content.activeSlide.setTitle("Slide")
        documents.append((impress, "impress.odp"))
        draw = ooolib.Draw()
        draw.content.activePage.addRectangle(1, 1, 2, 1, text="Box")
        documents.append((draw, "draw.odg"))
        for doc, filename in documents:
            with self.subTest(filename):
                pdf = self.convert(self.export(doc, filename), "pdf")
                self.assertGreater(os.path.getsize(pdf), 0)

    def test_draw_survives_libreoffice_save(self):
        doc = ooolib.Draw()
        page = doc.content.activePage
        doc.styles.setPageSize("8.5in", "11in")
        page.addRectangle(2, 3, 1.5, 0.5, name="Rotated", text="Turn", rotation=30)
        page.addPolygon([(1, 1), (3, 1), (2, 2.5)], name="Triangle")
        page.addLine(1, 5, 4, 5, name="Arrow", linestyle="dash", endarrow="arrow")
        group = page.addGroup(name="Group")
        group.addEllipse(5, 5, 1, 1, name="Inside", text="  two  spaces")
        # Deliberately start the connector on the left of "Triangle", which is far from "Inside"
        page.addConnector("Triangle", "Inside", startside="left", endside="top", endarrow="arrow", name="Link")
        filename = self.export(doc, "draw.odg")
        self.breakConnectorEnds(filename)
        resaved = ooolib.openDocument(self.convert(filename, "odg"))
        start, end, link = resaved.pages[0].getConnections()[0]
        self.assertEqual((start.name, end.name, link.startSide, link.endSide), ("Triangle", "Inside", "left", "top"))
        # LibreOffice routes connectors from the shapes; the ends must sit on the chosen sides
        self.assertAlmostEqual(link.getLength("x1"), 1, places=2)
        self.assertAlmostEqual(link.getLength("y1"), 1.75, places=2)
        self.assertAlmostEqual(link.getLength("x2"), 5.5, places=2)
        self.assertAlmostEqual(link.getLength("y2"), 5, places=2)
        self.assertEqual(resaved.pageLayout["orientation"], "portrait")
        rotated = resaved.pages[0].getShape("Rotated")
        self.assertAlmostEqual(rotated.rotation, 30, places=3)
        for actual, wanted in zip(rotated.getBounds("in"), (2, 3, 1.5, 0.5)):
            self.assertAlmostEqual(actual, wanted, places=2)
        self.assertEqual(rotated.text, "Turn")
        triangle = resaved.pages[0].getShape("Triangle")
        self.assertAlmostEqual(triangle.getPoints("in")[2][1], 2.5, places=2)
        arrow = resaved.pages[0].getShape("Arrow")
        self.assertEqual((arrow.strokeDash, arrow.markerEnd), ("Dash", "Arrow"))
        self.assertEqual(resaved.pages[0].getShape("Inside").text, "  two  spaces")
        self.assertEqual(resaved.pages[0].getShape("Group").type, "group")


if __name__ == "__main__":
    unittest.main()
