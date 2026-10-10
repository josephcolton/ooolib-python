#################################################################################
# ooolib-python - Python module for creating Open Document Format documents.    #
# Copyright (C) 2006-2026  Joseph Colton                                        #
#                                                                               #
# You can contact me by email at josephcolton@gmail.com                         #
#################################################################################
"""Tests for Draw documents. Most tests create a drawing, export it, and
read it back with ooolib.openDocument() to check the result."""

# Import python modules
import unittest

# Import test helpers (also puts code/ on the path)
import helpers
import ooolib

class DrawTestCase(helpers.DocumentTestCase):
    def setUp(self):
        helpers.DocumentTestCase.setUp(self)
        self.doc = ooolib.Draw()
        self.page = self.doc.content.activePage

    def readBack(self):
        return ooolib.openDocument(self.export(self.doc, "drawing.odg"))

    def assertBounds(self, shape, expected, places=3):
        bounds = shape.getBounds("in")
        self.assertIsNotNone(bounds, shape)
        for actual, wanted in zip(bounds, expected):
            self.assertAlmostEqual(actual, wanted, places=places)


class PageSetupTest(DrawTestCase):
    def test_default_page(self):
        layout = self.readBack().pageLayout
        self.assertEqual((layout["width"], layout["height"], layout["orientation"]), ("10in", "7.5in", "landscape"))
        self.assertEqual(layout["marginTop"], "0.5in")

    def test_portrait_page(self):
        self.doc.styles.setPageSize("8.5in", "11in")
        self.doc.styles.setPageMargins("1in", left="0.75in")
        layout = self.readBack().pageLayout
        self.assertEqual((layout["width"], layout["height"], layout["orientation"]), ("8.5in", "11in", "portrait"))
        self.assertEqual((layout["marginTop"], layout["marginLeft"], layout["marginRight"]), ("1in", "0.75in", "1in"))

    def test_page_size_uses_document_unit(self):
        self.doc.content.setUnit("mm")
        self.doc.styles.setPageSize(297, 210)
        self.assertEqual(self.doc.styles.getPageSize(), ("297mm", "210mm"))

    def test_bad_orientation(self):
        self.assertRaises(ValueError, self.doc.styles.setPageSize, 5, 5, "sideways")


class UnitsAndNamesTest(DrawTestCase):
    def test_numeric_coordinates(self):
        self.page.addRectangle(1, 1.5, 3, 2, name="Inches")
        self.doc.content.setUnit("cm")
        self.page.addRectangle(2.54, 5.08, 2.54, 2.54, name="Centimeters")
        self.page.addRectangle("1in", "1in", "25.4mm", "72pt", name="Strings")
        page = self.readBack().pages[0]
        self.assertBounds(page.getShape("Inches"), (1, 1.5, 3, 2))
        self.assertBounds(page.getShape("Centimeters"), (1, 2, 1, 1))
        self.assertBounds(page.getShape("Strings"), (1, 1, 1, 1))

    def test_bad_values(self):
        self.assertRaises(ValueError, self.doc.content.setUnit, "furlong")
        self.assertRaises(TypeError, self.page.addRectangle, True, 1, 1, 1)

    def test_names_and_text(self):
        self.page.addRectangle(1, 1, 3, 1, fillcolor="#dae3f3", name="Box & <1>", text="Hello\nWorld")
        self.page.addEllipse(1, 3, 2, 1, name="Oval", text="  spaced  ")
        self.page.addTextBox(1, 5, 4, 1, "Title", name="Title", bold=True, fontsize=20, align="center")
        self.page.addLine(0, 0, 1, 1, name="Line")
        page = self.readBack().pages[0]
        self.assertEqual(page.getShape("Box & <1>").text, "Hello\nWorld")
        self.assertEqual(page.getShape("Box & <1>").fillColor, "#dae3f3")
        self.assertEqual(page.getShape("Oval").text, "  spaced  ")
        self.assertEqual(page.getShape("Title").type, "textbox")
        self.assertEqual(page.getShape("Line").type, "line")

    def test_shape_handles(self):
        box = self.page.addRectangle(1, 1, 2, 1, name="Step")
        self.assertIs(self.page.getShape("Step"), box)
        box.setText("Changed", italic=True).setName("Renamed")
        self.assertIsNone(self.page.getShape("Step"))
        self.assertEqual(self.readBack().pages[0].getShape("Renamed").text, "Changed")

    def test_text_errors(self):
        self.assertRaises(ValueError, self.page.addRectangle, 1, 1, 1, 1, text="x", align="middle")
        line = self.page.addLine(0, 0, 1, 1)
        self.assertRaises(ValueError, line.setText, "x")

    def test_text_alignment_style(self):
        self.page.addRectangle(1, 1, 2, 1, text="Right", align="right")
        content = self.readZipString(self.export(self.doc, "align.odg"), "content.xml")
        self.assertIn('fo:text-align="end"', content)
        self.assertIn('draw:textarea-vertical-align="middle"', content)


class PolygonTest(DrawTestCase):
    def test_polygon_points(self):
        points = [(1.25, 0.5), (2.25, 2.25), (0.25, 2.25)]
        self.page.addPolygon(points, fillcolor="#ffd966", name="Triangle", text="Tri")
        shape = self.readBack().pages[0].getShape("Triangle")
        self.assertEqual(shape.type, "polygon")
        self.assertEqual(shape.text, "Tri")
        self.assertBounds(shape, (0.25, 0.5, 2.0, 1.75))
        for actual, wanted in zip(shape.getPoints("in"), points):
            self.assertAlmostEqual(actual[0], wanted[0], places=3)
            self.assertAlmostEqual(actual[1], wanted[1], places=3)

    def test_polyline_points_in_mm(self):
        self.doc.content.setUnit("mm")
        self.page.addPolyline([(10, 10), (20, 30), (30, 10)], name="Zigzag")
        shape = self.readBack().pages[0].getShape("Zigzag")
        self.assertEqual(shape.type, "polyline")
        points = shape.getPoints("mm")
        self.assertEqual(len(points), 3)
        self.assertAlmostEqual(points[1][0], 20, places=1)
        self.assertAlmostEqual(points[1][1], 30, places=1)

    def test_straight_polyline(self):
        # A polyline with no height still gets a valid (non-zero) view box
        self.page.addPolyline([(1, 2), (3, 2), (5, 2)], name="Flat")
        shape = self.readBack().pages[0].getShape("Flat")
        self.assertNotIn(" 0", shape.viewBox.split(" ", 2)[-1])

    def test_too_few_points(self):
        self.assertRaises(ValueError, self.page.addPolygon, [(0, 0), (1, 1)])
        self.assertRaises(ValueError, self.page.addPolyline, [(0, 0)])
        self.assertRaises(ValueError, self.page.addPolygon, [(0, 0), (1, 1), ("50%", 1)])


class LineStyleTest(DrawTestCase):
    def test_dashes(self):
        expected = {"solid": None, "dash": "Dash", "dot": "Dot", "dashdot": "DashDot", "longdash": "LongDash"}
        for linestyle in expected:
            self.page.addLine(0, 0, 1, 0, name=linestyle, linestyle=linestyle)
        self.page.addRectangle(1, 1, 1, 1, strokecolor="#000000", linestyle="dot", name="Dotted Box")
        page = self.readBack().pages[0]
        for linestyle in expected:
            self.assertEqual(page.getShape(linestyle).strokeDash, expected[linestyle], linestyle)
        self.assertEqual(page.getShape("Dotted Box").strokeDash, "Dot")

    def test_line_ends(self):
        self.page.addLine(0, 0, 1, 0, name="Both", startarrow="circle", endarrow="arrow")
        self.page.addPolyline([(0, 0), (1, 1), (2, 0)], name="End", endarrow="diamond", arrowsize="0.3in")
        self.page.addLine(0, 1, 1, 1, name="None")
        page = self.readBack().pages[0]
        self.assertEqual((page.getShape("Both").markerStart, page.getShape("Both").markerEnd), ("Circle", "Arrow"))
        self.assertEqual(page.getShape("End").markerEnd, "Diamond")
        self.assertEqual(page.getShape("End").graphicProperties["draw:marker-end-width"], "0.3in")
        self.assertIsNone(page.getShape("None").markerEnd)

    def test_bad_line_values(self):
        self.assertRaises(ValueError, self.page.addLine, 0, 0, 1, 1, linestyle="wavy")
        self.assertRaises(ValueError, self.page.addLine, 0, 0, 1, 1, endarrow="star")


class RotationTest(DrawTestCase):
    def test_rotation_round_trip(self):
        for angle in (30, 90, 200, -45):
            self.page.addRectangle(2, 3, 1.5, 0.5, name="R%d" % angle, rotation=angle)
        self.page.addTextBox(1, 1, 2, 1, "Turned", name="Text", rotation=45)
        self.page.addPolygon([(5, 5), (6, 5), (5.5, 6)], name="Poly", rotation=10)
        page = self.readBack().pages[0]
        for angle in (30, 90, 200, -45):
            shape = page.getShape("R%d" % angle)
            self.assertAlmostEqual(shape.rotation, angle % 360, places=4)
            self.assertBounds(shape, (2, 3, 1.5, 0.5))
        self.assertAlmostEqual(page.getShape("Text").rotation, 45, places=4)
        self.assertBounds(page.getShape("Poly"), (5, 5, 1, 1))

    def test_no_rotation(self):
        self.page.addRectangle(1, 1, 1, 1, name="Plain", rotation=360)
        shape = self.readBack().pages[0].getShape("Plain")
        self.assertIsNone(shape.rotation)
        self.assertIsNone(shape.transform)

    def test_rotation_errors(self):
        line = self.page.addLine(0, 0, 1, 1)
        self.assertRaises(ValueError, line.setRotation, 45)
        self.assertRaises(TypeError, self.page.addRectangle, 1, 1, 1, 1, rotation="45")


class GroupTest(DrawTestCase):
    def test_groups(self):
        group = self.page.addGroup(name="Server")
        group.addRectangle(7, 5.5, 2, 1.2, name="Server Box", text="Server")
        inner = group.addGroup(name="Lights")
        inner.addEllipse(7.1, 5.6, 0.2, 0.2, name="Light")
        self.page.addRectangle(0, 0, 1, 1, name="Outside")
        # Generator side: groups are searched by getShape
        self.assertIs(self.page.getShape("Light").getType(), "ellipse")
        self.assertEqual(len(self.page.getShapes()), 2)
        # Reader side
        page = self.readBack().pages[0]
        self.assertEqual([shape.type for shape in page.shapes], ["group", "rect"])
        server = page.getShape("Server")
        self.assertEqual([child.name for child in server.children], ["Server Box", "Lights"])
        self.assertEqual(page.getShape("Light").type, "ellipse")
        self.assertEqual(server.getText(), "Server")
        self.assertEqual(len(page.getShapes()), 5)


class ConnectorTest(DrawTestCase):
    def addBoxes(self):
        self.page.addRectangle(1, 1, 2, 1, name="A", text="A")
        self.page.addRectangle(6, 4, 2, 1, name="B", text="B")
        group = self.page.addGroup(name="Group")
        group.addEllipse(1, 5, 2, 1, name="C")

    def test_connections_round_trip(self):
        self.addBoxes()
        self.page.addConnector("A", "B", startside="right", endside="top", endarrow="arrow",
                               name="A to B", text="calls")
        self.page.addConnector(self.page.getShape("A"), "C", connectortype="curve", linestyle="dash")
        connections = self.readBack().pages[0].getConnections()
        self.assertEqual([(start.name, end.name) for (start, end, connector) in connections],
                         [("A", "B"), ("A", "C")])
        first = connections[0][2]
        self.assertEqual((first.name, first.text, first.connectorType), ("A to B", "calls", "standard"))
        self.assertEqual((first.startSide, first.endSide, first.markerEnd), ("right", "top", "Arrow"))
        second = connections[1][2]
        self.assertEqual((second.connectorType, second.startSide, second.endSide, second.strokeDash),
                         ("curve", None, None, "Dash"))

    def test_fallback_end_points(self):
        self.addBoxes()
        self.page.addConnector("A", "B", startside="bottom", endside="left", name="Explicit")
        self.page.addConnector("A", "B", name="Automatic")
        page = self.readBack().pages[0]
        explicit = page.getShape("Explicit")
        self.assertEqual([explicit.getLength(name) for name in ("x1", "y1", "x2", "y2")], [2, 2, 6, 4.5])
        # Automatic sides face each other: B is further right than down, so right -> left
        automatic = page.getShape("Automatic")
        self.assertEqual([automatic.getLength(name) for name in ("x1", "y1", "x2", "y2")], [3, 1.5, 6, 4.5])

    def test_rotated_shape_glue_point(self):
        self.page.addRectangle(1, 1, 2, 1, name="Turned", rotation=90)
        self.page.addRectangle(5, 1, 1, 1, name="Plain")
        self.page.addConnector("Turned", "Plain", startside="right", endside="left", name="Link")
        link = self.readBack().pages[0].getShape("Link")
        # After a 90 degree turn the right side faces up: center (2, 1.5) moved up by half the width
        self.assertAlmostEqual(link.getLength("x1"), 2, places=3)
        self.assertAlmostEqual(link.getLength("y1"), 0.5, places=3)

    def test_shared_ids(self):
        self.addBoxes()
        self.page.addConnector("A", "B")
        self.page.addConnector("A", "C")
        content = self.readZipString(self.export(self.doc, "ids.odg"), "content.xml")
        self.assertEqual(content.count('draw:id="id1"'), 1) # A keeps one id for both connectors
        self.assertIn('xml:id="id1"', content)
        info = ooolib.openDocument(self.path("ids.odg"))
        self.assertEqual(len([shape for shape in info.getShapes() if shape.id]), 3)

    def test_connector_errors(self):
        self.addBoxes()
        line = self.page.addLine(0, 0, 1, 1, name="Line")
        self.assertRaises(ValueError, self.page.addConnector, "A", "Missing")
        self.assertRaises(ValueError, self.page.addConnector, "A", line)
        self.assertRaises(ValueError, self.page.addConnector, "A", "Group")
        self.assertRaises(ValueError, self.page.addConnector, "A", "B", startside="middle")
        self.assertRaises(ValueError, self.page.addConnector, "A", "B", connectortype="zigzag")
        connector = self.page.addConnector("A", "B")
        self.assertRaises(ValueError, connector.setRotation, 45)


class ImageTest(DrawTestCase):
    def test_image_sizes(self):
        png = helpers.makePNG(192, 96)                      # 2in x 1in at 96 dpi
        self.page.addImage(png, 1, 1, name="Natural")
        self.page.addImage(png, 1, 3, width=4, name="Width")
        self.page.addImage(png, 1, 5, height=0.5, name="Height")
        self.page.addImage(png, 6, 1, width=1, height=3, name="Stretched")
        self.page.addImage(helpers.makePNG(300, 300, dpi=300), 6, 5, name="High Resolution")
        page = self.readBack().pages[0]
        self.assertBounds(page.getShape("Natural"), (1, 1, 2, 1))
        self.assertBounds(page.getShape("Width"), (1, 3, 4, 2))
        self.assertBounds(page.getShape("Height"), (1, 5, 1, 0.5))
        self.assertBounds(page.getShape("Stretched"), (6, 1, 1, 3))
        self.assertBounds(page.getShape("High Resolution"), (6, 5, 1, 1))

    def test_images_stored_once(self):
        filename = self.path("picture.png")
        with open(filename, "wb") as pngFile:
            pngFile.write(helpers.makePNG(10, 10))
        self.page.addImage(filename, 1, 1, name="From File")
        self.page.addImage(helpers.makePNG(10, 10), 3, 1, name="From Bytes")   # same image
        self.page.addImage(helpers.SVG, 5, 1, name="Vector")
        info = self.readBack()
        self.assertEqual(sorted(image["mediaType"] for image in info.images), ["image/png", "image/svg+xml"])
        page = info.pages[0]
        self.assertEqual(page.getShape("From File").href, page.getShape("From Bytes").href)
        self.assertEqual(info.readImage(page.getShape("From File")), helpers.makePNG(10, 10))
        self.assertEqual(info.readImage(page.getShape("Vector").href), helpers.SVG)
        self.assertBounds(page.getShape("Vector"), (5, 1, 2, 1))
        png = [image for image in info.images if image["mediaType"] == "image/png"][0]
        self.assertEqual((png["pixelWidth"], png["pixelHeight"]), (10, 10))

    def test_manifest_and_storage(self):
        self.page.addImage(helpers.makePNG(10, 10), 1, 1)
        self.page.addImage(helpers.SVG, 3, 1)
        filename = self.export(self.doc, "images.odg")
        self.export(self.doc, "again.odg") # Exporting twice must not repeat manifest entries
        manifest = self.readZipString(self.path("again.odg"), "META-INF/manifest.xml")
        self.assertEqual(manifest.count('manifest:media-type="image/png"'), 1)
        self.assertEqual(manifest.count('manifest:media-type="image/svg+xml"'), 1)
        with helpers.zipfile.ZipFile(filename) as archive:
            pictures = {info.filename: info.compress_type for info in archive.infolist()
                        if info.filename.startswith("Pictures/")}
        self.assertEqual(sorted(pictures.values()), [helpers.zipfile.ZIP_STORED, helpers.zipfile.ZIP_DEFLATED])

    def test_image_options(self):
        self.page.addImage(helpers.makePNG(96, 96), 1, 1, name="Framed", strokecolor="#ff0000",
                           linestyle="dash", rotation=30)
        self.page.addRectangle(4, 1, 1, 1, name="Box")
        self.page.addConnector("Framed", "Box", endarrow="arrow")
        page = self.readBack().pages[0]
        framed = page.getShape("Framed")
        self.assertEqual((framed.type, framed.strokeColor, framed.strokeDash), ("image", "#ff0000", "Dash"))
        self.assertAlmostEqual(framed.rotation, 30, places=4)
        self.assertBounds(framed, (1, 1, 1, 1))
        self.assertEqual([(start.name, end.name) for (start, end, connector) in page.getConnections()],
                         [("Framed", "Box")])

    def test_image_errors(self):
        self.assertRaises(ValueError, self.page.addImage, b"not an image", 1, 1)
        self.assertRaises(FileNotFoundError, self.page.addImage, self.path("missing.png"), 1, 1)


class ReaderTest(DrawTestCase):
    def test_pages_and_text(self):
        self.page.addTextBox(1, 1, 3, 1, "First page")
        self.doc.content.addPage("R&D \"Q3\"")
        self.doc.content.activePage.addRectangle(1, 1, 2, 1, text="Second page")
        info = self.readBack()
        self.assertEqual(info.getPageNames(), ["page1", 'R&D "Q3"'])
        self.assertEqual(info.getPage('R&D "Q3"').getText(), "Second page")
        self.assertEqual(info.getPage(1).getText(), "First page")
        self.assertEqual(info.getText(), "First page\nSecond page")
        self.assertEqual(info.meta["statistics"]["page-count"], 2)

    def test_fill_and_stroke(self):
        self.page.addRectangle(1, 1, 1, 1, fillcolor="#ff0000", strokecolor="#0000ff", strokewidth=0.05, name="Both")
        self.page.addRectangle(1, 1, 1, 1, name="Neither")
        page = self.readBack().pages[0]
        both = page.getShape("Both")
        self.assertEqual((both.fill, both.fillColor, both.strokeColor, both.strokeWidth),
                         ("solid", "#ff0000", "#0000ff", "0.05in"))
        neither = page.getShape("Neither")
        self.assertEqual((neither.fill, neither.fillColor, neither.stroke, neither.strokeColor),
                         ("none", None, "none", None))

    def test_toDict(self):
        import json
        self.page.addPolygon([(0, 0), (1, 0), (0, 1)], name="Tri", rotation=15)
        data = json.loads(json.dumps(self.readBack().toDict()))
        self.assertEqual(data["type"], "graphics")
        self.assertEqual(data["pages"][0]["shapes"][0]["name"], "Tri")


if __name__ == "__main__":
    unittest.main()
