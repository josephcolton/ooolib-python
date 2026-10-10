#################################################################################
# ooolib-python - Python module for creating Open Document Format documents.    #
# Copyright (C) 2006-2026  Joseph Colton                                        #
#                                                                               #
# You can contact me by email at josephcolton@gmail.com                         #
#################################################################################
"""Tests for Impress presentations, including the Draw shape features on slides."""

# Import python modules
import re
import unittest

# Import test helpers (also puts code/ on the path)
import helpers
import ooolib

class ImpressTestCase(helpers.DocumentTestCase):
    def setUp(self):
        helpers.DocumentTestCase.setUp(self)
        self.doc = ooolib.Impress()
        self.slide = self.doc.content.activeSlide

    def readBack(self):
        return ooolib.openDocument(self.export(self.doc, "slides.odp"))

    def assertBounds(self, shape, expected, places=3):
        bounds = shape.getBounds("in")
        self.assertIsNotNone(bounds, shape)
        for actual, wanted in zip(bounds, expected):
            self.assertAlmostEqual(actual, wanted, places=places)


class SlideTest(ImpressTestCase):
    def test_title_and_content(self):
        self.slide.setTitle("Title & More")
        self.slide.addContent("Point one", bold=True, fontsize=20)
        self.slide.addContent("  Point two")
        slide = self.readBack().pages[0]
        title = [shape for shape in slide.shapes if shape.presentationClass == "title"][0]
        self.assertEqual(title.text, "Title & More")
        self.assertBounds(title, (0.6, 0.3, 8.8, 1.2))
        content = [shape for shape in slide.shapes if shape.presentationClass == "subtitle"][0]
        self.assertEqual(content.paragraphs, ["Point one", "  Point two"])
        self.assertBounds(content, (0.6, 1.8, 8.8, 5.3))

    def test_widescreen_slides(self):
        self.doc.styles.setPageSize("13.333in", "7.5in")
        self.slide.setTitle("Wide")
        info = self.readBack()
        layout = info.pageLayout
        self.assertEqual((layout["width"], layout["height"], layout["orientation"]), ("13.333in", "7.5in", "landscape"))
        title = info.pages[0].shapes[0]
        self.assertBounds(title, (0.8, 0.3, 11.733, 1.2))
        self.assertEqual(self.doc.styles.getPageSize(), ("13.333in", "7.5in"))

    def test_master_placeholders(self):
        # Master page frames must be placeholders, or LibreOffice draws them on every slide
        styles = self.readZipString(self.export(self.doc, "master.odp"), "styles.xml")
        master = re.search(r"<style:master-page.*?</style:master-page>", styles).group(0)
        frames = re.findall(r"<draw:frame[^>]*>", master)
        self.assertEqual(len(frames), 2)
        for frame in frames:
            self.assertIn('presentation:placeholder="true"', frame)
            self.assertIn("presentation:style-name=", frame)

    def test_unique_style_names(self):
        self.slide.addContent("Bold", bold=True)
        self.slide.addRectangle(1, 1, 1, 1, text="Italic", italic=True)
        self.slide.setBackgroundColor("#ff0000")
        content = self.readZipString(self.export(self.doc, "styles.odp"), "content.xml")
        names = re.findall(r'<style:style style:name="([^"]+)"', content)
        self.assertEqual(len(names), len(set(names)), names)
        self.assertNotIn("dp1", names) # dp1 is the default background style in styles.xml


class SlideShapesTest(ImpressTestCase):
    def test_shapes_on_slides(self):
        self.doc.content.setUnit("cm")
        self.slide.addRectangle(2.54, 2.54, 5.08, 2.54, fillcolor="#dae3f3", name="Box", text="Hello", rotation=15)
        self.slide.addPolygon([(10, 10), (15, 10), (12.5, 15)], name="Triangle")
        self.slide.addLine(1, 1, 5, 1, name="Arrow", linestyle="dash", endarrow="arrow")
        slide = self.readBack().pages[0]
        box = slide.getShape("Box")
        self.assertEqual((box.type, box.text, box.fillColor), ("rect", "Hello", "#dae3f3"))
        self.assertAlmostEqual(box.rotation, 15, places=4)
        self.assertBounds(box, (1, 1, 2, 1))
        self.assertEqual(slide.getShape("Triangle").type, "polygon")
        arrow = slide.getShape("Arrow")
        self.assertEqual((arrow.strokeDash, arrow.markerEnd), ("Dash", "Arrow"))

    def test_groups_and_connectors(self):
        group = self.slide.addGroup(name="Pair")
        group.addRectangle(1, 3, 2, 1, name="A")
        group.addRectangle(5, 3, 2, 1, name="B")
        self.slide.addConnector("A", "B", startside="right", endside="left", endarrow="arrow", text="link")
        slide = self.readBack().pages[0]
        self.assertEqual([child.name for child in slide.getShape("Pair").children], ["A", "B"])
        connections = slide.getConnections()
        self.assertEqual([(start.name, end.name, connector.text) for (start, end, connector) in connections],
                         [("A", "B", "link")])

    def test_images(self):
        self.slide.addImage(helpers.makePNG(192, 96), 1, 1, name="Picture")
        self.doc.content.addSlide("Second")
        self.doc.content.activeSlide.addImage(helpers.makePNG(192, 96), 2, 2, width=1, name="Again")
        self.doc.content.activeSlide.addImage(helpers.SVG, 4, 2, name="Vector")
        filename = self.export(self.doc, "images.odp")
        self.export(self.doc, "again.odp")
        info = ooolib.openDocument(filename)
        self.assertEqual(sorted(image["mediaType"] for image in info.images), ["image/png", "image/svg+xml"])
        self.assertBounds(info.pages[0].getShape("Picture"), (1, 1, 2, 1))
        self.assertBounds(info.getPage("Second").getShape("Again"), (2, 2, 1, 0.5))
        self.assertEqual(info.readImage(info.getPage("Second").getShape("Vector")), helpers.SVG)
        manifest = self.readZipString(self.path("again.odp"), "META-INF/manifest.xml")
        self.assertEqual(manifest.count("Pictures/"), 2)

    def test_slide_text_includes_shapes(self):
        self.slide.setTitle("Title")
        self.slide.addRectangle(1, 4, 2, 1, text="In a shape")
        self.assertEqual(self.readBack().getText(), "Title\nIn a shape")


if __name__ == "__main__":
    unittest.main()
