#################################################################################
# ooolib-python - Python module for creating Open Document Format documents.    #
# Copyright (C) 2006-2026  Joseph Colton                                        #
#                                                                               #
# You can contact me by email at josephcolton@gmail.com                         #
#################################################################################
"""Tests that apply to every document type, plus Calc, Writer, and Impress."""

# Import python modules
import datetime
import re
import unittest

# Import test helpers (also puts code/ on the path)
import helpers
import ooolib

DOCUMENT_TYPES = [
    (ooolib.Calc, ".ods", "spreadsheet"),
    (ooolib.Writer, ".odt", "text"),
    (ooolib.Impress, ".odp", "presentation"),
    (ooolib.Draw, ".odg", "graphics"),
]

class AllDocumentsTest(helpers.DocumentTestCase):
    def test_blank_documents_are_valid(self):
        for documentClass, extension, documentType in DOCUMENT_TYPES:
            with self.subTest(documentType):
                filename = self.export(documentClass(), "blank" + extension)
                self.assertEqual(ooolib.openDocument(filename).type, documentType)

    def test_export_twice_gives_same_content(self):
        for documentClass, extension, documentType in DOCUMENT_TYPES:
            with self.subTest(documentType):
                doc = documentClass()
                first = self.export(doc, "first" + extension)
                second = self.export(doc, "second" + extension)
                self.assertEqual(self.readZipString(first, "content.xml"),
                                 self.readZipString(second, "content.xml"))

    def test_metadata_with_special_characters(self):
        for documentClass, extension, documentType in DOCUMENT_TYPES:
            with self.subTest(documentType):
                doc = documentClass()
                doc.meta.setTitle('A & B <"x">')
                doc.meta.setSubject("Subject & more")
                doc.meta.setDescription("Line <1>")
                doc.meta.addKeyword("R&D")
                doc.meta.addKeyword("R&D") # Duplicate keywords are ignored
                filename = self.export(doc, "meta" + extension)
                meta = ooolib.openDocument(filename).meta
                self.assertEqual(meta["title"], 'A & B <"x">')
                self.assertEqual(meta["subject"], "Subject & more")
                self.assertEqual(meta["description"], "Line <1>")
                self.assertEqual(meta["keywords"], ["R&D"])


class CalcTest(helpers.DocumentTestCase):
    def setUp(self):
        helpers.DocumentTestCase.setUp(self)
        self.doc = ooolib.Calc()
        self.table = self.doc.content.activeTable

    def test_column_ids(self):
        expected = {1: "A", 25: "Y", 26: "Z", 27: "AA", 52: "AZ", 53: "BA", 702: "ZZ", 703: "AAA", 16384: "XFD"}
        for col in expected:
            self.assertEqual(self.table.convertCellRowCol2Id(7, col), expected[col] + "7")

    def test_column_ids_round_trip(self):
        for col in range(1, 3000):
            cellId = self.table.convertCellRowCol2Id(12, col)
            self.assertEqual(self.table.convertCellId2RowCol(cellId), (12, col))

    def test_cell_text(self):
        self.table.setCellText(1, 1, "a\tb")
        self.table.setCellText(2, 1, "  two  spaces")
        self.table.setCellText(3, 1, "Tom & <Jerry>")
        filename = self.export(self.doc, "text.ods")
        self.assertEqual(ooolib.openDocument(filename).getText(), "a\tb\n  two  spaces\nTom & <Jerry>")

    def test_dates(self):
        self.table.setCellDate(1, 1, datetime.date(2026, 1, 2))
        self.table.setCellDate(2, 1, datetime.datetime(2026, 1, 2, 3, 4, 5))
        filename = self.export(self.doc, "dates.ods")
        content = self.readZipString(filename, "content.xml")
        self.assertIn('office:date-value="2026-01-02"', content)
        self.assertIn('office:date-value="2026-01-02T03:04:05"', content)
        self.assertNotIn("Date Error", content)

    def test_formula_with_quotes(self):
        self.table.setCellFormulaCustom(1, 1, 'of:=IF(1>0;"yes";"no")')
        filename = self.export(self.doc, "formula.ods")
        self.assertIn('table:formula="of:=IF(1&gt;0;&quot;yes&quot;;&quot;no&quot;)"',
                      self.readZipString(filename, "content.xml"))


class WriterTest(helpers.DocumentTestCase):
    def test_text_round_trip(self):
        doc = ooolib.Writer()
        doc.content.addHeading("Title & More", level=1)
        doc.content.addParagraph("  indented   gap\tTab")
        doc.content.addParagraph("line one\nline two", bold=True)
        filename = self.export(doc, "writer.odt")
        self.assertEqual(ooolib.openDocument(filename).getText(),
                         "Title & More\n  indented   gap\tTab\nline one\nline two")

    def test_lists_round_trip(self):
        doc = ooolib.Writer()
        doc.content.addListItem("One", numbered=True)
        doc.content.addListItem("One a", level=2)
        doc.content.addListItem("Two", numbered=True)
        doc.content.addParagraph("After")
        doc.content.addListItem("Bullet")
        filename = self.export(doc, "lists.odt")
        self.assertEqual(ooolib.openDocument(filename).getText(), "One\nOne a\nTwo\nAfter\nBullet")
        content = self.readZipString(filename, "content.xml")
        # Two top level lists, the first with a nested list inside its first item
        self.assertEqual(content.count('<text:list text:style-name="'), 3)
        self.assertRegex(content, r'<text:list-item><text:p text:style-name="Standard">One</text:p>'
                                  r'<text:list text:style-name="L\d+"><text:list-item>'
                                  r'<text:p text:style-name="Standard">One a</text:p>')

    def test_sub_list_kind_does_not_change_earlier_sub_lists(self):
        doc = ooolib.Writer()
        doc.content.addListItem("One", numbered=True)
        doc.content.addListItem("Bullet note", level=2)
        doc.content.addListItem("Two", numbered=True)
        doc.content.addListItem("Sub-step", level=2, numbered=True)
        content = doc.content.toString()
        styles = re.findall(r'<text:list text:style-name="(L\d+)">', content)
        self.assertEqual(len(styles), 3)
        self.assertNotEqual(styles[1], styles[2])
        self.assertIn('<text:list-level-style-number text:level="1"', content)
        self.assertIn('<text:list-level-style-bullet text:level="2"', content)

    def test_deep_list_item_without_parent(self):
        doc = ooolib.Writer()
        doc.content.addListItem("Deep", level=3)
        filename = self.export(doc, "deep.odt")
        content = self.readZipString(filename, "content.xml")
        self.assertEqual(content.count("<text:list>") + content.count("<text:list "), 3)
        self.assertEqual(ooolib.openDocument(filename).getText(), "Deep")

    def test_mixed_runs_and_paragraph_layout(self):
        doc = ooolib.Writer()
        doc.content.addParagraph([("Rent:", {"bold": True}), " <$800> & ", ("due", {"bold": True})],
                                 align="center", indent="0.5in", spaceAfter="0.1in", keepWithNext=True)
        doc.content.addParagraph("Plain")
        filename = self.export(doc, "runs.odt")
        self.assertEqual(ooolib.openDocument(filename).getText(), "Rent: <$800> & due\nPlain")
        content = self.readZipString(filename, "content.xml")
        # Both bold runs share one text style
        self.assertEqual(content.count('style:family="text"'), 1)
        self.assertIn('<text:span text:style-name="T1">Rent:</text:span><text:s/>&lt;$800&gt; &amp;<text:s/>', content)
        for attribute in ('fo:text-align="center"', 'fo:margin-left="0.5in"',
                          'fo:margin-bottom="0.1in"', 'fo:keep-with-next="always"'):
            self.assertIn(attribute, content)
        self.assertIn('<text:p text:style-name="Standard">Plain</text:p>', content)


class ImpressTest(helpers.DocumentTestCase):
    def test_slides_round_trip(self):
        doc = ooolib.Impress()
        doc.content.activeSlide.setTitle("First")
        doc.content.activeSlide.addContent("Point one")
        doc.content.addSlide("Second Slide")
        doc.content.activeSlide.setTitle("Second")
        filename = self.export(doc, "impress.odp")
        info = ooolib.openDocument(filename)
        self.assertEqual(len(info.pages), 2)
        self.assertEqual(info.pages[1].name, "Second Slide")
        self.assertEqual(info.pages[0].getText(), "First\nPoint one")
        titles = [shape.text for shape in info.getShapes() if shape.presentationClass == "title"]
        self.assertEqual(titles, ["First", "Second"])


if __name__ == "__main__":
    unittest.main()
