#!/usr/bin/python3
# Writer Example 01 - Demonstrates paragraphs, headings, and text formatting

import sys
sys.path.append('../code')
import ooolib

# Create the document
doc = ooolib.Writer()

# Document metadata
doc.meta.setTitle("Writer Example 1")
doc.meta.setSubject("ooolib-python Writer demonstration")
doc.meta.setDescription("Shows headings, paragraphs, and text formatting options.")
doc.meta.addKeyword("ooolib")
doc.meta.addKeyword("writer")

# Headings (level 1-4)
doc.content.addHeading("Document Title (Heading 1)", level=1)
doc.content.addHeading("Chapter One (Heading 2)", level=2)
doc.content.addHeading("Section 1.1 (Heading 3)", level=3)
doc.content.addHeading("Sub-section 1.1.1 (Heading 4)", level=4)

# Plain paragraph
doc.content.addParagraph("This is a normal paragraph with no formatting applied.")

# Text style options
# addParagraph(text, bold=False, italic=False, underline=False, fontsize=None, fontcolor=None)
doc.content.addHeading("Text Formatting", level=2)

doc.content.addParagraph("This paragraph is bold.", bold=True)
doc.content.addParagraph("This paragraph is italic.", italic=True)
doc.content.addParagraph("This paragraph is underlined.", underline=True)
doc.content.addParagraph("This paragraph is bold, italic, and underlined.", bold=True, italic=True, underline=True)

# Font sizes (specified as point strings)
doc.content.addHeading("Font Sizes", level=2)
doc.content.addParagraph("This paragraph uses 10pt font.", fontsize="10pt")
doc.content.addParagraph("This paragraph uses 14pt font.", fontsize="14pt")
doc.content.addParagraph("This paragraph uses 18pt font.", fontsize="18pt")
doc.content.addParagraph("This paragraph uses 24pt bold font.", fontsize="24pt", bold=True)

# Font colors (hex color strings)
doc.content.addHeading("Font Colors", level=2)
doc.content.addParagraph("This paragraph uses red text.", fontcolor="#cc0000")
doc.content.addParagraph("This paragraph uses green text.", fontcolor="#006600")
doc.content.addParagraph("This paragraph uses blue text.", fontcolor="#0000cc")
doc.content.addParagraph("This paragraph combines color and bold.", fontcolor="#cc6600", bold=True)

# Write out the document
doc.export("writer-example01.odt")
print("Created writer-example01.odt")
