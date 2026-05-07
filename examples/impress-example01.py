#!/usr/bin/python3
# Impress Example 01 - Demonstrates slides, titles, content text, and formatting

import sys
sys.path.append('../code')
import ooolib

# Create the document
doc = ooolib.Impress()

# Document metadata
doc.meta.setTitle("Impress Example 1")
doc.meta.setSubject("ooolib-python Impress demonstration")
doc.meta.setDescription("Shows slides, titles, content, and text formatting.")
doc.meta.addKeyword("ooolib")
doc.meta.addKeyword("impress")

# --- Slide 1: Title slide with background color ---
doc.content.activeSlide.setTitle("Welcome to ooolib Impress")
doc.content.activeSlide.addContent("An ooolib-python demonstration")
doc.content.activeSlide.addContent("Open Document Presentation format", italic=True)
doc.content.activeSlide.setBackgroundColor("#1f3864")

# --- Slide 2: Text formatting ---
# addContent(text, bold=False, italic=False, underline=False, fontsize=None, fontcolor=None)
doc.content.addSlide("Slide2")
doc.content.activeSlide.setTitle("Text Formatting")
doc.content.activeSlide.addContent("Normal content text")
doc.content.activeSlide.addContent("Bold content text", bold=True)
doc.content.activeSlide.addContent("Italic content text", italic=True)
doc.content.activeSlide.addContent("Underlined content text", underline=True)
doc.content.activeSlide.addContent("Bold and italic text", bold=True, italic=True)

# --- Slide 3: Font sizes ---
doc.content.addSlide("Slide3")
doc.content.activeSlide.setTitle("Font Sizes")
doc.content.activeSlide.addContent("Default size (28pt)")
doc.content.activeSlide.addContent("Larger text at 36pt", fontsize="36pt")
doc.content.activeSlide.addContent("Smaller text at 20pt", fontsize="20pt")
doc.content.activeSlide.addContent("Small text at 16pt", fontsize="16pt")

# --- Slide 4: Colors ---
doc.content.addSlide("Slide4")
doc.content.activeSlide.setTitle("Font Colors")
doc.content.activeSlide.addContent("Red text", fontcolor="#cc0000")
doc.content.activeSlide.addContent("Green text", fontcolor="#006600")
doc.content.activeSlide.addContent("Blue text", fontcolor="#0000cc")
doc.content.activeSlide.addContent("Orange bold text", fontcolor="#cc6600", bold=True)
doc.content.activeSlide.setBackgroundColor("#f0f4ff")

# Write out the document
doc.export("impress-example01.odp")
print("Created impress-example01.odp")
