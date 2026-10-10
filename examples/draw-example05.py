#!/usr/bin/python3
# Draw Example 05 - Images (PNG, JPEG, GIF, BMP, and SVG)

import sys, struct, zlib
sys.path.append('../code')
import ooolib

def makePNG(width, height, dpi=96):
    """Create a small gradient PNG image using only the standard library."""
    def chunk(chunkType, data):
        return (struct.pack(">I", len(data)) + chunkType + data +
                struct.pack(">I", zlib.crc32(chunkType + data) & 0xffffffff))
    rows = b""
    for y in range(height):
        rows += b"\x00" # No filter
        for x in range(width):
            rows += bytes([int(255 * x / width), int(255 * y / height), 200])
    pixelsPerMeter = int(round(dpi / 0.0254))
    return (b"\x89PNG\r\n\x1a\n" +
            chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)) +
            chunk(b"pHYs", struct.pack(">IIB", pixelsPerMeter, pixelsPerMeter, 1)) +
            chunk(b"IDAT", zlib.compress(rows)) +
            chunk(b"IEND", b""))

# Write a PNG file to use by file name (192 x 128 pixels at 96 dpi = 2in x 1.333in)
with open("gradient.png", "wb") as pngFile:
    pngFile.write(makePNG(192, 128))

# An SVG image kept in memory as bytes
logo = b"""<svg xmlns="http://www.w3.org/2000/svg" width="1.5in" height="1.5in" viewBox="0 0 100 100">
  <circle cx="50" cy="50" r="45" fill="#4472c4" stroke="#1f3864" stroke-width="4"/>
  <text x="50" y="62" font-size="32" text-anchor="middle" fill="white" font-family="sans-serif">ODF</text>
</svg>"""

# Create the document (10in x 7.5in landscape, coordinates in inches)
doc = ooolib.Draw()
doc.meta.setTitle("Draw Example 5")
page = doc.content.activePage
page.addTextBox(0.5, 0.3, 9, 0.6, "Images", bold=True, fontsize=24, fontcolor="#1f3864", align="center")

# Natural size: taken from the image's resolution
page.addImage("gradient.png", 0.5, 1.25, name="Natural Size")
page.addTextBox(0.5, 2.7, 2.5, 0.4, "Natural size", fontsize=11)

# Only a width: the height keeps the image's proportions
page.addImage("gradient.png", 3.25, 1.25, width=3, name="Width Only", strokecolor="#7f7f7f")
page.addTextBox(3.25, 3.35, 3, 0.4, "width=3 (with a border)", fontsize=11)

# Width and height: the image is stretched
page.addImage("gradient.png", 6.75, 1.25, width=2.75, height=1, name="Stretched")
page.addTextBox(6.75, 2.35, 2.75, 0.4, "width=2.75, height=1", fontsize=11)

# Images from bytes, rotated, and joined with a connector
page.addImage(logo, 1.0, 4.5, name="Logo")
page.addImage(logo, 6.5, 4.5, width=1.5, rotation=20, name="Rotated Logo")
page.addConnector("Logo", "Rotated Logo", startside="right", endside="left",
    endarrow="arrow", text="same image, stored once", fontsize=10)

# Write out the document
doc.export("draw-example05.odg")
print("Created draw-example05.odg")
