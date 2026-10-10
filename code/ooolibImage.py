#!/usr/bin/python3
#################################################################################
# ooolib-python - Python module for creating Open Document Format documents.    #
# Copyright (C) 2006-2026  Joseph Colton                                        #
#                                                                               #
# You can contact me by email at josephcolton@gmail.com                         #
#################################################################################

# Import python modules
import struct
import sys
import xml.etree.ElementTree as ElementTree

# Import ooolib-python modules
import ooolibUnits

###########################
# ooolib-python Images    #
###########################
# Resolution used when an image does not say
DEFAULT_DPI = 96.0

# Image formats: name -> (media type, file extension)
FORMATS = {
    "png": ("image/png", "png"),
    "jpeg": ("image/jpeg", "jpg"),
    "gif": ("image/gif", "gif"),
    "bmp": ("image/bmp", "bmp"),
    "svg": ("image/svg+xml", "svg"),
}

def readImageInfo(data):
    """Return information about an image from its bytes:
    {"format", "mediaType", "extension", "pixelWidth", "pixelHeight", "dpiX", "dpiY"}.

    Raises ValueError for formats that are not supported (PNG, JPEG, GIF, BMP, and SVG are)."""
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        info = readPNG(data)
    elif data[:3] == b"\xff\xd8\xff":
        info = readJPEG(data)
    elif data[:6] in (b"GIF87a", b"GIF89a"):
        width, height = struct.unpack("<HH", data[6:10])
        info = {"format": "gif", "pixelWidth": width, "pixelHeight": height}
    elif data[:2] == b"BM":
        info = readBMP(data)
    elif looksLikeSVG(data):
        info = readSVG(data)
    else:
        raise ValueError("Unsupported image format (expected PNG, JPEG, GIF, BMP, or SVG)")
    info["mediaType"], info["extension"] = FORMATS[info["format"]]
    info.setdefault("dpiX", DEFAULT_DPI)
    info.setdefault("dpiY", DEFAULT_DPI)
    return info

def getImageSize(info, unit="in"):
    """Return the natural (width, height) of an image in unit, using its resolution."""
    width = info["pixelWidth"] / info["dpiX"]
    height = info["pixelHeight"] / info["dpiY"]
    return (ooolibUnits.convertLength("%rin" % width, unit), ooolibUnits.convertLength("%rin" % height, unit))

##########
# Format #
##########
def readPNG(data):
    width, height = struct.unpack(">II", data[16:24])
    info = {"format": "png", "pixelWidth": width, "pixelHeight": height}
    # Look for the physical pixel size chunk (pHYs) before the image data
    position = 8
    while position + 8 <= len(data):
        length, chunkType = struct.unpack(">I4s", data[position:position + 8])
        if chunkType == b"pHYs" and length == 9:
            perUnitX, perUnitY, unit = struct.unpack(">IIB", data[position + 8:position + 17])
            if unit == 1 and perUnitX and perUnitY: # Pixels per meter
                info["dpiX"], info["dpiY"] = perUnitX * 0.0254, perUnitY * 0.0254
            break
        if chunkType in (b"IDAT", b"IEND"): break
        position += 12 + length
    return info

def readJPEG(data):
    info = {"format": "jpeg"}
    position = 2
    while position + 4 <= len(data):
        if data[position] != 0xFF:
            position += 1
            continue
        marker = data[position + 1]
        if marker in (0xD8, 0x01) or 0xD0 <= marker <= 0xD7 or marker == 0xFF:
            position += 1 if marker == 0xFF else 2
            continue
        length = struct.unpack(">H", data[position + 2:position + 4])[0]
        segment = data[position + 4:position + 2 + length]
        if marker == 0xE0 and segment[:5] == b"JFIF\x00" and len(segment) >= 12:
            units, densityX, densityY = struct.unpack(">BHH", segment[7:12])
            if densityX and densityY and units in (1, 2):
                scale = 1.0 if units == 1 else 2.54 # Dots per inch or per centimeter
                info["dpiX"], info["dpiY"] = densityX * scale, densityY * scale
        # Start of frame markers hold the size (C4, C8, and CC are other tables)
        if 0xC0 <= marker <= 0xCF and marker not in (0xC4, 0xC8, 0xCC):
            height, width = struct.unpack(">HH", segment[1:5])
            info["pixelWidth"], info["pixelHeight"] = width, height
            return info
        position += 2 + length
    raise ValueError("Unable to find the size of the JPEG image")

def readBMP(data):
    width, height = struct.unpack("<ii", data[18:26])
    info = {"format": "bmp", "pixelWidth": width, "pixelHeight": abs(height)}
    if len(data) >= 46:
        perMeterX, perMeterY = struct.unpack("<ii", data[38:46])
        if perMeterX > 0 and perMeterY > 0:
            info["dpiX"], info["dpiY"] = perMeterX * 0.0254, perMeterY * 0.0254
    return info

def looksLikeSVG(data):
    start = data[:1024].lstrip().lower()
    return start.startswith(b"<svg") or (start.startswith(b"<?xml") and b"<svg" in data[:4096].lower())

def readSVG(data):
    root = ElementTree.fromstring(data)
    width, height = root.get("width"), root.get("height")
    viewBox = (root.get("viewBox") or "").replace(",", " ").split()
    # SVG sizes without units are CSS pixels (96 per inch)
    widthIn = ooolibUnits.convertLength(width, "in", "px") if width and "%" not in width else None
    heightIn = ooolibUnits.convertLength(height, "in", "px") if height and "%" not in height else None
    if (widthIn is None or heightIn is None) and len(viewBox) == 4:
        viewWidth, viewHeight = float(viewBox[2]), float(viewBox[3])
        if widthIn is None and heightIn is not None and viewHeight:
            widthIn = heightIn * viewWidth / viewHeight
        elif heightIn is None and widthIn is not None and viewWidth:
            heightIn = widthIn * viewHeight / viewWidth
        else:
            widthIn, heightIn = viewWidth / DEFAULT_DPI, viewHeight / DEFAULT_DPI
    if widthIn is None or heightIn is None:
        raise ValueError("Unable to find the size of the SVG image")
    return {"format": "svg", "pixelWidth": widthIn * DEFAULT_DPI, "pixelHeight": heightIn * DEFAULT_DPI}


###################
# Execute to Test #
###################
if __name__ == "__main__":
    for filename in sys.argv[1:]:
        with open(filename, "rb") as imageFile:
            info = readImageInfo(imageFile.read())
        print(filename, info, "%.2fin x %.2fin" % getImageSize(info))
