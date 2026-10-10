#!/usr/bin/python3
# Draw Example 04 - Connectors that stay attached to shapes

import sys
sys.path.append('../code')
import ooolib

# Create the document (10in x 7.5in landscape, coordinates in inches)
doc = ooolib.Draw()
doc.meta.setTitle("Draw Example 4")
doc.meta.setDescription("Connectors stay attached when the shapes are moved in LibreOffice.")
page = doc.content.activePage

page.addTextBox(0.5, 0.3, 9, 0.6, "Web Application", bold=True, fontsize=24,
    fontcolor="#1f3864", align="center")

# Shapes to connect
boxStyle = {"fillcolor": "#dae3f3", "strokecolor": "#4472c4", "strokewidth": 0.03, "bold": True}
page.addEllipse(0.5, 3.0, 1.75, 1.0, fillcolor="#fce4d6", strokecolor="#ed7d31", name="User", text="User")
page.addRectangle(3.25, 3.0, 2.0, 1.0, name="Web Server", text="Web Server", **boxStyle)
page.addRectangle(6.75, 1.25, 2.25, 1.0, name="Cache", text="Cache", **boxStyle)
page.addPolygon([(7.875, 4.75), (9.25, 5.5), (7.875, 6.25), (6.5, 5.5)],
    fillcolor="#e2f0d9", strokecolor="#548235", name="Database", text="Database", bold=True)

# A group: connectors can attach to shapes inside it
group = page.addGroup(name="Monitoring")
group.addRectangle(3.25, 5.75, 2.0, 1.0, fillcolor="#f2f2f2", strokecolor="#7f7f7f",
    linestyle="dash", name="Logs", text="Logs")

# Connectors take shape objects or shape names
page.addConnector("User", "Web Server", startside="right", endside="left",
    endarrow="arrow", text="HTTPS", fontsize=10)
page.addConnector("Web Server", "Cache", startside="top", endside="left",
    endarrow="arrow", strokecolor="#4472c4")
page.addConnector("Web Server", "Database", startside="right", endside="top",
    connectortype="lines", startarrow="circle", endarrow="arrow", strokecolor="#548235")
page.addConnector("Web Server", "Logs", connectortype="curve", linestyle="dash",
    endarrow="arrow", strokecolor="#7f7f7f")   # sides chosen by LibreOffice
page.addConnector("Cache", "Database", startside="bottom", endside="top",
    connectortype="line", linestyle="dot", strokecolor="#bf9000")

# Write out the document
doc.export("draw-example04.odg")
print("Created draw-example04.odg")
