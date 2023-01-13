# ooolib-python

The ooolib-python library module is a tool to help programmers export data into OpenOffice.org or LibreOffice.org compatible documents using the Open Document Format.  The library module currently supports exporting documents as Calc spreadsheets.

# Installation
You can download the code, then add it to your path to test the code:

```
import sys
sys.path.append('PATH_TO_CODE')
import ooolib
doc = ooolib.Calc()
doc.export("test.ods")
```

This should create a blank document.  You can then look at the examples directory for more examples.

# History
I originally wrote ooolib-python in 2006 as I was switching from Perl to Python and needed to replace the ooolib-perl project I had originally created.  I wrote it and maintained it for a couple of years, then stopped working on it.  I recently needed to use the ooolib-python modules again, but discovered that the old code was a bit messy and was built for Python 2.  I decided to completely rewrite the code for Python 3 using a bit more of an object oriented approach.

I decided that I was more interested in working on exporting or creating spreadsheets, so that has been the focus of this version.

# Credits
Written by Joseph Colton <josephcolton@gmail.com>

# License
This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program. If not, see <https://www.gnu.org/licenses/>.
