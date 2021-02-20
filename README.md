The QuestradeReconcile macro is python code that uses the Questrade application programming interface (API) to fetch account, position, balance, equity, and 30 day activity into a LibreOffice spreadsheet file.

This script is meant to be run infrequently to help provide a dashboard view when, for example, re-balancing a portfolio. It is not meant to track real time market conditions.

![Figure 1: Run the QuestradeReconcile Python Macro](Documentation/RunQuestradeMacro.png?raw=True "Figure 1: Run the QuestradeReconcile Python Macro")

This project was last tested on February 12th, 2021 and continues  to be fully functional.

[PDF Documentation](Documentation/QuestradeMacroDocumentation.pdf?raw=True)

This repository is not meant to offer a turnkey application but instead is a useful reference. Nobody ought to allow any application, including this one, access to the Questrade API without first being able to completely understand and audit the code.

[Wiki Roadmap](https://github.com/kerouac01850/questrade-reconcile/wiki)

**Prerequistes**

<ul>
   <li>[LibreOffice for spreadsheet functionality](https://www.libreoffice.org/download/download/)</li>
   <li>[7Zip for development](https://www.7-zip.org/download.html)</li>
</ul>

**Notes**

Python and Basic macros are packed into the Sample spreadsheet file by default. LibreOffice is required to make use of the Sample spreadsheet and it represents a good starting point for a fully customized Questrade tracking application. The spreadsheet file is a standard ODS archive generated by LibreOffice.

LibreOffice has a built-in editor for Basic but not for Python macros. The following is necessary when making Python changes to the spreadsheet file.

<pre>
W:\Questrade>"C:\Program Files\7-Zip\7z.exe" l Sample.ods

7-Zip 18.05 (x64) : Copyright (c) 1999-2018 Igor Pavlov : 2018-04-30

Scanning the drive for archives:
1 file, 24113 bytes (24 KiB)

Listing archive: Sample.ods

Path = Sample.ods
Type = zip
Physical Size = 24113

   Date      Time    Attr         Size   Compressed  Name
------------------- ----- ------------ ------------  ------------------------
2021-02-20 17:05:18 .....           46           46  mimetype
2021-02-20 17:05:18 .....         8182         1546  Basic\Standard\QuestDashboard.xml
2021-02-20 17:05:18 .....          355          218  Basic\Standard\script-lb.xml
2021-02-20 17:05:18 .....          338          211  Basic\script-lc.xml
2021-02-20 17:05:18 .....        36391         8667  Scripts\python\QuestradeReconcile.py
2021-02-20 17:05:18 D....            0            0  Configurations2\images\Bitmaps
2021-02-20 17:05:18 D....            0            0  Configurations2\floater
2021-02-20 17:05:18 D....            0            0  Configurations2\accelerator
2021-02-20 17:05:18 D....            0            0  Configurations2\menubar
2021-02-20 17:05:18 D....            0            0  Configurations2\progressbar
2021-02-20 17:05:18 D....            0            0  Configurations2\popupmenu
2021-02-20 17:05:18 D....            0            0  Configurations2\statusbar
2021-02-20 17:05:18 D....            0            0  Configurations2\toolbar
2021-02-20 17:05:18 D....            0            0  Configurations2\toolpanel
2021-02-20 17:05:18 .....          899          261  manifest.rdf
2021-02-20 17:05:18 .....        55733         4187  styles.xml
2021-02-20 17:05:18 .....          879          441  meta.xml
2021-02-20 17:05:18 .....        31506         3672  content.xml
2021-02-20 17:05:18 .....        20970         1930  settings.xml
2021-02-20 17:05:18 .....         1577          362  META-INF\manifest.xml
------------------- ----- ------------ ------------  ------------------------
2021-02-20 17:05:18             156876        21541  11 files, 9 folders
</pre>

To extract and make changes to the QuestradeReconcile.py file embedded within the Sample.ods file use 7zip or any other archive program:

<pre>
W:\Questrade>"C:\Program Files\7-Zip\7z.exe" x -aoa Sample.ods Scripts\python\QuestradeReconcile.py

7-Zip 18.05 (x64) : Copyright (c) 1999-2018 Igor Pavlov : 2018-04-30

Scanning the drive for archives:
1 file, 24113 bytes (24 KiB)

Extracting archive: Sample.ods

Path = Sample.ods
Type = zip
Physical Size = 24113

Everything is Ok

Size:       36391
Compressed: 24113
</pre>

To overwrite the existing Scripts\python\QuestradeReconcile.py file in the Sample.ods file with your changes using 7zip:

<pre>
W:\Questrade>"C:\Program Files\7-Zip\7z.exe" u Sample.ods Scripts\python\QuestradeReconcile.py

7-Zip 18.05 (x64) : Copyright (c) 1999-2018 Igor Pavlov : 2018-04-30

Open archive: Sample.ods
--
Path = Sample.ods
Type = zip
Physical Size = 24113

Scanning the drive:
1 file, 36391 bytes (36 KiB)

Updating archive: Sample.ods

Keep old data in archive: 9 folders, 11 files, 156876 bytes (154 KiB)
Add new data to archive: 0 files, 0 bytes


Files read from disk: 0
Archive size: 24113 bytes (24 KiB)
Everything is Ok
</pre>

QuestradeReconcile is free software: you can redistribute and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

QuestradeReconcile is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details. You should have received a copy of the GNU General Public License along with this program.  If not, see https://www.gnu.org/licenses/.
