# NARA-Project

2018/5/6

Main goal
Transforming metadata entered by archivists into spreadsheets into National Archives Catalog-compliant XML


How to Run "transform.py"
Run "transform.py" in Python IDE after library "pandas" and "tkinter" are already correctly installed.It will pop up a window for the user to select the file they want to convert. The file must be standard spreadsheet used by NAtional Archives Catalog for a successful conversion.The output XML file will be saved in the same directory where input file locates.


Notes
	1. fixed the matching issue. It will work with or without the whitespace in headers.
	2. check all yellow (required) columns and remind the user if some of them are blank.
	3. if a grey column has no data, xml file won't contain the corresponding tags.
	4. handled both the case where the spreadsheet contains additional columns "variantControlNumber"  and the case where the spreadsheet doesn't contain. Yet it only works when headers are variantControlNumberType, variantControlNumberNum, and variantControlNumberNote.
