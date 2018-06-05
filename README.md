# Excel-VBA
A collection of VBA scripts for data transformations, data cleaning, and/or metadata exports in Microsoft Excel or Microsoft Access.

# Quickstart
*Verticalize()* is the most popular and useful Excel script in this repository. To get started, open Excel; ensure that one of the sheet contains a rectangular dataset with field names in row 1 starting in cell A1, and data in the rows below; then copy, paste, and run the [verticalize()](https://github.com/jcoffeepot/Excel-VBA/blob/master/Excel%20VBA%20Scripts/sub%20verticalize.txt) subroutine from your Excel VBA editor.

# VBA Scripts for Excel
* sub verticalize(): Transform an Excel table from row-based to column-based.  Allows for the data to be more easily summarized with a Pivot Table.
* sub horizontalize(): Transform an Excel table from column-based to row-based.

# VBA Scripts  for Access
* exportAccessMetadata.txt: Old starter code for exporting metadata Access database metadata in a tabular format.
* exportAccessMetadata_v2.txt: Old starter code for exporting metadata Access database metadata in a tabular format.
* exportForSASProcFormat.txt: Old starter code for exporting Access database metadata for SAS Proc Format.

# License
The content of this repository is licensed according to an [Apache License 2.0](https://github.com/jcoffeepot/Excel-VBA/blob/master/LICENSE).
