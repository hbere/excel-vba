Attribute VB_Name = "basHorizontalize"
'*****************************************************
'Sub Horizontalize()
'https://github.com/jcoffeepot/Excel-VBA
'For use in Excel
'
'Results:
'   1.  Inputs a Vertical table with N dimension (categorization) columns, 1 data column, R rows, and S rows per entity
'   2.  Outputs a Vertical table with N+S columns and R/S rows
'
'Purpose:
'   1.  Transpose data from column-based to a row-based format.
'
'Usage/Assumptions:
'   1.  Input table must begin with dimension columns, reading left-to-right
'   2.  To the right of the column is 1 data column
'   3.  Each entity has the same number of rows
'   4.  Formatting doesn't need to come through.  (For the moment, formatting is not perserved through the transformation.)
'
'
'/**
' * Copyright 2018 jcoffeepot
' *
' * Licensed under the Apache License, Version 2.0 (the "License");
' * you may not use this file except in compliance with the License.
' * You may obtain a copy of the License at
' *
' *    http://www.apache.org/licenses/LICENSE-2.0
' *
' * Unless required by applicable law or agreed to in writing, software
' * distributed under the License is distributed on an "AS IS" BASIS,
' * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' * See the License for the specific language governing permissions and
' * limitations under the License.
' */
'
'*****************************************************
Sub Horizontalize()
On Error GoTo ErrorHandler:

    'Declare variables
    Dim InputRange As Range     'Range for input table
    Dim OutputRange As Range    'Range (upper-left corner) for output table

    Dim RowsPerEntity As Long   'Number of rows per entity
    Dim DataColFieldNo As Long  'Column containing the data
    Dim DimColTot As Long       'Number of dimension columns (Static)
    Dim EntityTot As Long       'Number of entities

    Dim EntityNo As Long        'Counter for entities
    Dim InRowStart As Long      'Pointer for input table topmost row to be copied in operation
    Dim InColStart As Long      'Pointer for input table leftmost column to be copied in operation
    Dim InRowEnd As Long        'Pointer for input table bottommost row to be copied in operation
    Dim InColEnd As Long        'Pointer for input table rightmost column to be copied in operation
    Dim OutRowStart As Long     'Pointer for output table paste location for operation
    Dim OutColStart As Long     'Pointer for output table paste location for operation

    Dim i As Long               'All-purpose counter

    'Initialize variables
    MsgBox "Note: This Macro operates under four assumptions: 1.  Input table must begin with dimension columns, reading left-to-right.  2.  The last column on the right contains data.  3.  The second-to-last column on the right contains field names corresponding to the data in the last column on the right.  4.  Each entity has the same number of rows."
    Set InputRange = Application.InputBox(Title:="Input Data Range", prompt:="Please select or enter the input table range.  It can be a normal Excel range (e.g. $A$1:$D$5) or a named range (e.g. MyTable).  For this macro to work correctly, the first row must contain *data*, not column names.", Type:=8)
    RowsPerEntity = Application.InputBox(Title:="Dimensions", prompt:="How many rows are there per entity?", Type:=1)
    Set OutputRange = Application.InputBox(Title:="Output Data Range.", prompt:="Please select the upper-left cell of where you'd like the output to appear.", Type:=8)
    DataColFieldNo = InputRange.Columns.Count - 1
    DimColTot = InputRange.Columns.Count - 2
    EntityTot = InputRange.Rows.Count / RowsPerEntity

    'Configure environment
    Application.ScreenUpdating = False

    'Transpose field names
    For i = 1 To DataColFieldNo - 1
        OutputRange.Cells(1, i) = "DIM" & i
    Next i
    For i = 1 To RowsPerEntity
        OutputRange.Cells(1, DataColFieldNo + i - 1) = InputRange.Cells(i, DataColFieldNo)
    Next i

    'Transpose data
    For EntityNo = 1 To EntityTot
        'DIMENSIONS
        If DimColTot > 0 Then
            'Define input range
            InRowStart = (EntityNo - 1) * RowsPerEntity + 1
            InRowEnd = InRowStart
            InColStart = 1
            InColEnd = DimColTot
            'Define output range
            OutRowStart = EntityNo + 1
            OutColStart = 1
            'Copy input range to output range
            Workbooks(InputRange.Worksheet.Parent.Name).Sheets(InputRange.Parent.Name).Activate     'Activate input range
            InputRange.Activate
            InputRange.Range(Cells(InRowStart, InColStart), Cells(InRowEnd, InColEnd)).Copy
            Workbooks(OutputRange.Worksheet.Parent.Name).Sheets(OutputRange.Parent.Name).Activate   'Activate output range
            OutputRange.Activate
            OutputRange.Cells(OutRowStart, OutColStart).PasteSpecial
        End If
        'DATA
        'Define input range
        InRowStart = (EntityNo - 1) * RowsPerEntity + 1
        InRowEnd = InRowStart + RowsPerEntity - 1
        InColStart = InputRange.Columns.Count
        InColEnd = InputRange.Columns.Count
        'Define output range
        OutRowStart = EntityNo + 1
        OutColStart = DimColTot + 1
        'Copy input range to output range
        Workbooks(InputRange.Worksheet.Parent.Name).Sheets(InputRange.Parent.Name).Activate     'Activate input range
        InputRange.Activate
        InputRange.Range(Cells(InRowStart, InColStart), Cells(InRowEnd, InColEnd)).Copy
        Workbooks(OutputRange.Worksheet.Parent.Name).Sheets(OutputRange.Parent.Name).Activate   'Activate output range
        OutputRange.Activate
        OutputRange.Cells(OutRowStart, OutColStart).PasteSpecial Transpose:=True
    Next EntityNo

Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description & ".  Please try again.", vbOKOnly

End Sub

