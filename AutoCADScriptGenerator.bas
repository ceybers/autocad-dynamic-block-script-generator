Attribute VB_Name = "AutoCADScriptGenerator"
Option Explicit

Const BLOCK_NAME As String = "MY_BLOCK"
Const POINTS_COL As Integer = 1 ' Index of column where points (x,y coordinates) are stored
Const FIRST_ATTRIBUTE_COLUMN = 2 ' Where your first attribute column is in the table
Const ATTRIBUTE_COUNT As Integer = 3 ' Number of attributes (assume first column is points)
Const FILENAME As String = "drawing1.scr"

Public Sub GenerateScript()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim dbr As Range
    Dim i As Integer, attribIdx As Integer
    
    ' Check if table exists
    Set ws = ThisWorkbook.Worksheets(1)
    If ws.ListObjects.Count = 0 Then
        MsgBox "Cannot find table"
        Exit Sub
    End If
    Set lo = ws.ListObjects(1)
    Set dbr = lo.DataBodyRange
    
    ' Check constants versus columns in the table
    If lo.ListColumns.Count < (FIRST_ATTRIBUTE_COLUMN + ATTRIBUTE_COUNT) Then
        MsgBox "Check FIRST_ATTRIBUTE_COLUMN and COL_COUNT versus table"
        Exit Sub
    End If
    
    ' Clear existing file
    Open FILENAME For Output As #1
    Close #1
    
    ' Loop through table and write each row to file in AutoCAD command line script format
    Open FILENAME For Append As #1
    For i = 1 To lo.DataBodyRange.Rows.Count
        Print #1, "-INSERT" ' Insert dynamic block from command line
        Print #1, BLOCK_NAME ' Name of dynamic block
        Print #1, dbr.Cells(i, POINTS_COL) 'v(1, 1) ' Coordinates in x,y format
        Print #1, "1" ' Scale X
        Print #1, "1" ' Scale Y
        Print #1, "0" ' Rotation
        For attribIdx = FIRST_ATTRIBUTE_COLUMN To FIRST_ATTRIBUTE_COLUMN + ATTRIBUTE_COUNT - 1
            Print #1, dbr.Cells(i, attribIdx) ' Loop through all attributes
        Next attribIdx
        
        ' Alternatively, reference columns by name
        ' Print #1, lo.ListColumns("cadAttribute1").DataBodyRange.Cells(i, 1).Value
    Next i
    Close #1
End Sub
