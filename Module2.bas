Attribute VB_Name = "Module2"
Option Explicit
Sub color()
    Sheets("Sheet1").Activate
    Range("A1").Select
    Dim nc As Integer
    Dim nr As Integer
    Dim x As Integer
    Dim y As Integer
    Dim cell As Range

    nr = 275 ' Number of rows
    nc = 126 ' Number of columns

    For x = 1 To nr
        For y = 1 To nc
            Set cell = Sheets("Sheet1").Cells(x, y)
            If cell.Value = "Y" Then
                cell.Interior.color = RGB(128, 0, 128) ' Purple color
            End If
        Next y
    Next x
End Sub
