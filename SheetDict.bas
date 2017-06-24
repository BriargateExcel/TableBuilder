Attribute VB_Name = "SheetDict"
Option Explicit

Public Sheets As Scripting.Dictionary

Public Sub BuildSheets(ByVal Wkbk As Workbook, ByVal Tables As Scripting.Dictionary)
Dim Sht As Variant
Dim Tbl As Variant
Dim Fld As Variant
Dim Sheet As SheetClass
Dim WS As Worksheet
Dim NextCol As Integer
Dim RngStr As String

    Set Sheets = New Scripting.Dictionary
    
    For Each Tbl In Tables
        If Sheets.Exists(Tables.Item(Tbl).SheetName) Then
            Sht = Tables.Item(Tbl).SheetName
            If Not Sheets.Item(Sht).TableDict.Exists(Tbl) Then
                Sheets.Item(Sht).AddNewTableToSheet Tables.Item(Tbl)
                NextCol = Tables.Item(Tbl).FirstCol
                For Each Fld In Tables.Item(Tbl).Fields
                    Wkbk.Worksheets(Sht).Cells(1, NextCol) = Fld.FieldLabel
                    NextCol = NextCol + 1
                Next Fld
            Else
                Stop
            End If
        Else
            Set Sheet = New SheetClass
            Sheet.AddNewSheet Tables.Item(Tbl).SheetName, Tables.Item(Tbl)
            Sheets.Add Tables.Item(Tbl).SheetName, Sheet
            Set WS = Wkbk.Worksheets.Add
            WS.Name = Tables.Item(Tbl).SheetName
            NextCol = 1
            For Each Fld In Tables.Item(Tbl).Fields
                WS.Cells(1, NextCol) = Fld.FieldLabel
                NextCol = NextCol + 1
            Next Fld
        End If
    Next Tbl
    
    Wkbk.Activate
    For Each Sht In Wkbk.Worksheets
        If Not Sheets.Exists(Sht.Name) Then
            Application.DisplayAlerts = False
            Sht.Delete
            Application.DisplayAlerts = True
        End If
    Next Sht
    
    For Each Sht In Sheets
        If Sheets.Exists(Sht) Then
            Set Sheet = Sheets.Item(Sht)
            
            For Each Tbl In Sheet.TableDict
                RngStr = "$" & ConvertToLetter(Sheets.Item(Sht).TableDict.Item(Tbl).FirstCol) & "$1:$"
                RngStr = RngStr & _
                    ConvertToLetter(Sheets.Item(Sht).TableDict.Item(Tbl).FirstCol + _
                        Sheets.Item(Sht).TableDict.Item(Tbl).NumCols - 1) & _
                    "$1"
                Wkbk.Worksheets(Sht).Activate
                Wkbk.Worksheets(Sht).ListObjects.Add(xlSrcRange, Range(RngStr) _
                    , , xlYes).Name = Sheets.Item(Sht).TableDict.Item(Tbl).TableName
            Next Tbl
            FMT_AutoFit Worksheets(Sht).Cells.EntireColumn
            Wkbk.Worksheets(Sht).Activate
            Wkbk.Worksheets(Sht).Cells(2, 1).Select
            ActiveWindow.FreezePanes = True
        End If
    Next Sht
    
    FMT_SortSheets Wkbk, 1
    
End Sub
