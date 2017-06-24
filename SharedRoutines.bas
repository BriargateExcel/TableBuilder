Attribute VB_Name = "SharedRoutines"
Option Explicit

Public Function FindLastRow(ByVal ColLetter As String, ByVal RowNumber As Long, ByVal Sheet As Worksheet) As Long
    Dim RegionRow As Long: RegionRow = Sheet.Range(ColLetter & RowNumber).CurrentRegion.Rows.Count
    Dim ColumnRow As Long: ColumnRow = Sheet.Range(ColLetter & Sheet.Rows.Count).End(xlUp).Row
    Dim ColumnNumber As Integer: ColumnNumber = Sheet.Range(ColLetter & 1).Column
    Dim I As Long
    Dim CurrentCell As Range

    If RegionRow = ColumnRow Then
        FindLastRow = ColumnRow
    Else
        For I = Application.Max(ColumnRow, RegionRow) To Application.Min(ColumnRow, RegionRow) Step -1
            Set CurrentCell = Sheet.Cells(I, ColumnNumber)
            If Not IsEmpty(CurrentCell) Then
                FindLastRow = I
                Exit For
            End If
        Next I
    End If
End Function ' FindLastRow

Public Function FindLastColumn(ByVal RowNumber As Integer, _
    ByVal Sheet As Worksheet) As Integer
    
    FindLastColumn = Sheet.Cells(RowNumber, Sheet.Columns.Count).End(xlToLeft).Column
End Function ' FindLastColumn

Public Function ConvertToLetter(ByVal iCol As Long) As String
    Dim iAlpha As Integer
    Dim iRemainder As Integer

    iAlpha = Int(iCol / 27)
    iRemainder = iCol - (iAlpha * 26)
    
    If iAlpha > 0 Then
        ConvertToLetter = Chr(iAlpha + 64)
    End If
    
    If iRemainder > 0 Then
        ConvertToLetter = ConvertToLetter & Chr(iRemainder + 64)
    End If

End Function ' ConvertToLetter

Public Sub FMT_AutoFit(ByVal Rng As Range)
    Rng.AutoFit
End Sub ' FMT_AutoFit

Public Sub FMT_SortSheets(ByVal WB As Workbook, ByVal StartSht As Integer)
'
' StartSht=1 to sort all sheets
' StartSht=2 to skip the sheet right after the table of contents
'
Dim I As Integer
Dim J As Integer

    For I = StartSht To WB.Sheets.Count - 1
      For J = I + 1 To WB.Sheets.Count
        If WB.Sheets(J).Name < WB.Sheets(I).Name Then
           WB.Sheets(J).Move Before:=WB.Sheets(I)
        End If
      Next J
    Next I
End Sub ' FMT_SortSheets

Public Sub ShowAll(ByVal Sht As Worksheet)
Dim CurrentSheet As Worksheet
Dim Vis As Boolean
Dim Tables As Scripting.Dictionary
Dim StartCell As String
Dim Tbl As Variant

    Vis = Sht.Visible
    Set CurrentSheet = ThisWorkbook.ActiveSheet
    Sht.Activate
    Set Tables = Sheets.Item(Sht.Name).TableDict
    For Each Tbl In Tables
        StartCell = ConvertToLetter(Tables.Item(Tbl).FirstCol) & "1"
        Sht.Range(StartCell).Activate
        On Error Resume Next
        Sht.ShowAllData
        On Error GoTo 0
    Next Tbl
    Sht.Visible = Vis
    Sht.Visible = Vis
    CurrentSheet.Activate
End Sub ' ShowAll

Public Sub AddReferences(ByVal Wkbk As Workbook)
Const VBA = "{000204EF-0000-0000-C000-000000000046}"
Const Excel = "{00020813-0000-0000-C000-000000000046}"
Const stdole = "{00020430-0000-0000-C000-000000000046}"
Const Office = "{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}"
Const MSForms = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}"
Const VBIDE = "{0002E157-0000-0000-C000-000000000046}"
Const AdHocReportingExcelClientLib = "{8E47F3A2-81A4-468E-A401-E1DEBBAE2D8D}"
Const MSComCtl2 = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}"
Const Scripting = "{420B2830-E718-11CF-893D-00A0C9054228}"

Dim ID As Object

 On Error Resume Next

 Set ID = Wkbk.VBProject.References

 ID.AddFromGuid VBA, 0, 0
 ID.AddFromGuid Excel, 0, 0
 ID.AddFromGuid stdole, 0, 0
 ID.AddFromGuid Office, 0, 0
 ID.AddFromGuid MSForms, 0, 0
 ID.AddFromGuid VBIDE, 0, 0
 ID.AddFromGuid AdHocReportingExcelClientLib, 0, 0
 ID.AddFromGuid MSComCtl2, 0, 0
 ID.AddFromGuid Scripting, 0, 0

End Sub ' AddReferences
