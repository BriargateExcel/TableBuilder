Attribute VB_Name = "CodeBuilder"
Option Explicit

Sub ButtonBuildCode(control As IRibbonControl)
    BuildCode
End Sub

Public Sub BuildCode()
    Dim NewProj As VBProject
    Dim ThisProj As VBProject
    Dim Wkbk As Workbook
    Dim DataWkbk As Workbook
    Dim intChoice As Integer
    Dim strPath As String
    Dim Tbl As Variant
    Dim Tables As Scripting.Dictionary
    
    Set ThisProj = ThisWorkbook.VBProject
    
    Set Wkbk = Workbooks.Add
    
    Set NewProj = Wkbk.VBProject
    
    CopyModuleCode "CommonCode", ThisProj, NewProj, False
    
    Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
    intChoice = Application.FileDialog(msoFileDialogOpen).Show
    If intChoice <> 0 Then
        strPath = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
        Set DataWkbk = Workbooks.Open(strPath, , True)
    Else
        Stop
    End If

    Set Tables = New Scripting.Dictionary
    Set Tables = TBLTBL_Initialize(DataWkbk)

    DataWkbk.Close SaveChanges:=False
    
    Set NewProj = Wkbk.VBProject

    For Each Tbl In Tables
        BuildModule Tbl, NewProj, Tables
    Next Tbl

    AddReferences Wkbk
    
    BuildSheets Wkbk, Tables

    ThisWorkbook.Activate
    
End Sub ' BuildCode

Private Sub BuildModule(ByVal TableName As String, VBProj As VBProject, ByVal Dict As Scripting.Dictionary)
    Dim ShtName As String:      ShtName = TBLTBL_Get_SheetName_TableName(TableName, Dict)
    Dim ModName As String:      ModName = TBLTBL_Get_ModName_TableName(TableName, Dict)
    Dim TblAbbrev As String:    TblAbbrev = TBLTBL_Get_TblAbbrev_TableName(TableName, Dict)
    Dim PrimKey As String:      PrimKey = TBLTBL_Get_PrimKey_TableName(TableName, Dict)
    Dim UpperPref As String:    UpperPref = UCase(TBLTBL_Get_CodePref_TableName(TableName, Dict)) & "_"
    Dim LowerPref As String:    LowerPref = LCase(UpperPref)
    Const DQUOTE = """"     ' one " character
    Const SQUOTE = "'"      ' one ' character
    Dim Line As String
    Dim LineNum As Long
    Dim Key As String
    Dim Fields As Collection
    Dim Fld As Variant
    Dim CodeMod As VBIDE.CodeModule
    
    AddModuleToProject ModName, VBProj

    Set CodeMod = VBProj.VBComponents(ModName).CodeModule

'Private smp_SampleAbbrevSheet As Worksheet
'Private <LowerPref><TblAbbrev>Sheet As Worksheet
    Line = "Private " & LowerPref & TblAbbrev & "Sheet As Worksheet"
    AddLineToModule Line, CodeMod, 3

'Private smp_SampleAbbrevTable As ListObject
'Private <LowerPref><TblAbbrev>Table As ListObject
    Line = "Private " & LowerPref & TblAbbrev & "Table As ListObject"
    AddLineToModule Line, CodeMod, 4

'Const smp_SheetName = "SampleSheet"
'Const <LowerPref>SheetName = "<ShtName>"
    Line = "Const " & LowerPref & "SheetName = " & DQUOTE & ShtName & DQUOTE
    AddLineToModule Line, CodeMod, 6

'Const smp_TableName = "SampleTable"
'Const <LowerPref>TableName = "<TableName>"
    Line = "Const " & LowerPref & "TableName = " & DQUOTE & TableName & DQUOTE
    AddLineToModule Line, CodeMod, 7

'Const smp_Smp1Title = "Sample 1"
'Const <LowerPref><FieldLabel>Title = "<FldLabel>"
    LineNum = 8
    Set Fields = TBLTBL_Get_Fields_TableName(TableName, Dict)
    For Each Fld In Fields
        Line = "Const " & LowerPref & Fld.FieldAbbrev & "Title = " & DQUOTE & Fld.FieldLabel & DQUOTE
        AddLineToModule Line, CodeMod, LineNum
        LineNum = LineNum + 1
    Next Fld
    
'Private smp_Smp1Rng As Range
'Private <LowerPref><FieldAbbrev>Rng As Range
    LineNum = LineNum + 1
    For Each Fld In Fields
        Line = "Private " & LowerPref & Fld.FieldAbbrev & "Rng as Range"
        AddLineToModule Line, CodeMod, LineNum
        LineNum = LineNum + 1
    Next Fld

'Private smp_Smp1Col As Integer
'Private <LowerPref><FldAbbrev>Col As Integer
    LineNum = LineNum + 1
    For Each Fld In Fields
        Line = "Private " & LowerPref & Fld.FieldAbbrev & "Col as Integer"
        AddLineToModule Line, CodeMod, LineNum
        LineNum = LineNum + 1
    Next Fld

'Public Sub SMP_Initialize(ByVal Wkbk as Workbook)
'Public Sub <UpperPref>Initialize(ByVal Wkbk as Workbook)
    LineNum = LineNum + 1
    Line = "Public Sub " & UpperPref & "Initialize(ByVal Wkbk as Workbook)"
    AddLineToModule Line, CodeMod, LineNum

'    Set smp_SampleAbbrevSheet = ThisWorkbook.Worksheets(smp_SheetName)
'    Set <LowerPref><TblAbbrev>Sheet = Wkbk.Worksheets(<LowerPref>SheetName)
    LineNum = LineNum + 2
    Line = "    Set " & LowerPref & TblAbbrev & "Sheet = Wkbk.Worksheets(" & LowerPref & "SheetName)"
    AddLineToModule Line, CodeMod, LineNum
    
'    Set smp_SampleAbbrevTable = smp_SampleAbbrevSheet.ListObjects(smp_TableName)
'    Set <LowerPreef><TblAbbrev>Table = <LowerPref><TblAbbrev>Sheet.ListObjects(<LowerPref>TableName)
    LineNum = LineNum + 2
    Line = "    Set " & LowerPref & TblAbbrev & "Table = " & LowerPref & TblAbbrev & "Sheet.ListObjects(" & LowerPref & "TableName)"
    AddLineToModule Line, CodeMod, LineNum

    LineNum = LineNum + 1

'    Set smp_Smp1Rng = smp_SampleAbbrevTable.ListColumns(smp_Smp1Title).DataBodyRange
'    Set <LowerPref><FldAbbrev>Rng = <LowerPref><FldAbbrev>Table.ListColumns(<LowerPref><FldAbbrev>Title).DataBodyRange
    LineNum = LineNum + 1
    For Each Fld In Fields
        Line = "    Set " & LowerPref & Fld.FieldAbbrev & "Rng = " & LowerPref & TblAbbrev & _
            "Table.ListColumns(" & LowerPref & Fld.FieldAbbrev & "Title).DataBodyRange"
        AddLineToModule Line, CodeMod, LineNum
        LineNum = LineNum + 1
    Next Fld

'    smp_Smp1Col = VBAMatch(smp_Smp1Title, smp_SampleAbbrevTable.HeaderRowRange)
'    <LowerPref><FldAbbrev>Col = VBAMatch(<LowerPref><FldAbbrev>Title, <LowerPref><TblAbbrev>Table.HeaderRowRange)
    LineNum = LineNum + 1
    For Each Fld In Fields
        Line = "    " & LowerPref & Fld.FieldAbbrev & "Col = VBAMatch(" & _
            LowerPref & Fld.FieldAbbrev & "Title, " & LowerPref & TblAbbrev & "Table.HeaderRowRange)"
        AddLineToModule Line, CodeMod, LineNum
        LineNum = LineNum + 1
    Next Fld
    
'End Sub ' SMP_Initialize
'End Sub ' <UpperPref>Initialize
    LineNum = LineNum + 1
    Line = "End Sub " & SQUOTE & " " & UpperPref & "Initialize"
    AddLineToModule Line, CodeMod, LineNum

'Public Function SMP_NumCols() As Integer
'Public Function <UpperPref>NumCols() As Integer
    LineNum = LineNum + 2
    Line = "Public Function " & UpperPref & "NumCols() As Integer"
    AddLineToModule Line, CodeMod, LineNum

'    SMP_NumCols = smp_SampleAbbrevTable.ListColumns.Count
'    <UpperPref>NumCols = <LowerPref><TblAbbrev>Table.ListColumns.Count
    LineNum = LineNum + 1
    Line = "    " & UpperPref & "NumCols = " & LowerPref & TblAbbrev & "Table.ListColumns.Count"
    AddLineToModule Line, CodeMod, LineNum

'End Function 'SMP_NumCols
'End Function '<UpperPref>NumCols
    LineNum = LineNum + 1
    Line = "End Function " & SQUOTE & UpperPref & "NumCols"
    AddLineToModule Line, CodeMod, LineNum

'Public Function SMP_NumRows() As Integer
'Public Function <UpperPref>NumRows() As Integer
    LineNum = LineNum + 2
    Line = "Public Function " & UpperPref & "NumRows() As Integer"
    AddLineToModule Line, CodeMod, LineNum

'    SMP_NumRows = smp_SampleAbbrevTable.ListRows.Count
'    <UpperPref>NumRows = <LowerPref><TblAbbrev>Table.ListRows.Count
    LineNum = LineNum + 1
    Line = "    " & UpperPref & "NumRows = " & LowerPref & TblAbbrev & "Table.ListRows.Count"
    AddLineToModule Line, CodeMod, LineNum

'End Function 'SMP_NumRows
'End Function '<UpperPref>NumRows
    LineNum = LineNum + 1
    Line = "End Function " & SQUOTE & UpperPref & "NumRows"
    AddLineToModule Line, CodeMod, LineNum

    LineNum = LineNum + 3
    
    If PrimKey = "None" Then
        ' Do nothing; nothing to test for existence
    Else
        Key = TBLTBL_Get_PrimKey_TableName(TableName, Dict)
    End If
    
    If Key <> "" Then
'Public Function SMP_Exists_SampleAbbrev(ByVal Smp1Name As String) As Boolean
'Public Function <UpperPref>Exists_<TblAbbrev>(ByVal <Key>As String) As Boolean
        Line = "Public Function " & UpperPref & "Exists_" & TblAbbrev & "(ByVal " & Key & " As String) As Boolean"
        AddLineToModule Line, CodeMod, LineNum

'    SMP_Exists_SampleAbbrev = (VBAMatch(Smp1Name, smp_Smp1Rng) <> 0)
'    <UpperPref>Exists_<TblAbbrev> = (VBAMatch(<Key>, <LowerPref><Key>Rng) <> 0)
        LineNum = LineNum + 1
        Line = "    " & UpperPref & "Exists_" & TblAbbrev & " = (VBAMatch(" & Key & ", " & LowerPref & Key & "Rng) <> 0)"
        AddLineToModule Line, CodeMod, LineNum

'End Function 'SMP_Exists_SampleAbbrev
'End Function '<UpperPref>Exists_<TblAbbrev>
        LineNum = LineNum + 1
        Line = "End Function " & SQUOTE & UpperPref & "Exists_" & TblAbbrev
        AddLineToModule Line, CodeMod, LineNum
    End If
    
    LineNum = LineNum + 3
    For Each Fld In Fields
        If PrimKey = "None" Then
            Key = TableName & "_" & Fld.FieldAbbrev
        Else
            Key = TBLTBL_Get_PrimKey_TableName(TableName, Dict)
        End If
    
    
'Public Function SMP_Get_Smp2_Smp1(ByVal Smp1 As String) As String
'Public Function <UpperPref>Get_<FldAbbrev>_<Key>(ByVal <Key> As String) As Fld.<VBAType>
        If (Fld.FieldAbbrev <> Key Or (TableName & "_" & Fld.FieldAbbrev) <> Key) And PrimKey <> "None" Then
            Line = "Public Function " & UpperPref & "Get_" & Fld.FieldAbbrev & "_" & Key & _
                "(ByVal " & Key & " As String) As " & Fld.VBAType
            AddLineToModule Line, CodeMod, LineNum

'    Dim RowNum As Long
'    Dim RowNum As Long
            LineNum = LineNum + 1
            Line = "    Dim RowNum As Long"
            AddLineToModule Line, CodeMod, LineNum

'    RowNum = VBAMatch(Smp1, smp_Smp1Rng, True)
'    RowNum = VBAMatch(<Key>, <LowerPref><Key>Rng, True)
            LineNum = LineNum + 1
            Line = "    RowNum = VBAMatch(" & Key & ", " & LowerPref & Key & "Rng, True)"
            AddLineToModule Line, CodeMod, LineNum

'    SMP_Get_Smp2_Smp1 = smp_Smp2Rng(RowNum)
'    <UpperPref>Get_<FldAbbrev>_<Key> = <LowerPref><FldAbbrev>Rng(RowNum)
            LineNum = LineNum + 1
            Line = "    " & UpperPref & "Get_" & Fld.FieldAbbrev & "_" & Key & " = " & _
                LowerPref & Fld.FieldAbbrev & "Rng(RowNum)"
            AddLineToModule Line, CodeMod, LineNum

'End Function 'SMP_Get_Smp2_Smp1
'End Function '<UpperPref>Get_<FldAbbrev>_<Key>
            LineNum = LineNum + 1
            Line = "End Function " & SQUOTE & UpperPref & "Get_" & Fld.FieldAbbrev & "_" & Key
            AddLineToModule Line, CodeMod, LineNum

            LineNum = LineNum + 2

        End If

'Public Sub SMP_Let_Smp2_Smp1(ByVal Smp1 As String, ByVal NewVal as String)
'Public Sub <UpperPref>Get_<FldAbbrev>_<Key>(ByVal <Key> As String, ByVal NewVal as <VBAType>)
        If PrimKey = "None" Then
            LineNum = LineNum - 3
        Else
            Line = "Public Sub " & UpperPref & "Let_" & Fld.FieldAbbrev & "_" & Key & _
                "(ByVal " & Key & " As String, " & _
                "ByVal NewVal as " & Fld.VBAType & ")"
            AddLineToModule Line, CodeMod, LineNum

'    Dim RowNum As Long
'    Dim RowNum As Long
            LineNum = LineNum + 1
            Line = "    Dim RowNum As Long"
            AddLineToModule Line, CodeMod, LineNum

'    RowNum = VBAMatch(Smp1, smp_Smp1Rng, True)
'    RowNum = VBAMatch(<Key>, <LowerPref><Key>Rng, True)
            LineNum = LineNum + 1
            Line = "    RowNum = VBAMatch(" & Key & ", " & LowerPref & Key & "Rng, True)"
            AddLineToModule Line, CodeMod, LineNum

'    smp_Smp2Rng(RowNum) = NewVal
'    <LowerPref><FldAbbrev>Rng(RowNum) = NewVal
            LineNum = LineNum + 1
            Line = "    " & LowerPref & Fld.FieldAbbrev & "Rng(RowNum) = NewVal"
            AddLineToModule Line, CodeMod, LineNum

'End Sub 'SMP_Let_Smp2_Smp1
'End Sub '<UpperPref>Let_<FldAbbrev>_<Key>
            LineNum = LineNum + 1
            Line = "End Sub " & SQUOTE & UpperPref & "Let_" & Fld.FieldAbbrev & "_" & Key
            AddLineToModule Line, CodeMod, LineNum

            LineNum = LineNum + 3
        End If
    Next Fld

'Public Function SMP_CheckStructure() As Boolean
'Public Function <UpperPref>CheckStructure() As Boolean
    Line = "Public Function " & UpperPref & "CheckStructure() As Boolean"
    AddLineToModule Line, CodeMod, LineNum

'    Dim Header As Range
'    Dim Header As Range
    LineNum = LineNum + 1
    Line = "    Dim Header As Range"
    AddLineToModule Line, CodeMod, LineNum

'    Set Header = smp_SampleAbbrevTable.HeaderRowRange
'    Set Header = <LowerPref><TblAbbrev>Table.HeaderRowRange
    LineNum = LineNum + 1
    Line = "    Set Header = " & LowerPref & TblAbbrev & "Table.HeaderRowRange"
    AddLineToModule Line, CodeMod, LineNum

    For Each Fld In Fields
    
'    If VBAMatch("Sample 1", Header, True) = 0 Then
'    If VBAMatch("<FldLabel>", Header, True) = 0 Then
        LineNum = LineNum + 1
        Line = "    If VBAMatch(" & DQUOTE & Fld.FieldLabel & DQUOTE & ", Header, True) = 0 Then"
        AddLineToModule Line, CodeMod, LineNum

'        MsgBox "Sample 1 not found"
'        MsgBox "<FldLabel> not found"
        LineNum = LineNum + 1
        Line = "        MsgBox" & DQUOTE & SQUOTE & Fld.FieldLabel & SQUOTE & _
            " not found in " & TableName & " header" & DQUOTE
        AddLineToModule Line, CodeMod, LineNum

'    End If
'    End If
        LineNum = LineNum + 1
        Line = "    End If"
        AddLineToModule Line, CodeMod, LineNum
    Next Fld

'End Function 'SMP_CheckStructure
'End Function '<UpperPref>CheckStructure
    LineNum = LineNum + 1
    Line = "End Function " & SQUOTE & UpperPref & "CheckStructure"
    AddLineToModule Line, CodeMod, LineNum

End Sub ' BuildModule


