Attribute VB_Name = "TableTableDict"
Option Explicit

Private tbltbl_TableTableSheet As Worksheet

Private tbltbl_TableTableTable As ListObject

Const tbltbl_SheetName = "Tables"
Const tbltbl_TableName = "TableTable"

Const tbltbl_TableNameTitle = "Table Name"
Const tbltbl_SheetNameTitle = "Sheet Name"
Const tbltbl_ModNameTitle = "Module Name"
Const tbltbl_TblAbbrevTitle = "Table Abbreviation"
Const tbltbl_CodePrefTitle = "Code Prefix"
Const tbltbl_PrimKeyTitle = "Primary Key"

Private tbltbl_TableNameCol As Integer
Private tbltbl_SheetNameCol As Integer
Private tbltbl_ModNameCol As Integer
Private tbltbl_TblAbbrevCol As Integer
Private tbltbl_CodePrefCol As Integer
Private tbltbl_PrimKeyCol As Integer

Public Function TBLTBL_Initialize(ByVal Wkbk As Workbook) As Scripting.Dictionary
    Dim TblAry() As Variant
    Dim LeftCol As Integer
    Dim TopRow As Long
    Dim LastRow As Long
    Dim LastCol As Integer
    Dim TblRng As Range
    Dim I As Long
    Dim Tbl As TableClass
    Dim Fields As Collection
    Dim Dict As Scripting.Dictionary

    Set tbltbl_TableTableSheet = Wkbk.Worksheets(tbltbl_SheetName)

    Set tbltbl_TableTableTable = tbltbl_TableTableSheet.ListObjects(tbltbl_TableName)

    tbltbl_TableNameCol = VBAMatch(tbltbl_TableNameTitle, tbltbl_TableTableTable.HeaderRowRange)
    tbltbl_SheetNameCol = VBAMatch(tbltbl_SheetNameTitle, tbltbl_TableTableTable.HeaderRowRange)
    tbltbl_ModNameCol = VBAMatch(tbltbl_ModNameTitle, tbltbl_TableTableTable.HeaderRowRange)
    tbltbl_TblAbbrevCol = VBAMatch(tbltbl_TblAbbrevTitle, tbltbl_TableTableTable.HeaderRowRange)
    tbltbl_CodePrefCol = VBAMatch(tbltbl_CodePrefTitle, tbltbl_TableTableTable.HeaderRowRange)
    tbltbl_PrimKeyCol = VBAMatch(tbltbl_PrimKeyTitle, tbltbl_TableTableTable.HeaderRowRange)

    TopRow = tbltbl_TableTableTable.Range.Row
    LeftCol = tbltbl_TableTableTable.Range.Column

    LastRow = FindLastRow(ConvertToLetter(LeftCol), TopRow, tbltbl_TableTableSheet)
    LastCol = FindLastColumn(TopRow, tbltbl_TableTableSheet)
    
    With tbltbl_TableTableSheet
        Set TblRng = .Range(.Cells(TopRow + 1, LeftCol), .Cells(LastRow, LastCol))
    End With
    
    TblAry = TblRng
    
    Set Dict = New Scripting.Dictionary
    
    FLDTBL_Initialize Wkbk
    
    For I = 1 To UBound(TblAry, 1)
        Set Tbl = New TableClass
        Set Fields = New Collection
        Set Fields = FLDTBL_Get_Coll_TableName(TblAry(I, tbltbl_TableNameCol))
        
        Tbl.AddNewTable _
            TblAry(I, tbltbl_TableNameCol), _
            TblAry(I, tbltbl_SheetNameCol), _
            TblAry(I, tbltbl_ModNameCol), _
            TblAry(I, tbltbl_TblAbbrevCol), _
            TblAry(I, tbltbl_CodePrefCol), _
            TblAry(I, tbltbl_PrimKeyCol), _
            Fields
            
        Dict.Add TblAry(I, tbltbl_TableNameCol), Tbl
    Next I
    
    Set TBLTBL_Initialize = Dict

End Function ' TBLTBL_Initialize

Public Function TBLTBL_NumCols() As Integer
    TBLTBL_NumCols = tbltbl_TableTableTable.ListColumns.Count
End Function 'TBLTBL_NumCols

Public Function TBLTBL_Get_SheetName_TableName(ByVal TableTableName As String, ByVal Dict As Scripting.Dictionary) As String
    If Dict.Exists(TableTableName) Then
        TBLTBL_Get_SheetName_TableName = Dict.Item(TableTableName).SheetName
    Else
        Stop
    End If
End Function 'TBLTBL_Get_SheetName_TableName

Public Function TBLTBL_Get_ModName_TableName(ByVal TableTableName As String, ByVal Dict As Scripting.Dictionary) As String
    If Dict.Exists(TableTableName) Then
        TBLTBL_Get_ModName_TableName = Dict.Item(TableTableName).ModName
    Else
        Stop
    End If
End Function 'TBLTBL_Get_ModName_TableName

Public Function TBLTBL_Get_TblAbbrev_TableName(ByVal TableTableName As String, ByVal Dict As Scripting.Dictionary) As String
    If Dict.Exists(TableTableName) Then
        TBLTBL_Get_TblAbbrev_TableName = Dict.Item(TableTableName).TblAbbrev
    Else
        Stop
    End If
End Function 'TBLTBL_Get_TblAbbrev_TableName

Public Function TBLTBL_Get_CodePref_TableName(ByVal TableTableName As String, ByVal Dict As Scripting.Dictionary) As String
    If Dict.Exists(TableTableName) Then
        TBLTBL_Get_CodePref_TableName = Dict.Item(TableTableName).CodePref
    Else
        Stop
    End If
End Function 'TBLTBL_Get_CodePref_TableName

Public Function TBLTBL_Get_PrimKey_TableName(ByVal TableTableName As String, ByVal Dict As Scripting.Dictionary) As String
    If Dict.Exists(TableTableName) Then
        TBLTBL_Get_PrimKey_TableName = Dict.Item(TableTableName).PrimKey
    Else
        Stop
    End If
End Function 'TBLTBL_Get_PrimKey_TableName

Public Function TBLTBL_Get_Fields_TableName(ByVal TableTableName As String, ByVal Dict As Scripting.Dictionary) As Collection
    If Dict.Exists(TableTableName) Then
        Set TBLTBL_Get_Fields_TableName = Dict.Item(TableTableName).Fields
    Else
        Stop
    End If
End Function 'TBLTBL_Get_PrimKey_TableName

Public Function TBLTBL_CheckStructure() As Boolean
    Dim Header As Range

    Set Header = tbltbl_TableTableTable.HeaderRowRange

    If VBAMatch("Table Name", Header, True) = 0 Then
        MsgBox "Table Name not found"
    End If

    If VBAMatch("Sheet Name", Header, True) = 0 Then
        MsgBox "Sheet Name not found"
    End If

    If VBAMatch("Module Name", Header, True) = 0 Then
        MsgBox "Module Name not found"
    End If

    If VBAMatch("Table Abbreviation", Header, True) = 0 Then
        MsgBox "Table Abbreviation not found"
    End If

    If VBAMatch("Table Prefix", Header, True) = 0 Then
        MsgBox "Table Prefix not found"
    End If

    If VBAMatch("Primary Key", Header, True) = 0 Then
        MsgBox "Primary Key not found"
    End If

End Function 'TBLTBL_CheckStructure



