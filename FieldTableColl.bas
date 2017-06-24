Attribute VB_Name = "FieldTableColl"
Option Explicit

Private fldtbl_FieldTableSheet As Worksheet

Private fldtbl_FieldTableTable As ListObject

Private fldtbl_FldAry() As Variant

Const fldtbl_SheetName = "Fields"
Const fldtbl_TableName = "FieldTable"

Const fldtbl_TableNameTitle = "Table Name"
Const fldtbl_FieldLabelTitle = "Field Label"
Const fldtbl_FieldAbbrevTitle = "Field Abbreviation"
Const fldtbl_VBATypeTitle = "VBA Type"

Private fldtbl_TableNameCol As Integer
Private fldtbl_FieldLabelCol As Integer
Private fldtbl_FieldAbbrevCol As Integer
Private fldtbl_VBATypeCol As Integer

Public Sub FLDTBL_Initialize(ByVal Wkbk As Workbook)
    Dim LeftCol As Integer
    Dim TopRow As Long
    Dim LastRow As Long
    Dim LastCol As Integer
    Dim FldRng As Range

    Set fldtbl_FieldTableSheet = Wkbk.Worksheets(fldtbl_SheetName)

    Set fldtbl_FieldTableTable = fldtbl_FieldTableSheet.ListObjects(fldtbl_TableName)
    
    fldtbl_TableNameCol = VBAMatch(fldtbl_TableNameTitle, fldtbl_FieldTableTable.HeaderRowRange)
    fldtbl_FieldLabelCol = VBAMatch(fldtbl_FieldLabelTitle, fldtbl_FieldTableTable.HeaderRowRange)
    fldtbl_FieldAbbrevCol = VBAMatch(fldtbl_FieldAbbrevTitle, fldtbl_FieldTableTable.HeaderRowRange)
    fldtbl_VBATypeCol = VBAMatch(fldtbl_VBATypeTitle, fldtbl_FieldTableTable.HeaderRowRange)

    TopRow = fldtbl_FieldTableTable.Range.Row
    LeftCol = fldtbl_FieldTableTable.Range.Column

    LastRow = FindLastRow(ConvertToLetter(LeftCol), TopRow, fldtbl_FieldTableSheet)
    LastCol = FindLastColumn(TopRow, fldtbl_FieldTableSheet)
    
    With fldtbl_FieldTableSheet
        Set FldRng = .Range(.Cells(TopRow + 1, LeftCol), .Cells(LastRow, LastCol))
    End With
    
    fldtbl_FldAry = FldRng
    
End Sub ' FLDTBL_Initialize

Public Function FLDTBL_Get_Coll_TableName(ByVal TableName As String) As Collection
    Dim I As Long
    Dim Field As FieldClass
    Dim Coll As Collection

    Set Coll = New Collection
    
    For I = 1 To UBound(fldtbl_FldAry, 1)
        If TableName = fldtbl_FldAry(I, fldtbl_TableNameCol) Then
            Set Field = New FieldClass
            Field.AddNewField fldtbl_FldAry(I, fldtbl_TableNameCol), _
                fldtbl_FldAry(I, fldtbl_FieldLabelCol), _
                fldtbl_FldAry(I, fldtbl_FieldAbbrevCol), _
                fldtbl_FldAry(I, fldtbl_VBATypeCol)
                
            Coll.Add Field
        End If
    Next I
    
    Set FLDTBL_Get_Coll_TableName = Coll

End Function

Public Function FLDTBL_CheckStructure() As Boolean
    Dim Header As Range

    Set Header = fldtbl_FieldTableTable.HeaderRowRange

    If VBAMatch("Table Name", Header, True) = 0 Then
        MsgBox "Table Name not found"
    End If

    If VBAMatch("Field Label", Header, True) = 0 Then
        MsgBox "Field Label not found"
    End If

    If VBAMatch("Field Abbreviation", Header, True) = 0 Then
        MsgBox "Field Abbreviation not found"
    End If

    If VBAMatch("VBA Type", Header, True) = 0 Then
        MsgBox "VBA Type not found"
    End If

End Function 'FLDTBL_CheckStructure


