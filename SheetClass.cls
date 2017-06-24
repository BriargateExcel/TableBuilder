VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pSheetName As String
Private pTbl As TableClass
Private pFirstCol As Integer
Private pNextTopRow As Long
Private pNextCol As Integer
Private pTableDict As Scripting.Dictionary

Public Sub AddNewSheet(ByVal SheetName As String, ByVal Tbl As TableClass)
    pSheetName = SheetName
    Set pTbl = Tbl
    Set pTableDict = New Scripting.Dictionary
    pTableDict.Add Tbl.TableName, Tbl
    pFirstCol = 1
    pNextTopRow = 1
    pNextCol = Tbl.NumCols + 2
End Sub

Public Sub AddNewTableToSheet(ByVal Tbl As TableClass)
    If Not pTableDict.Exists(Tbl.TableName) Then
        pTableDict.Add Tbl.TableName, Tbl
        pNextTopRow = 1
        Tbl.FirstCol = pNextCol
        pNextCol = pNextCol + Tbl.NumCols + 1
    Else
        Stop
    End If
End Sub


Property Get SheetName() As String
    SheetName = pSheetName
End Property

Property Let SheetName(ByVal SN As String)
    pSheetName = SN
End Property


Property Get Tbl() As TableClass
    Set Tbl = pTbl
End Property

Property Let Tbl(ByVal TC As TableClass)
    Set pTbl = TC
End Property


Property Get FirstCol() As Integer
    FirstCol = pFirstCol
End Property

Property Let FirstCol(ByVal FC As Integer)
    pFirstCol = FC
End Property


Property Get NextTopRow() As Long
    NextTopRow = pNextTopRow
End Property

Property Let NextTopRow(ByVal NTR As Long)
    pNextTopRow = NTR
End Property


Property Get TableDict() As Scripting.Dictionary
    Set TableDict = pTableDict
End Property

Property Let TableDict(ByVal TD As Scripting.Dictionary)
    Set pTableDict = TD
End Property


Property Get NextCol() As Integer
    NextCol = pNextCol
End Property

Property Let NextCol(ByVal NC As Integer)
    pNextCol = NC
End Property



