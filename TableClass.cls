VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pTableName As String
Private pSheetName As String
Private pModName As String
Private pTblAbbrev As String
Private pCodePref As String
Private pPrimKey As String
Private pFields As Collection
Private pFirstCol As Integer
Private pNumCols As Integer

Public Sub AddNewTable( _
    ByVal TableName As String, _
    ByVal SheetName As String, _
    ByVal ModName As String, _
    ByVal TblAbbrev As String, _
    ByVal CodePref As String, _
    ByVal PrimKey As String, _
    ByVal Fields As Collection)
    
    pTableName = TableName
    pSheetName = SheetName
    pModName = ModName
    pTblAbbrev = TblAbbrev
    pCodePref = CodePref
    pPrimKey = PrimKey
    Set pFields = Fields
    pNumCols = Fields.Count
    pFirstCol = 1
    
End Sub ' AddNewField


Property Get TableName() As String
    TableName = pTableName
End Property

Property Let TableName(ByVal TN As String)
    pTableName = TN
End Property


Property Get SheetName() As String
    SheetName = pSheetName
End Property

Property Let SheetName(ByVal SN As String)
    pSheetName = SN
End Property


Property Get ModName() As String
    ModName = pModName
End Property

Property Let ModName(ByVal MN As String)
    pModName = MN
End Property


Property Get TblAbbrev() As String
    TblAbbrev = pTblAbbrev
End Property

Property Let TblAbbrev(ByVal TA As String)
    pTblAbbrev = TA
End Property


Property Get CodePref() As String
    CodePref = pCodePref
End Property

Property Let CodePref(ByVal CP As String)
    pCodePref = CP
End Property


Property Get PrimKey() As String
    PrimKey = pPrimKey
End Property

Property Let PrimKey(ByVal PK As String)
    pPrimKey = PK
End Property


Property Get Fields() As Collection
    Set Fields = pFields
End Property

Property Let Fields(ByVal Fld As Collection)
    Set pFields = Fld
End Property


Property Get FirstCol() As Integer
    FirstCol = pFirstCol
End Property

Property Let FirstCol(ByVal FC As Integer)
    pFirstCol = FC
End Property


Property Get NumCols() As Integer
    NumCols = pNumCols
End Property

Property Let NumCols(ByVal NC As Integer)
    pNumCols = NC
End Property



