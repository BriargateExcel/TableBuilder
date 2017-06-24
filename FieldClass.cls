VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FieldClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pTableName As String
Private pFieldLabel As String
Private pFieldAbbrev As String
Private pVBAType As String

Public Sub AddNewField(ByVal TableName As String, ByVal FieldLabel As String, ByVal FieldAbbrev As String, ByVal VBAType As String)
    ' ToDo: check the table name exists
    pTableName = TableName
    pFieldLabel = FieldLabel
    pFieldAbbrev = FieldAbbrev
    ' ToDo: Need to check for valid types
    pVBAType = VBAType
End Sub ' AddNewField


Property Get TableName() As String
    TableName = pTableName
End Property

Property Let TableName(ByVal TN As String)
    pTableName = TN
End Property


Property Get FieldLabel() As String
    FieldLabel = pFieldLabel
End Property

Property Let FieldLabel(ByVal FL As String)
    pFieldLabel = FL
End Property

Property Get FieldAbbrev() As String
    FieldAbbrev = pFieldAbbrev
End Property

Property Let FieldAbbrev(ByVal FA As String)
    pFieldAbbrev = FA
End Property


Property Get VBAType() As String
    VBAType = pVBAType
End Property

Property Let VBAType(ByVal VT As String)
    pVBAType = VT
End Property


