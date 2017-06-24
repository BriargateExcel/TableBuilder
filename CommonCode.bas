Attribute VB_Name = "CommonCode"
Option Explicit

Public Function VBAMatch(ByVal Target As Variant, ByVal SearchRange As Range, Optional ByVal TreatAsString As Boolean = False) As Long

    On Error GoTo NotFound
    
    If IsDate(Target) And Not TreatAsString Then
        VBAMatch = Application.Match(CLng(Target), SearchRange, 0)
        Exit Function
    Else
        VBAMatch = Application.Match(Target, SearchRange, 0)
        Exit Function
    End If

NotFound:
    VBAMatch = 0
    
End Function ' VBAMatch


