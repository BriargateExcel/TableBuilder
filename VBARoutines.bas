Attribute VB_Name = "VBARoutines"
Option Explicit
' source = http://www.cpearson.com/Excel/VBE.aspx

Sub AddModuleToProject(ByVal ModuleName As String, ByVal VBProj As VBIDE.VBProject)
    Dim VBComp As VBIDE.VBComponent

    Set VBComp = VBProj.VBComponents.Add(vbext_ct_StdModule)
    VBComp.Name = ModuleName
End Sub

Sub AddLineToModule(ByVal Line As String, ByVal CodeMod As VBIDE.CodeModule, ByVal LineNum As Long)
    If CodeMod.CountOfLines = LineNum Then
        CodeMod.InsertLines LineNum, Line
    Else
        CodeMod.InsertLines LineNum, vbCrLf
        CodeMod.InsertLines LineNum, Line
    End If

    If CodeMod.CountOfLines >= LineNum + 1 Then
        CodeMod.DeleteLines LineNum + 1, CodeMod.CountOfLines - LineNum
    End If
End Sub

Function CopyModuleCode(ModuleName As String, _
    FromVBProject As VBIDE.VBProject, _
    ToVBProject As VBIDE.VBProject, _
    OverwriteExisting As Boolean) As Boolean
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' CopyModule
    ' This function copies a module from one VBProject to
    ' another. It returns True if successful or False
    ' if an error occurs.
    '
    ' Parameters:
    ' --------------------------------
    ' FromVBProject         The VBProject that contains the module
    '                       to be copied.
    '
    ' ToVBProject           The VBProject into which the module is
    '                       to be copied.
    '
    ' ModuleName            The name of the module to copy.
    '
    ' OverwriteExisting     If True, the VBComponent named ModuleName
    '                       in ToVBProject will be removed before
    '                       importing the module. If False and
    '                       a VBComponent named ModuleName exists
    '                       in ToVBProject, the code will return
    '                       False.
    '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim VBComp As VBIDE.VBComponent
    Dim FName As String
    Dim CompName As String
    Dim S As String
    Dim SlashPos As Long
    Dim ExtPos As Long
    Dim TempVBComp As VBIDE.VBComponent

    '''''''''''''''''''''''''''''''''''''''''''''
    ' Do some housekeeping validation.
    '''''''''''''''''''''''''''''''''''''''''''''
    If FromVBProject Is Nothing Then
        CopyModuleCode = False
        Exit Function
    End If

    If Trim(ModuleName) = vbNullString Then
        CopyModuleCode = False
        Exit Function
    End If

    If ToVBProject Is Nothing Then
        CopyModuleCode = False
        Exit Function
    End If

    If FromVBProject.Protection = vbext_pp_locked Then
        CopyModuleCode = False
        Exit Function
    End If

    If ToVBProject.Protection = vbext_pp_locked Then
        CopyModuleCode = False
        Exit Function
    End If

    On Error Resume Next
    Set VBComp = FromVBProject.VBComponents(ModuleName)
    If Err.Number <> 0 Then
        CopyModuleCode = False
        Exit Function
    End If

    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' FName is the name of the temporary file to be
    ' used in the Export/Import code.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    FName = Environ("Temp") & "\" & ModuleName & ".bas"
    If OverwriteExisting = True Then
        ''''''''''''''''''''''''''''''''''''''
        ' If OverwriteExisting is True, Kill
        ' the existing temp file and remove
        ' the existing VBComponent from the
        ' ToVBProject.
        ''''''''''''''''''''''''''''''''''''''
        If Dir(FName, vbNormal + vbHidden + vbSystem) <> vbNullString Then
            Err.Clear
            Kill FName
            If Err.Number <> 0 Then
                CopyModuleCode = False
                Exit Function
            End If
        End If
        With ToVBProject.VBComponents
            .Remove .Item(ModuleName)
        End With
    Else
        '''''''''''''''''''''''''''''''''''''''''
        ' OverwriteExisting is False. If there is
        ' already a VBComponent named ModuleName,
        ' exit with a return code of False.
        ''''''''''''''''''''''''''''''''''''''''''
        Err.Clear
        Set VBComp = ToVBProject.VBComponents(ModuleName)
        If Err.Number <> 0 Then
            If Err.Number = 9 Then
                ' module doesn't exist. ignore error.
            Else
                ' other error. get out with return value of False
                CopyModuleCode = False
                Exit Function
            End If
        End If
    End If

    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Do the Export and Import operation using FName
    ' and then Kill FName.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    FromVBProject.VBComponents(ModuleName).Export FileName:=FName

    '''''''''''''''''''''''''''''''''''''
    ' Extract the module name from the
    ' export file name.
    '''''''''''''''''''''''''''''''''''''
    SlashPos = InStrRev(FName, "\")
    ExtPos = InStrRev(FName, ".")
    CompName = Mid(FName, SlashPos + 1, ExtPos - SlashPos - 1)

    ''''''''''''''''''''''''''''''''''''''''''''''
    ' Document modules (SheetX and ThisWorkbook)
    ' cannot be removed. So, if we are working with
    ' a document object, delete all code in that
    ' component and add the lines of FName
    ' back in to the module.
    ''''''''''''''''''''''''''''''''''''''''''''''
    Set VBComp = Nothing
    Set VBComp = ToVBProject.VBComponents(CompName)

    If VBComp Is Nothing Then
        ToVBProject.VBComponents.Import FileName:=FName
    Else
        If VBComp.Type = vbext_ct_Document Then
            ' VBComp is destination module
            Set TempVBComp = ToVBProject.VBComponents.Import(FName)
            ' TempVBComp is source module
            With VBComp.CodeModule
                .DeleteLines 1, .CountOfLines
                S = TempVBComp.CodeModule.Lines(1, TempVBComp.CodeModule.CountOfLines)
                .InsertLines 1, S
            End With
            On Error GoTo 0
            ToVBProject.VBComponents.Remove TempVBComp
        End If
    End If
    Kill FName
    CopyModuleCode = True
End Function ' CopyModule

