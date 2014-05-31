
Option Explicit


Sub LoggingConsole_1()
    Dim IsLoaded  As Boolean
    Dim i         As Integer
    
    If (Not IsLoadedActionsNET()) Then
        MsgBox "Actions.NET.xll ist nicht geladen"
    Else
        For i = 1 To 1000
            Call Application.Run("LoggingConsoleLogInfo", "Info Message from " & ThisWorkbook.Name)
        Next
        Call Application.Run("LoggingConsoleShow")
    End If
End Sub

Sub LoggingConsole_Hide()
    If (IsLoadedActionsNET()) Then
        Call Application.Run("LoggingConsoleHide")
    End If
End Sub


Function IsLoadedActionsNET() As Boolean
    
    Dim xlAddIn   As AddIn
    Dim xlAddins2 As AddIns2
    Dim IsLoaded  As Boolean
    
    Const AddInName_1 As String = "Actions.NET-AddIn-packed.xll"
    Const AddInName_2 As String = "Actions.NET-AddIn.xll"
    Const AddInName_3 As String = "Actions.NET.xll"
    
    Set xlAddins2 = Application.AddIns2
    
    For Each xlAddIn In xlAddins2
        If ((xlAddIn.Name = AddInName_3) Or (xlAddIn.Name = AddInName_2) Or (xlAddIn.Name = AddInName_1)) Then
            IsLoaded = xlAddIn.IsOpen
            Exit For
        End If
    Next
    
    IsLoadedActionsNET = IsLoaded
End Function


