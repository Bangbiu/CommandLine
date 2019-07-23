Attribute VB_Name = "Utility"
Option Explicit

Public CmdDealer As New CommandDealer
Public CurPath As String
Public CurInput As String
Public History As String

Public Function Middle(Src As String, Start As Long, Finish As Long)

End Function

Public Function LastDirectory() As String
    Dim Index As Integer
    For Index = Len(CurPath) - 4 To 1 Step -1
        If Mid(CurPath, Index, 1) = "\" Then
            LastDirectory = Left(CurPath, Index) & "> "
            Exit Function
        End If
    Next Index
    LastDirectory = CurPath
End Function

Public Sub Main()
    If Command = Empty Then
        CurPath = App.Path & "\> "
    ElseIf Dir(Command, vbDirectory) <> "" Then
        CurPath = Command & IIf(Right(Command, 1) = "\", Empty, "\") & "> "
    Else
        
    End If
    Mainfrm.Show
    Mainfrm.CurPathCom.Path = Left(CurPath, Len(CurPath) - 2)
End Sub
