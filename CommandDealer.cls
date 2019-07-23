VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommandDealer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Commands()

Private Sub Class_Initialize()
    Commands = Array("version", "help", "ls", "touch", "cd")
End Sub

Public Function GetCommands(Optional ParamCount As Integer = -1) As String
    Dim Output As String
    Dim Index As Integer
    For Index = 0 To UBound(Commands)
        Output = Output & Commands(Index) & vbCrLf
    Next Index
    GetCommands = Output
End Function

Public Function GetVersionDescription() As String
    GetVersionDescription = App.EXEName & " [ Version " & App.Major & "." & App.Minor & "." & App.Revision & " ]" & vbCrLf & App.LegalCopyright & vbCrLf
End Function

Public Function Deal() As String
    Dim Cmds
    Dim CurrentPath As String

    If CurInput = Empty Then Exit Function
    Cmds = Split(CurInput, " ")
    Cmds(0) = LCase(Cmds(0))
    Dim Index As Integer
    For Index = 0 To UBound(Commands)
        If Commands(Index) = Cmds(0) Then Exit For
    Next Index
    If Index > UBound(Commands) Then
        Deal = "Can't Find Command """ & Cmds(0) & """" & vbCrLf
    End If
    
    CurrentPath = Left(CurPath, Len(CurPath) - 2)
    
    Select Case Cmds(0)
    Case Commands(0)
        Deal = GetVersionDescription()
    Case Commands(1)
        Deal = GetCommands()
        
    Case Commands(2)
        Mainfrm.CurPathCom.Refresh
        For Index = 0 To Mainfrm.CurPathCom.ListCount - 1
            Deal = Deal & Mainfrm.CurPathCom.List(Index) & vbCrLf
        Next Index
    Case Commands(3)
        If UBound(Cmds) < 1 Then
            Deal = "Error: No File Name" & vbCrLf
        Else
            Open CurrentPath & Cmds(1) For Output As #1
            Close
        End If
        
    Case Commands(4)
        If UBound(Cmds) < 1 Then
            Deal = "Error: No Path" & vbCrLf
        Else
            If Cmds(1) = ".." Then
                CurPath = LastDirectory()
            ElseIf Cmds(1) = "." Then
                
            Else
                If Dir(Cmds(1), vbDirectory) = "" Then
                    Deal = "Error: Path """ & Cmds(1) & """ does not exists" & vbCrLf
                Else
                    CurPath = Cmds(1) & IIf(Right(Cmds(1), 1) = "\", Empty, "\") & "> "
                End If
            End If
        End If
        If Left(Deal, 5) <> "Error" Then Mainfrm.CurPathCom.Path = Left(CurPath, Len(CurPath) - 2)
    End Select
    Deal = vbCrLf & Deal
    
End Function
