VERSION 5.00
Begin VB.Form Mainfrm 
   Caption         =   "Command"
   ClientHeight    =   4785
   ClientLeft      =   6000
   ClientTop       =   2010
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   7470
   Begin VB.TextBox Console 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "System"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4695
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Mainfrm.frx":0000
      Top             =   0
      Width           =   7455
   End
End
Attribute VB_Name = "Mainfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public CurPath As String
Public CurInput As String
Public History As String

Private Sub Console_Change()
        Dim Pref As Long
        Pref = Len(History) + Len(CurPath)
        If Pref < Len(Console.Text) Then
            CurInput = Right(Console.Text, Len(Console.Text) - Pref)
        Else
            CurInput = Empty
        End If
End Sub

Private Sub Console_KeyDown(KeyCode As Integer, Shift As Integer)
        
        If KeyCode = vbKeyBack Then
            Reset
        ElseIf KeyCode = vbKeyReturn Then
            History = Console.Text
            Console.Text = History & CurPath
            Console.SelStart = Len(History)
            History = History & vbCrLf
        End If
End Sub


Private Sub Reset(Optional Force As Boolean = False)
        If CurInput = Empty Or Force Then
            Console.Text = History & CurPath
            Console.SelStart = Len(Console.Text)
        End If
End Sub

Private Sub Console_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack Then
        Reset
    ElseIf KeyCode = vbKeyReturn Then
        Console.SelStart = Len(Console.Text)
        Reset True
    End If
End Sub

Private Sub Form_Load()
    Helper.Show
    CurPath = App.Path & "> "
    Console.Text = CurPath
    Console.SelStart = Len(CurPath)
End Sub

Private Sub Form_Resize()
    Console.Move 0, 0, ScaleWidth, ScaleHeight
End Sub
