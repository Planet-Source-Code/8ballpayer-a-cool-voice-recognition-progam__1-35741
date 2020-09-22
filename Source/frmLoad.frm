VERSION 5.00
Object = "{4E3D9D11-0C63-11D1-8BFB-0060081841DE}#1.0#0"; "Xlisten.dll"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Command-By-Word"
   ClientHeight    =   4905
   ClientLeft      =   9825
   ClientTop       =   7230
   ClientWidth     =   9825
   Icon            =   "frmLoad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   9825
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVELISTENPROJECTLibCtl.DirectSR DirectSR2 
      Height          =   255
      Left            =   1560
      OleObjectBlob   =   "frmLoad.frx":030A
      TabIndex        =   15
      Top             =   4560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Exit"
      Height          =   315
      Left            =   8400
      TabIndex        =   13
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Status"
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   240
      Width           =   9615
      Begin VB.Timer Timer2 
         Interval        =   1
         Left            =   4680
         Top             =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Welcome to Command-By-Word"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   9015
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   7550
      Left            =   7440
      Top             =   4440
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1695
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   5655
      Begin VB.CommandButton Command1 
         Caption         =   "Edit Word List"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         ToolTipText     =   "Eddit your Word/Phrase and Command list"
         Top             =   960
         Width           =   2295
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Add New Word/Phrase"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         ToolTipText     =   "Click here to add a new word/phrase and it's command."
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Disable Command-By-Word"
         Height          =   375
         Left            =   2760
         TabIndex        =   8
         ToolTipText     =   "Use this button enabled and disable the Command-By-Word program"
         Top             =   960
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Hide Command-By-Word"
         Height          =   375
         Left            =   2760
         TabIndex        =   7
         ToolTipText     =   "Hide the Command-By-Word program(it will still run)"
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.ListBox List3 
      Height          =   1230
      Left            =   3360
      TabIndex        =   3
      Top             =   3240
      Width           =   1455
   End
   Begin VB.ListBox List2 
      Height          =   1230
      Left            =   4920
      TabIndex        =   2
      Top             =   3240
      Width           =   4815
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   3135
   End
   Begin ACTIVELISTENPROJECTLibCtl.DirectSR DirectSR1 
      Height          =   255
      Left            =   2040
      OleObjectBlob   =   "frmLoad.frx":032E
      TabIndex        =   0
      Top             =   4560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "File\Message"
      Height          =   255
      Left            =   4920
      TabIndex        =   14
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Command"
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label wordph 
      Caption         =   "Word/Phrase"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2


Private Sub Command4_Click()
Form1.Hide
End Sub

Private Sub Command5_Click()
End
End Sub

Private Sub DirectSR1_PhraseFinish(ByVal flags As Long, ByVal beginhi As Long, ByVal beginlo As Long, ByVal endhi As Long, ByVal endlo As Long, ByVal Phrase As String, ByVal parsed As String, ByVal results As Long)
    Dim noth
    List1.ListIndex = -1

    
    For i = 0 To List1.ListCount
        If Phrase = "" Then
            List1.ListIndex = -1
            List2.ListIndex = -1
            List3.ListIndex = -1
            Exit Sub
        End If

If Phrase = List1.List(i) Then
            List1.ListIndex = i
            List2.ListIndex = i
            List3.ListIndex = i
            Label4.Caption = "Welcome to Command-By-Click: Command-By-Word is currently executing command: " + List3.List(i)
            Timer1.Enabled = True
            If List3.Text = "Open" Then
            noth = Shell(List2.List(i), vbNormalNoFocus)
            ElseIf List3.Text = "Delete" Then
            Kill List2.Text
            ElseIf List3.Text = "Message Prompt" Then
            MsgBox (List2.Text)
            End If
        End If
    Next i
End Sub

Private Sub Form_Load()

    Dim junk, windir$
    
    Label4.Caption = ""
        
    windir = Space(144)
    junk = getwindir(windir, 144)
    windir = Trim(windir)
    i = InStr(windir$, vbNullChar)
    windir$ = Mid$(windir$, 1, i - 1)
    
    words = windir$ & "\words.txt"
    dirs = windir$ & "\dirs.txt"
    descrip = windir$ & "\desc.txt"
    
    test = Dir(words)
    If test = "" Then
        Open words For Output As #1
        Close #1
    End If
    
    test = Dir(dirs)
    If test = "" Then
        Open dirs For Output As #1
        Close #1
    End If
    
    test = Dir(descrip)
    If test = "" Then
        Open descrip For Output As #1
        Close #1
    End If
    
    Call loadfiles
    Label4.Caption = "Welcome to Command-By-Word: Command-By-Word is currently Enabled"
           
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call savefiles
    Set Form1 = Nothing
    Set Form2 = Nothing
    Set Form3 = Nothing
    End
End Sub

Private Sub command1_click()
    Form2.Show
End Sub

Private Sub Command2_Click()
    Form3.Show
End Sub

Private Sub Command3_Click()
    Select Case Command3.Caption
    Case Is = "Disable Command-By-Word"
        DirectSR1.Deactivate
        Command3.Caption = "Enable Command-By-Word"
        Label4.Caption = "Welcome to Command-By-Word: Command-By-Word is currently Disabled"
    Case Is = "Enable Command-By-Word"
        DirectSR1.Activate
        Command3.Caption = "Disable Command-By-Word"
        Label4.Caption = "Welcome to Command-By-Word: Command-By-Word is currently Enabled"
    End Select
End Sub

Private Sub Timer1_Timer()
Label4.Caption = "Welcome to Command-By-Word: Command-By-Word is currently Enabled"
Timer1.Enabled = False
End Sub

