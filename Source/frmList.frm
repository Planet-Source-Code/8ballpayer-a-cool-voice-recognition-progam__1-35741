VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Words/Phrases List"
   ClientHeight    =   2745
   ClientLeft      =   3015
   ClientTop       =   2730
   ClientWidth     =   8745
   Icon            =   "frmList.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   8745
   Begin VB.CommandButton Command2 
      Caption         =   "Remove"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.ListBox List3 
      Height          =   1815
      Left            =   4080
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Filename of executable. Double click to see full directory"
      Top             =   240
      Width           =   4455
   End
   Begin VB.ListBox List2 
      Height          =   1815
      Left            =   2760
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "The description of the executable"
      Top             =   240
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "The sound recognised"
      Top             =   240
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Command"
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Label3"
      Height          =   255
      Left            =   3600
      TabIndex        =   6
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   255
      Left            =   -360
      TabIndex        =   3
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SendText As String
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2




Private Sub Command6_Click()

End Sub

Private Sub Form_Load()
    
    Label1.Caption = "Word/Phrase"
    Label3.Caption = "File/Message"
    Command1.Caption = "Close"
    Command2.Caption = "Delete"
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

Private Sub List1_Click()
    List2.ListIndex = List1.ListIndex
    List3.ListIndex = List1.ListIndex
End Sub

Private Sub List2_Click()
    On Error GoTo exitsub
    List1.ListIndex = List2.ListIndex
    List3.ListIndex = List2.ListIndex
exitsub:
End Sub

Private Sub command1_click()
    Form2.Hide
End Sub

Private Sub Command2_Click()
    Dim index As Integer
    index = List3.ListIndex
    On Error Resume Next
    List1.RemoveItem List1.ListIndex
    List2.RemoveItem List2.ListIndex
    List3.RemoveItem List3.ListIndex
    Form1.List1.RemoveItem (index)
    Form1.List2.RemoveItem (index)
    Form1.List3.RemoveItem (index)
    
    Call savefiles
    Form1.List1.Clear
    Form1.List2.Clear
    Form1.List3.Clear
    Form2.List1.Clear
    Form2.List2.Clear
    Form2.List3.Clear
    Call loadfiles
    
End Sub

Private Sub List1_DblClick()
    MsgBox Form1.List2.List(List3.ListIndex)
End Sub

Private Sub List2_DblClick()
    MsgBox Form1.List2.List(List3.ListIndex)
End Sub

Private Sub List3_Click()
On Error GoTo exitsub
List1.ListIndex = List3.ListIndex

List2.ListIndex = List3.ListIndex
exitsub:
End Sub

Private Sub List3_DblClick()
    MsgBox Form1.List2.List(List3.ListIndex)
End Sub
