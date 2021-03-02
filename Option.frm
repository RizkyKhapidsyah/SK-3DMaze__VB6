VERSION 5.00
Begin VB.Form Opton 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   2265
   ClientLeft      =   3390
   ClientTop       =   2010
   ClientWidth     =   3750
   ControlBox      =   0   'False
   Icon            =   "Option.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2265
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Rand 
      Caption         =   "Start in random spot in maze when finished"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save Changes"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
   Begin VB.FileListBox File2 
      Height          =   480
      Left            =   2040
      Pattern         =   "*.dmo"
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.FileListBox File1 
      Height          =   870
      Left            =   120
      Pattern         =   "*.map"
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Opton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TX
Dim TY
Private Sub Command1_Click()
'saves the data
Dim Sav As SavOp
MazFile = File1.List(File1.ListIndex)
Demfile = File2.List(File2.ListIndex)

If Trim(Text1.Text) <> "" Then
    Chk = Dir(App.Path & "\" & LCase(Text1.Text) & ".dmo")
Else
    Chk = "Not"
End If

If Chk = "" Then
    Demfile = Text1.Text & ".dmo"
    Open App.Path & "\" & Demfile For Output As #1
    Close #1
End If
    
File1.Refresh
File2.Refresh

Open App.Path & "\Maze.dat" For Random As #1 Len = Len(Sav)
    Get #1, 1, Sav
    If Trim(MazFile) <> "" Then Sav.MFile = MazFile
    If Trim(Demfile) <> "" Then Sav.DFile = Demfile
    If Rand.Value = 1 Then Sav.RndS = True
    If Rand.Value = 0 Then Sav.RndS = False
    Put #1, 1, Sav
Close #1
Opton.Hide
MazFrm.Enabled = True
MazFrm.Show
Command1.Enabled = False
Command2.Enabled = True
Text1.Text = ""
File1.ListIndex = -1
File2.ListIndex = -1
End Sub

Private Sub Command2_Click()
Dim Sav As SavOp
File1.Refresh
File2.Refresh
Opton.Hide
MazFrm.Enabled = True
MazFrm.Show
Command1.Enabled = False
Text1.Text = ""
File1.ListIndex = -1
File2.ListIndex = -1
Open App.Path & "\Maze.dat" For Random As #1 Len = Len(Sav)
    Get #1, 1, Sav
Close #1
If Sav.RndS = True Then Rand.Value = 1
End Sub


Private Sub File1_Click()
Command1.Enabled = True

End Sub


Private Sub File2_Click()
Command1.Enabled = True
If File2.ListIndex <> -1 Then Text1.Text = Left(File2.List(File2.ListIndex), Len(File2.List(File2.ListIndex)) - 4)
End Sub


Private Sub File2_KeyDown(KeyCode As Integer, Shift As Integer)
'see if the delete button was pushed
If KeyCode = 46 Then
    If File2.ListCount > 1 Then
        Kill App.Path & "\" & File2.List(File2.ListIndex)
        Text1.Text = ""
        File2.Refresh
        File2.ListIndex = File2.ListCount - 1
        Command2.Enabled = False
    Else
        MsgBox "Need some file to record into"
    End If
End If
End Sub


Private Sub Form_Load()
Dim Sav As SavOp
File1.Path = App.Path
File2.Path = App.Path
Open App.Path & "\Maze.dat" For Random As #1 Len = Len(Sav)
    Get #1, 1, Sav
Close #1
If Sav.RndS = True Then Rand.Value = 1
End Sub

Private Sub Rand_Click()
Command1.Enabled = True
End Sub

Private Sub Text1_Change()
If Text1.Text <> "" Then Command1.Enabled = True
End Sub


