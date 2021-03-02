VERSION 5.00
Begin VB.Form MazFrm 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "3D Maze:"
   ClientHeight    =   5565
   ClientLeft      =   2805
   ClientTop       =   3975
   ClientWidth     =   5520
   Icon            =   "MazForm.frx":0000
   LinkTopic       =   "Form1"
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5565
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton UnPause 
      Caption         =   "Click To Continue"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   5160
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox DrwBrd 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H0000FF00&
      FillStyle       =   4  'Upward Diagonal
      ForeColor       =   &H80000008&
      Height          =   5000
      Left            =   240
      ScaleHeight     =   331
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   331
      TabIndex        =   0
      Top             =   360
      Width           =   5000
   End
   Begin VB.Label Tim 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   5055
   End
   Begin VB.Menu MMaze 
      Caption         =   "Maze"
      Begin VB.Menu StartOver 
         Caption         =   "Start Over"
      End
      Begin VB.Menu BTime 
         Caption         =   "Best Time"
      End
      Begin VB.Menu UPause 
         Caption         =   "Pause Game"
      End
      Begin VB.Menu OpenM 
         Caption         =   "Options"
      End
      Begin VB.Menu Deme 
         Caption         =   "Demo"
         Begin VB.Menu Recor 
            Caption         =   "Record"
         End
         Begin VB.Menu PLy 
            Caption         =   "Play"
         End
      End
      Begin VB.Menu Sizing 
         Caption         =   "Window Size"
         Begin VB.Menu SizL 
            Caption         =   "Large"
         End
         Begin VB.Menu SizM 
            Caption         =   "Medium"
         End
         Begin VB.Menu SizS 
            Caption         =   "Small"
         End
         Begin VB.Menu SizB 
            Caption         =   "Max"
         End
      End
      Begin VB.Menu mnuSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu InFu 
         Caption         =   "Info"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "MazFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Store As String ' stores the moves made for the demo
Dim Paus As Boolean ' Pause?
Sub KeyMove(TheKey)
' Move
TH = H
TW = W
Select Case TheKey
Case 37 'left
    D = D - 1
    If D = 0 Then D = 4
        Store = Store & 1
Case 38 'up
    If D = 1 Then TH = TH - 1
    If D = 2 Then TW = TW + 1
    If D = 3 Then TH = TH + 1
    If D = 4 Then TW = TW - 1
        Store = Store & 2
Case 39 'right
    D = D + 1
    If D = 5 Then D = 1
        Store = Store & 3
Case 40 'left
    If D = 1 Then TH = TH + 1
    If D = 2 Then TW = TW - 1
    If D = 3 Then TH = TH - 1
    If D = 4 Then TW = TW + 1
        Store = Store & 4
End Select
'check to see if you acn move there
'if you can sets it where you want to be
If Check(TH, TW) = True Then H = TH: W = TW
'draw the spot that you are at now
Call TMove
End Sub

Sub Size(HBig, Max As Boolean)
' sixze the form
Dim TheSiz As SavOp
MazFrm.Visible = False
DrwBrd.Width = HBig
DrwBrd.Height = HBig
If Max = False Then MazFrm.Width = HBig + HBig / 8
DrwBrd.Left = (MazFrm.Width - HBig) / 2
If Max = False Then MazFrm.Height = HBig + 1000 + HBig / 32
Tim.Width = HBig
Tim.Left = (MazFrm.Width - HBig) / 2
If Max = False Then
    Top = Screen.Height / 2 - Height / 2
    Left = Screen.Width / 2 - Width / 2
End If
'centers unpause button, still not visible
UnPause.Top = Height / 2 - UnPause.Height / 2
UnPause.Left = Width / 2 - UnPause.Width / 2
'saves the size of the window to a file
Open App.Path & "\Maze.dat" For Random As #1 Len = Len(TheSiz)
    Get #1, 1, TheSiz
    TheSiz.Max = Max
    TheSiz.SSize = HBig
    Put #1, 1, TheSiz
Close #1
DoEvents
Enabled = True
MazFrm.Visible = True
If MazeFile <> "" Then Call TMove
End Sub

Private Sub BTime_Click()
Dim BT As TheMap
PauseAd = Timer - Start + PauseAd
Paus = True
Open App.Path & "\" & MazeFile For Random As #1 Len = Len(BT)
Get #1, 1, BT
Close #1
M = Mid(BT.Map(0), 17, 2)
S = Mid(BT.Map(0), 19, 2)
If Val(M) + Val(S) = 0 Then MsgBox "Maze Has Never Been Completed!": Exit Sub
MsgBox "The Best Time For This Maze is " & M & ":" & S
End Sub

Private Sub Exit_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim GetSiz As SavOp
' Get the last size used from the maze.dat file
Open App.Path & "\Maze.dat" For Random As #1 Len = Len(GetSiz)
    Get #1, 1, GetSiz
Close #1
'check off the size used in the menu
If GetSiz.SSize = 5000 Then SizS.Checked = True
If GetSiz.SSize = 7500 Then SizM.Checked = True
If GetSiz.SSize = 10000 Then SizL.Checked = True
If Max = True Then SizB.Checked = True
'change the size
Call Size(GetSiz.SSize, GetSiz.Max)
'start rolling the game
Call StartOver_Click
End Sub


Private Sub Form_Resize()
Enabled = False
If WindowState = 2 Then
    Call SizB_Click
ElseIf WindowState <> 1 Then
    Call Size((8 * Width / 9), False)
    SizB.Checked = False
    If DrwBrd.Width = 5000 Then SizS.Checked = True
    If DrwBrd.Width = 7500 Then SizM.Checked = True
    If DrwBrd.Width = 10000 Then SizL.Checked = True
End If
Caption = FrmCap & MazeFile
End Sub


Private Sub Form_Unload(Cancel As Integer)
Unload Opton
End Sub

Private Sub InFu_Click()
Paus = True
Enabled = False
PauseAd = Timer - Start + adpause
MsgBox "This game was made by Ryan Condon, Email me your thoughts or if you have any ideas  KingBeeXC@hotmail.com"
Enabled = True
End Sub

Private Sub OpenM_Click()
' Get the options form
' pause the game
MazFrm.Enabled = False
Opton.Show
PauseAd = Timer - Start + PauseAd
Paus = True
End Sub

Private Sub DrwBrd_KeyDown(KeyCode As Integer, Shift As Integer)
'if the demo isn't running then send the key pressed to be processed
If Demo = False Then Call KeyMove(KeyCode)
End Sub

Private Sub DrwBrd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Button = 1 Then Caption = DrwBrd.Width / X & " " & (DrwBrd.Width / Y)
End Sub


Private Sub PLy_Click()
'play the demo
Demo = True
PLy.Enabled = False
Store = ""
'Call Timer1_Timer
Call StartOver_Click
End Sub

Public Sub Recor_Click()
'set it up so you record your moves
If Recor.Checked = False Then
    Recor.Checked = True
    PLy.Enabled = False
    Store = ""
    Call StartOver_Click
Else
    Recor.Checked = False
    PLy.Enabled = True
    Open App.Path & "\" & DemoFile For Output As #1
    Print #1, Store
    Close #1
End If
End Sub

Private Sub SizB_Click()
'change size of box
SizL.Checked = False
SizM.Checked = False
SizS.Checked = False
SizB.Checked = True
MazFrm.WindowState = 2
Enabled = False
Siz = -((32000 - 32 * MazFrm.Height) / 33)
Call Size(Siz, True)
End Sub

Private Sub SizL_Click()
'change size of box
MazFrm.WindowState = 0
SizL.Checked = True
SizM.Checked = False
SizS.Checked = False
SizB.Checked = False
Enabled = False
Call Size(10000, False)
End Sub

Private Sub SizM_Click()
'change size of box, maximize the form
MazFrm.WindowState = 0
SizL.Checked = False
SizM.Checked = True
SizS.Checked = False
SizB.Checked = False
Enabled = False
Call Size(7500, False)
End Sub




Private Sub SizS_Click()
'change size of box
MazFrm.WindowState = 0
SizL.Checked = False
SizM.Checked = False
SizS.Checked = True
SizB.Checked = False
Enabled = False
Call Size(5000, False)
End Sub


Private Sub StartOver_Click()
Dim Maze As TheMap
Dim Dat As SavOp
Timer1.Enabled = False
'get the samed data like the maze file and the demo file
If Dir(App.Path & "\Maze.dat") = "" Then MsgBox "File Missing 'Maze.dat'": Exit Sub
Open App.Path & "\Maze.dat" For Random As #1 Len = Len(Dat)
    Get #1, 1, Dat
Close #1
   MazeFile = Dat.MFile
   DemoFile = Dat.DFile
   MazeFile = UCase(Left(MazeFile, 1)) & LCase(Mid(MazeFile, 2))
   Caption = FrmCap & MazeFile
If Dir(App.Path & "\" & MazeFile) = "" Then
    MsgBox "File Not Found '" & App.Path & "\" & MazeFile
    Call OpenM_Click
    Exit Sub
End If
Open App.Path & "\" & MazeFile For Random As #1 Len = Len(Maze)
If LOF(1) / Len(Maze) = 0 Then: a = MsgBox("File Not Found!", vbCritical, "Error"): Exit Sub
Get #1, 1, Maze
Close #1
Store = ""
H = Mid(Maze.Map(0), 1, 2)
W = Mid(Maze.Map(0), 3, 2)
D = Mid(Maze.Map(0), 5, 1)
PauseAd = 0
If UnPause.Visible = True Then Call UnPause_Click
Start = Timer
Timer1.Enabled = True
Tim.Caption = ""
Tim.Visible = True
Call TMove
End Sub



Private Sub Timer1_Timer()
Static Ov As Boolean
Static Ove As Boolean
Static MN As Integer
Dim dem As String
Dim Dat As SavOp
'this checks to see if the program is still paused
If Paus = True Then
    If Paus = True And Enabled = True And UnPause.Visible = False Then
        'starts over if the map is changed
        Open App.Path & "\Maze.dat" For Random As #1 Len = Len(Dat)
            Get #1, 1, Dat
        Close #1
        Start = Timer
        Paus = False
        If LCase(Dat.MFile) <> LCase(MazeFile) Then Call StartOver_Click
    End If
    Exit Sub
End If
TheEnd = Timer
TotalTime = TheEnd - Start + PauseAd
'sets the time to minutes and secs
If TotalTime >= 60 Then
    M = Int(TotalTime / 60)
    S = TotalTime Mod 60
Else
    M = 0
    S = Int(TotalTime)
End If
Ov = Recor.Checked
'if recording flash it w/ the time
If Ov = True And Ove = True Then
    Tim.Caption = "Recording"
    Ove = False
Else ' Displays Time
    If S < 10 Then S = "0" & S
    If M < 10 Then M = "0" & M
    Tim.Caption = M & ":" & S
    If Ov = True Then Ove = True
End If
'if the form in minimized show it in the caption
If WindowState = 1 Then Caption = M & ":" & S
'if the demo is going read the file and tell where to move next
If Demo = False Then MN = 0: PLy.Enabled = True
If Demo = True And Paus = False Then
    Open App.Path & "\" & DemoFile For Input As #1
        If Not EOF(1) Then Input #1, dem
    Close #1
    MN = MN + 1
    If dem = "" Then a = MsgBox("Nothing Recorded In This File!", vbCritical, "Error"): Demo = False: Call StartOver_Click: Exit Sub
    If Len(dem) < MN And dem <> "" Then a = MsgBox("End Of Recording!", vbCritical, "Error"): Demo = False: Call StartOver_Click: Exit Sub
    a = Mid(dem, MN, 1)
    If a = 1 Then a = 37
    If a = 2 Then a = 38
    If a = 3 Then a = 39
    If a = 4 Then a = 40
    Call KeyMove(a)
End If
End Sub


Private Sub UnPause_Click()
UPause.Checked = False
UnPause.Visible = False
End Sub

Private Sub UPause_Click()
If UPause.Checked = True Then
    Call UnPause_Click
Else
    UPause.Checked = True
    UnPause.Visible = True
    PauseAd = Timer - Start + PauseAd
    Paus = True
End If
End Sub


