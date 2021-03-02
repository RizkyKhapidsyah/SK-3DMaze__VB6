Attribute VB_Name = "MazeMod"
Global D As Integer 'Direction
Global H As Integer ' Height
Global W As Integer 'Width
Global Start ' for the timing
Global TheEnd 'for the timing
Global PauseAd ' For the timing
Global Demo As Boolean ' is the demo going?
Public Const FrmCap = "3D Maze: "
Public Const MH = 20 ' max lenght of maze
Public Const MW = 20 ' max width of maze
'Tells how to read the maze file
Type TheMap
    Map(MH) As String * MW
End Type
' the name of files
Global MazeFile As String
Global DemoFile As String
'Tells how to read the options menu
Type SavOp
    MFile As String * 13 'maze file name
    DFile As String * 13 'demo file name
    SSize As Integer ' the picture size
    Max As Boolean 'is the form maximized
    RndS As Boolean ' random start
End Type
'Sub to color in the maze
Declare Sub FloodFill Lib "GDI32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long)

Function Check(TH, TW)
'this function sees if where the player is tring to move
'is a legal spot
Dim Maze As TheMap
Open App.Path & "\" & MazeFile For Random As #1 Len = Len(Maze)
Get #1, 1, Maze
Close #1
' if the player can't move there exit set check to false
If TH > MH Or TW > MW Or TH <= 0 Or TW <= 0 Then Check = False: Exit Function
Temp = Mid(Maze.Map(TH), TW, 1)
If Temp >= 1 Then Check = True
If Temp = 0 Then Check = False
End Function


Sub Fill(Left, Right, Center, Level)
MazFrm.DrwBrd.FillStyle = 0
f = MazFrm.DrwBrd.ScaleHeight
backup = MazFrm.DrwBrd.FillColor
Rep = 255 / 6
Col = 255
Select Case Level
Case 0
   ' this Case colors the finish if its in the center
   ' don't ask me why i put it here, i know its confusing
    If Center <> -1 Then
        MazFrm.DrwBrd.FillColor = RGB(Rep * Center, Col, Rep * Center)
    Else
        MazFrm.DrwBrd.FillColor = RGB(Rep * 5.5, Col, Rep * 5.5)
    End If
    Call FloodFill(MazFrm.DrwBrd.hdc, f / 2, f / 2, 0)
    If Left = 2 And Center <> -1 Then
    MazFrm.DrwBrd.FillColor = RGB(Rep * Center, Rep * Center, Rep * Center)
        If Center = 1 Then
            MazFrm.DrwBrd.Line (f - f / 5, f - f / 5)-(f / 5, f / 5), , BF
            MazFrm.DrwBrd.Line (f - f / 5, f - f / 5)-(f / 5, f / 5), 0, B
        End If
        If Center = 2 Then
            MazFrm.DrwBrd.Line (f - f / 3, f - f / 3)-(f / 3, f / 3), , BF
            MazFrm.DrwBrd.Line (f - f / 3, f - f / 3)-(f / 3, f / 3), 0, B
        End If
        If Center = 3 Then
            MazFrm.DrwBrd.Line (f - f / 2.5, f - f / 2.5)-(f / 2.5, f / 2.5), , BF
            MazFrm.DrwBrd.Line (f - f / 2.5, f - f / 2.5)-(f / 2.5, f / 2.5), 0, B
        End If
        If Center = 4 Then
            MazFrm.DrwBrd.Line (f - f / 2.25, f - f / 2.25)-(f / 2.25, f / 2.25), , BF
            MazFrm.DrwBrd.Line (f - f / 2.25, f - f / 2.25)-(f / 2.25, f / 2.25), 0, B
        End If
        If Center = 5 Then
             MazFrm.DrwBrd.Line (f - f / 2.12, f - f / 2.12)-(f / 2.12, f / 2.12), , BF
             MazFrm.DrwBrd.Line (f - f / 2.12, f - f / 2.12)-(f / 2.12, f / 2.12), 0, B
        End If
    End If
   
Case 1 'first spot
    MazFrm.DrwBrd.FillColor = RGB(0, Col, 0)
    'colors sides
    Call FloodFill(MazFrm.DrwBrd.hdc, f - 1, f / 2, 0)
    Call FloodFill(MazFrm.DrwBrd.hdc, 0, f / 2, 0)
    MazFrm.DrwBrd.FillColor = RGB(255, 0, 0)
    'colors top
    Call FloodFill(MazFrm.DrwBrd.hdc, f / 2, f / 2 + f / 3, 0)
    Call FloodFill(MazFrm.DrwBrd.hdc, f / 2, f / 2 - f / 3, 0)
    If Right = 1 Then 'colors rights if hallway is there
        Call FloodFill(MazFrm.DrwBrd.hdc, f - 1, f - f / 18, 0)
        Call FloodFill(MazFrm.DrwBrd.hdc, f - 1, f / 18, 0)
    End If
    If Left = 1 Then
        Call FloodFill(MazFrm.DrwBrd.hdc, 0, f - f / 18, 0)
        Call FloodFill(MazFrm.DrwBrd.hdc, 0, f / 18, 0)
    End If
    If Left = 2 Then 'color if finish is on the left
        MazFrm.DrwBrd.FillColor = RGB(OCol, OCol, OCol)
        Call FloodFill(MazFrm.DrwBrd.hdc, 0, f / 2, 0)
    End If
    If Right = 2 Then
        MazFrm.DrwBrd.FillColor = RGB(OCol, OCol, OCol)
        Call FloodFill(MazFrm.DrwBrd.hdc, f - 1, f / 2, 0)
    End If
Case 2
    OCol = Rep * 2 ' this changes the color so it looks like it fades down the hall
    MazFrm.DrwBrd.FillColor = RGB(OCol, Col, OCol)
    Call FloodFill(MazFrm.DrwBrd.hdc, f / 2 + f / 4.5, f / 2, 0)
    Call FloodFill(MazFrm.DrwBrd.hdc, f / 2 - f / 4.5, f / 2, 0)
    MazFrm.DrwBrd.FillColor = RGB(Col, OCol, OCol)
    Call FloodFill(MazFrm.DrwBrd.hdc, f / 2, f / 2 + f / 4.5, 0)
    Call FloodFill(MazFrm.DrwBrd.hdc, f / 2, f / 2 - f / 4.5, 0)
    If Right = 1 Then
        Call FloodFill(MazFrm.DrwBrd.hdc, f / 1.3, f / 2 + f / 4.5, 0)
        Call FloodFill(MazFrm.DrwBrd.hdc, f / 1.3, f / 2 - f / 4.5, 0)
    End If
    If Left = 1 Then
        Call FloodFill(MazFrm.DrwBrd.hdc, f / 4, f / 2 + f / 4.5, 0)
        Call FloodFill(MazFrm.DrwBrd.hdc, f / 4, f / 2 - f / 4.5, 0)
    End If
    If Left = 2 Then
        MazFrm.DrwBrd.FillColor = RGB(OCol, OCol, OCol)
        Call FloodFill(MazFrm.DrwBrd.hdc, f / 2 - f / 4.5, f / 2, 0)
    End If
    If Right = 2 Then
        MazFrm.DrwBrd.FillColor = RGB(OCol, OCol, OCol)
        Call FloodFill(MazFrm.DrwBrd.hdc, f / 2 + f / 4.5, f / 2, 0)
    End If
Case 3
    OCol = Rep * 3
    MazFrm.DrwBrd.FillColor = RGB(OCol, Col, OCol)
    Call FloodFill(MazFrm.DrwBrd.hdc, f / 2 + f / 7.5, f / 2, 0)
    Call FloodFill(MazFrm.DrwBrd.hdc, f / 2 - f / 7.5, f / 2, 0)
    MazFrm.DrwBrd.FillColor = RGB(Col, OCol, OCol)
    Call FloodFill(MazFrm.DrwBrd.hdc, f / 2, f / 2 + f / 7.5, 0)
    Call FloodFill(MazFrm.DrwBrd.hdc, f / 2, f / 2 - f / 7.5, 0)
    If Right = 1 Then
        Call FloodFill(MazFrm.DrwBrd.hdc, f / 1.55, f / 2 + f / 7.5, 0)
        Call FloodFill(MazFrm.DrwBrd.hdc, f / 1.55, f / 2 - f / 7.5, 0)
    End If
    If Left = 1 Then
        Call FloodFill(MazFrm.DrwBrd.hdc, f / 2.9, f / 2 + f / 7.5, 0)
        Call FloodFill(MazFrm.DrwBrd.hdc, f / 2.9, f / 2 - f / 7.5, 0)
    End If
    If Left = 2 Then
        MazFrm.DrwBrd.FillColor = RGB(OCol, OCol, OCol)
        Call FloodFill(MazFrm.DrwBrd.hdc, f / 2 - f / 7.5, f / 2, 0)
    End If
    If Right = 2 Then
        MazFrm.DrwBrd.FillColor = RGB(OCol, OCol, OCol)
        Call FloodFill(MazFrm.DrwBrd.hdc, f / 2 + f / 7.5, f / 2, 0)
    End If
Case 4
    OCol = Rep * 4
    MazFrm.DrwBrd.FillColor = RGB(OCol, Col, OCol)
    Call FloodFill(MazFrm.DrwBrd.hdc, f / 2 + f / 13.5, f / 2, 0)
    Call FloodFill(MazFrm.DrwBrd.hdc, f / 2 - f / 13.5, f / 2, 0)
    MazFrm.DrwBrd.FillColor = RGB(Col, OCol, OCol)
    Call FloodFill(MazFrm.DrwBrd.hdc, f / 2, f / 2 + f / 13.5, 0)
    Call FloodFill(MazFrm.DrwBrd.hdc, f / 2, f / 2 - f / 13.5, 0)
    If Right = 1 Then
        Call FloodFill(MazFrm.DrwBrd.hdc, f / 1.7, f / 2 + f / 13.5, 0)
        Call FloodFill(MazFrm.DrwBrd.hdc, f / 1.7, f / 2 - f / 13.5, 0)
    End If
    If Left = 1 Then
        Call FloodFill(MazFrm.DrwBrd.hdc, f / 2.4, f / 2 + f / 13.5, 0)
        Call FloodFill(MazFrm.DrwBrd.hdc, f / 2.4, f / 2 - f / 13.5, 0)
    End If
    If Left = 2 Then
        MazFrm.DrwBrd.FillColor = RGB(OCol, OCol, OCol)
        Call FloodFill(MazFrm.DrwBrd.hdc, f / 2 - f / 13.5, f / 2, 0)
    End If
    If Right = 2 Then
        MazFrm.DrwBrd.FillColor = RGB(OCol, OCol, OCol)
        Call FloodFill(MazFrm.DrwBrd.hdc, f / 2 + f / 13.5, f / 2, 0)
    End If
Case 5
    OCol = Rep * 5
    MazFrm.DrwBrd.FillColor = RGB(OCol, Col, OCol)
    Call FloodFill(MazFrm.DrwBrd.hdc, f / 2 + f / 19.5, f / 2, 0)
    Call FloodFill(MazFrm.DrwBrd.hdc, f / 2 - f / 19.5, f / 2, 0)
    MazFrm.DrwBrd.FillColor = RGB(Col, OCol, OCol)
    Call FloodFill(MazFrm.DrwBrd.hdc, f / 2, f / 2 + f / 19.5, 0)
    Call FloodFill(MazFrm.DrwBrd.hdc, f / 2, f / 2 - f / 19.5, 0)
    If Right = 1 Then
        Call FloodFill(MazFrm.DrwBrd.hdc, f / 1.83, f / 2.15, 0)
        Call FloodFill(MazFrm.DrwBrd.hdc, f / 1.83, f / 1.86, 0)
    End If
    If Left = 1 Then
        Call FloodFill(MazFrm.DrwBrd.hdc, f / 2.22, f / 2.16, 0)
        Call FloodFill(MazFrm.DrwBrd.hdc, f / 2.22, f / 1.87, 0)
    End If
    If Left = 2 Then
        MazFrm.DrwBrd.FillColor = RGB(OCol, OCol, OCol)
        Call FloodFill(MazFrm.DrwBrd.hdc, f / 2 - f / 19.5, f / 2, 0)
    End If
    If Right = 2 Then
        MazFrm.DrwBrd.FillColor = RGB(OCol, OCol, OCol)
        Call FloodFill(MazFrm.DrwBrd.hdc, f / 2 + f / 19.5, f / 2, 0)
    End If
End Select
MazFrm.DrwBrd.FillStyle = 1
MazFrm.DrwBrd.FillColor = backup
End Sub

Sub Finish()
Dim Maze As TheMap
Dim Dat As SavOp
MazFrm.Tim.Visible = False
If Demo = True Then Demo = False: MazFrm.DrwBrd.Cls: Exit Sub
Open App.Path & "\" & MazeFile For Random As #1 Len = Len(Maze)
Get #1, 1, Maze
'gets the password for the level, if there is one
Pass = Mid(Maze.Map(0), 6)
If Mid(Pass, 1, 1) <> "|" Then Pass = Left(Pass, InStr(1, Pass, "|") - 1): Pas = True
TheEnd = Timer
TotalTime = TheEnd - Start
If TotalTime >= 60 Then
    M = Int(TotalTime / 60)
    S = TotalTime Mod 60
Else
    M = 0
    S = Int(TotalTime)
End If
' this Displays the time it took you to complete the game
MsgBox "It Took You " & M & " Minutes and " & S & " Seconds To Complete The Maze Congratulations"
'Gets The Best Time For The Maze From The File
'min
BM = Mid(Maze.Map(0), 17, 2) 'min
BS = Mid(Maze.Map(0), 19, 2) 'sec
'Converts it all to seconds
BT = Val(BM) * 60 + Val(BS)
'Converts The Time The Player Just Did, This Way It doesn't Deal w/ decimals
TT = Val(M) * 60 + Val(S)
'Compares to see if it beats the best time,
' if its zero that means no one has beaten it
If (BT > TT Or BT = 0) And M < 99 Then
    If BT = 0 Then
        MsgBox "You Are The First One To Beat This Maze, Nice Job!"
    Else
        MsgBox "Congratulations, You Have Completed This Maze In Record Time!!"
    End If
    'Writes the new best time in the file
    If S < 10 Then S = "0" & S
    If M < 10 Then M = "0" & M
    Maze.Map(0) = Left(Maze.Map(0), 16) & M & S
    Put 1, 1, Maze
End If
If BT = TT Then MsgBox "You Tied The Best Time!"
Close #1
   'Displays passwords
   If Pas = True Then MsgBox "The Password For This Maze Is '" & Pass & "'"
'if recording the maze, stop recording
If MazFrm.Recor.Checked = True Then Call MazFrm.Recor_Click
Demo = False
' If you want to you can have it choose a random space in the maze
' go to options, this is the code that finds the random space
Open App.Path & "\Maze.dat" For Random As #1 Len = Len(Dat)
    Get #1, 1, Dat
Close #1
If Dat.RndS = True Then
    Do
      TH = Int(Rnd * MH + 1)
      TW = Int(Rnd * MW + 1)
      D = Int(Rnd * 4 + 1)
    Loop Until Check(TH, TW) = True
    Start = Timer
    MazFrm.Tim.Visible = True
    If Check(TH, TW) = True Then H = TH: W = TW
    Call TMove
Else
    MazFrm.DrwBrd.Cls
End If
End Sub

Public Sub TMove()
Dim Maze As TheMap
Dim Cen As Boolean
Dim L(6) As Integer
Dim R(6) As Integer
Dim C(6) As Integer
' this code reads the mazefile then sets it up to draw
Open App.Path & "\" & MazeFile For Random As #1 Len = Len(Maze)
Get #1, 1, Maze
Close #1
Select Case D
'read it accorcding to the way the player is facing
    Case 1
        For i = H To H - 5 Step -1
        If i > MH Or i <= 0 Then Exit For
          S = S + 1
          If W + 1 > MW Then
            R(S) = 0
          Else
            R(S) = Mid(Maze.Map(i), W + 1, 1)
          End If
          If W - 1 <= 0 Then
                L(S) = 0
            Else
                L(S) = Mid(Maze.Map(i), W - 1, 1)
        End If
            C(S) = Mid(Maze.Map(i), W, 1)
        Next i
    Case 2
        For i = W To W + 5
        If i > MW Or i <= 0 Then Exit For
          S = S + 1
            If H - 1 <= 0 Then
                L(S) = 0
            Else
                L(S) = Mid(Maze.Map(H - 1), i, 1)
            End If
            If H + 1 > MH Then
                R(S) = 0
            Else
                R(S) = Mid(Maze.Map(H + 1), i, 1)
            End If
            C(S) = Mid(Maze.Map(H), i, 1)
        Next i
    Case 3
        For i = H To H + 5
        If i > MH Or i <= 0 Then Exit For
          S = S + 1
            If W + 1 > MW Then
                L(S) = 0
            Else
                L(S) = Mid(Maze.Map(i), W + 1, 1)
            End If
            If W - 1 <= 0 Then
                R(S) = 0
            Else
                R(S) = Mid(Maze.Map(i), W - 1, 1)
            End If
            C(S) = Mid(Maze.Map(i), W, 1)
        Next i
    Case 4
        For i = W To W - 5 Step -1
            If i > MW Or i <= 0 Then Exit For
             S = S + 1
              If H - 1 <= 0 Then
                R(S) = 0
            Else
                R(S) = Mid(Maze.Map(H - 1), i, 1)
            End If
            If H + 1 > MH Then
                L(S) = 0
            Else
                L(S) = Mid(Maze.Map(H + 1), i, 1)
            End If
                C(S) = Mid(Maze.Map(H), i, 1)
            Next i
End Select
'Draws the basic maze
f = MazFrm.DrwBrd.ScaleHeight
MazFrm.DrwBrd.Cls
MazFrm.DrwBrd.Line (f - f / 5, f - f / 5)-(f / 5, f / 5), , B
MazFrm.DrwBrd.Line (f - f / 3, f - f / 3)-(f / 3, f / 3), , B
MazFrm.DrwBrd.Line (f - f / 2.5, f - f / 2.5)-(f / 2.5, f / 2.5), , B
MazFrm.DrwBrd.Line (f - f / 2.25, f - f / 2.25)-(f / 2.25, f / 2.25), , B
MazFrm.DrwBrd.Line (f - f / 2.12, f - f / 2.12)-(f / 2.12, f / 2.12), , B
MazFrm.DrwBrd.Line (0, 0)-(f / 2.12, f / 2.12)
MazFrm.DrwBrd.Line (f, f)-(f - f / 2.12, f - f / 2.12)
MazFrm.DrwBrd.Line (0, f)-(f / 2.12, f - f / 2.12)
MazFrm.DrwBrd.Line (f, 0)-(f - f / 2.12, f / 2.12)
'if the player is in the finish box go to the finish sub
If C(1) = "2" Then
    MazFrm.DrwBrd.Cls
    Call Finish
Else 'if not draw the lines
For i = 6 To 1 Step -1
    Call Draw(Val(L(i)), Val(R(i)), Val(C(i)), i)
    Call Fill(Val(L(i)), Val(R(i)), Val(C(i)), i)
    If C(i) = 0 Or C(i) = 2 Or (C(i) = 3 And i <> 1) Then pant = i
Next i
    'color in everything
    Call Fill(C(pant), Val(C(pant)), pant - 1, 0)
End If
End Sub

Sub Draw(Left, Right, Center, Level)
f = MazFrm.DrwBrd.ScaleHeight
'draw the hallways to the left and right, if needed
' and blocks the center, if needed
Select Case Level
 Case 1
    If Left = 1 Then
        MazFrm.DrwBrd.Line (f - f / 5, f - f / 5)-(-10, f - f / 5)
        MazFrm.DrwBrd.Line (f - f / 5, f / 5)-(-10, f / 5)
    End If
    If Right = 1 Then
        MazFrm.DrwBrd.Line (f - f / 5, f - f / 5)-(f, f - f / 5)
        MazFrm.DrwBrd.Line (f - f / 5, f / 5)-(f, f / 5)
    End If
 Case 2
    If Left = 1 Then
        MazFrm.DrwBrd.Line (f - f / 3, f - f / 3)-(f / 5, f - f / 3)
        MazFrm.DrwBrd.Line (f - f / 3, f / 3)-(f / 5, f / 3)
    End If
    If Right = 1 Then
        MazFrm.DrwBrd.Line (f - f / 3, f - f / 3)-(f - f / 5, f - f / 3)
        MazFrm.DrwBrd.Line (f - f / 3, f / 3)-(f - f / 5, f / 3)
    End If
    If Center = 0 Or Center = 3 Then MazFrm.DrwBrd.Line (f - f / 5, f - f / 5)-(f / 5, f / 5), vbWhite, BF
    If Center = 0 Or Center = 3 Then MazFrm.DrwBrd.Line (f - f / 5, f - f / 5)-(f / 5, f / 5), , B
Case 3
    If Left = 1 Then
        MazFrm.DrwBrd.Line (f - f / 2.5, f - f / 2.5)-(f / 3, f - f / 2.5)
        MazFrm.DrwBrd.Line (f - f / 2.5, f / 2.5)-(f / 3, f / 2.5)
    End If
    If Right = 1 Then
        MazFrm.DrwBrd.Line (f - f / 2.5, f - f / 2.5)-(f - f / 3, f - f / 2.5)
        MazFrm.DrwBrd.Line (f - f / 2.5, f / 2.5)-(f - f / 3, f / 2.5)
    End If
    If Center = 0 Or Center = 3 Then MazFrm.DrwBrd.Line (f - f / 3, f - f / 3)-(f / 3, f / 3), vbWhite, BF
    If Center = 0 Or Center = 3 Then MazFrm.DrwBrd.Line (f - f / 3, f - f / 3)-(f / 3, f / 3), , B
Case 4
    If Left = 1 Then
        MazFrm.DrwBrd.Line (f - f / 2.25, f - f / 2.25)-(f / 2.5, f - f / 2.25)
        MazFrm.DrwBrd.Line (f - f / 2.25, f / 2.25)-(f / 2.5, f / 2.25)
    End If
    If Right = 1 Then
        MazFrm.DrwBrd.Line (f - f / 2.25, f - f / 2.25)-(f - f / 2.5, f - f / 2.25)
        MazFrm.DrwBrd.Line (f - f / 2.25, f / 2.25)-(f - f / 2.5, f / 2.25)
    End If
    If Center = 0 Or Center = 3 Then MazFrm.DrwBrd.Line (f - f / 2.5, f - f / 2.5)-(f / 2.5, f / 2.5), vbWhite, BF
    If Center = 0 Or Center = 3 Then MazFrm.DrwBrd.Line (f - f / 2.5, f - f / 2.5)-(f / 2.5, f / 2.5), , B
Case 5
    If Left = 1 Then
        MazFrm.DrwBrd.Line (f - f / 2.12, f - f / 2.12)-(f / 2.25, f - f / 2.12)
        MazFrm.DrwBrd.Line (f - f / 2.12, f / 2.12)-(f / 2.25, f / 2.12)
    End If
    If Right = 1 Then
        MazFrm.DrwBrd.Line (f - f / 2.12, f - f / 2.12)-(f - f / 2.25, f - f / 2.12)
        MazFrm.DrwBrd.Line (f - f / 2.12, f / 2.12)-(f - f / 2.25, f / 2.12)
    End If
    If Center = 0 Or Center = 3 Then MazFrm.DrwBrd.Line (f - f / 2.25, f - f / 2.25)-(f / 2.25, f / 2.25), vbWhite, BF
    If Center = 0 Or Center = 3 Then MazFrm.DrwBrd.Line (f - f / 2.25, f - f / 2.25)-(f / 2.25, f / 2.25), , B
End Select
End Sub


