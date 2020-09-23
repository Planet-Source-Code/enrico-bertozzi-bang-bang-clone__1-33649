VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "Bang Bang Clone!"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7215
   Icon            =   "FrmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   368
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   481
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Back 
      AutoRedraw      =   -1  'True
      Height          =   5055
      Left            =   120
      ScaleHeight     =   333
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   461
      TabIndex        =   0
      Top             =   240
      Width           =   6975
      Begin VB.PictureBox pExplsrc 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   255
         Left            =   840
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   17
         Top             =   2520
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox pExplmsk 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   255
         Left            =   360
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   16
         Top             =   2520
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox pTree2msk 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   255
         Left            =   1800
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   15
         Top             =   2040
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox pTree2src 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   255
         Left            =   1320
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   14
         Top             =   2040
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox pTree1msk 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   255
         Left            =   840
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   13
         Top             =   2040
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox pTree1src 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   255
         Left            =   360
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   12
         Top             =   2040
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox picCopy 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   3840
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   11
         Top             =   2040
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox picMask 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   300
         Left            =   3360
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   10
         Top             =   2040
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.PictureBox picSrc 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   300
         Left            =   2880
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   9
         Top             =   2040
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start game"
         Height          =   375
         Left            =   5160
         TabIndex        =   2
         Top             =   4560
         Width           =   1335
      End
      Begin VB.CommandButton cmdRegenerate 
         Caption         =   "Regenerate"
         Default         =   -1  'True
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   4560
         Width           =   1335
      End
      Begin VB.PictureBox pWind 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2640
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   109
         TabIndex        =   6
         Top             =   4560
         Visible         =   0   'False
         Width           =   1695
         Begin VB.Line Line1 
            BorderWidth     =   2
            X1              =   57
            X2              =   57
            Y1              =   0
            Y2              =   24
         End
      End
      Begin VB.Label lWind 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Wind"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   4320
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lAct 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Press Space"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   5
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Image imgP2 
         Height          =   135
         Left            =   1680
         Top             =   120
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Image imgP1 
         Height          =   135
         Left            =   1320
         Top             =   120
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lsInfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmMain.frx":0442
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4935
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   6615
      End
   End
   Begin VB.Label lP2Score 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   5280
      Width           =   6975
   End
   Begin VB.Label lP1Score 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Zoom% = 100 'PLEASE LEAVE 100 IN THERE OR GAME FREEZE!

Private Sub cmdRegenerate_Click()
Dim I%, IP%, SP%, SPB%, SPP%, NS%, SST%, CTC!, TP&, A!, C!, brange%, ntx%, ntl%
On Error GoTo Restart
Restart:
StClear
Back.Cls
VarResetFade
TP = Back.Point(1, 1)
SP = Int(Rnd * 7 / 8 * Back.ScaleHeight) + (1 / 8 * Back.ScaleHeight) '<< start level
SPB = SP
SPP = SP
NS = Int(Rnd * 1 / 12 * Back.ScaleHeight) '<< 1/12 (fraction) = complexity, lower to increase, modify also the below one...
brange = Back.ScaleWidth / 2 - 60
P1pe = Rnd * brange + 60
P1ps = P1pe - 40
P2pe = Rnd * brange + Back.ScaleWidth / 2 + 40
P2ps = P2pe - 40
For I = 1 To Back.ScaleWidth
 If I = NS Then
  IP = I
ExBound:
  A = Rnd * 2 - 1 '<< COS max height (parameter #1)
  C = Rnd * 2 + 0.5 '<< COS max lenght (parameter #2)
  CTC = Zoom * A
  If (TerLvl(I - 1) - 2 * CTC > Me.ScaleHeight) Or (TerLvl(I - 1) - 2 * CTC < 3 / 16 * Me.ScaleHeight) Then GoTo ExBound '<< lower and higher mountain bound
  SST = I + Int(180 / C)
  Do
   If (I > P1ps And I < P1pe) Or (I > P2ps And I < P2pe) Then SPP = SP: GoTo BaseDraw '<< cannons positions
   SPP = SP
   SP = Int(SPB + Zoom * (A * Cos(C * ((I - IP) * 3.141592 / 180)))) - CTC
BaseDraw:
   Back.Line (I, SP + 1)-(I, Back.ScaleWidth), RGB(0, 170, 0) '<< terrain color, modify also the below one
   Back.Line (I - 1, SPP)-(I, SP), RGB(0, 0, 0)
   TerLvl(I) = SP
   I = I + 1
  Loop Until I = SST
  Back.Line (I, SP + 1)-(I, Back.ScaleWidth), RGB(0, 170, 0) '<< terrain color, modify also the below one
  Back.Line (I - 1, SP)-(I + 1, SP), RGB(0, 0, 0)
  NS = Int(Rnd * 1 / 12 * Back.ScaleHeight) + I '<< 1/12 (fraction) = complexity, lower to increase
  SPB = SP
  TerLvl(I) = SP
 Else
  Back.PSet (I, SP), RGB(0, 0, 0)
  Back.Line (I, SP + 1)-(I, Back.ScaleWidth), RGB(0, 170, 0) '<< terrain color
  TerLvl(I) = SP
 End If
Next I
ntn = Int(Rnd * 10) + 5
For I = 1 To ntn
 ntx = Int(Rnd * (Back.ScaleWidth - 48)) + 24
 If Int(Rnd * 2) = 0 Then DrwTranspSpriteBlt Back, ntx - 24, TerLvl(ntx) - 46, pTree1src, pTree1msk Else DrwTranspSpriteBlt Back, ntx - 24, TerLvl(ntx) - 46, pTree2src, pTree2msk
Next I
TPR = LongToR(Back.Point(2, 2))
TPG = LongToG(Back.Point(2, 2))
TPB = LongToB(Back.Point(2, 2))
lsInfo.ForeColor = RGB(256 - TPR, 256 - TPG, 256 - TPB)
AllZero = False
StartX(1) = P1pe - 7
StartX(2) = P2ps + 7
StartY(1) = TerLvl(P1pe - 1) - 32
StartY(2) = TerLvl(P2ps + 1) - 32
imgP1.Top = TerLvl(P1ps + 1) - 32: imgP1.Left = P1ps + 1
imgP2.Top = TerLvl(P2pe - 1) - 32: imgP2.Left = P2ps + 7
imgP1.Visible = True
imgP2.Visible = True
lAct.ForeColor = RGB(256 - TPR, 256 - TPG, 256 - TPB)
On Error GoTo 0
End Sub

Private Sub cmdStart_Click()
If AllZero Then MsgBox "Click 'Regenerate' first!", vbExclamation: Exit Sub
CanShoot = True
PlrShoot(1).Angle = Empty
PlrShoot(1).Power = Empty
PlrShoot(2).Angle = Empty
PlrShoot(2).Power = Empty
PlrShoots(1) = 0
PlrShoots(2) = 0
If Not InGame Then
 TurnOf = Int(Rnd * 2) + 1
Else
 If HasWon = 1 Then TurnOf = 2 Else TurnOf = 1
End If
HasWon = 0
InGame = True
ObjSet osPlaying
W = 0.07 * Rnd - 0.035
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Dim Incr!
If KeyAscii = 32 And CanShoot And InGame Then
 StClear
 lAct.Visible = False
ReValue:
 Randomize (Time)
 Incr = 0.014 * Rnd - 0.007
 If W + Incr < -0.035 Or W + Incr > 0.035 Then GoTo ReValue
 W = W + Incr
 pWind.Cls
 R(1) = 255
 G(1) = 192
 B(1) = 0
 R(2) = 192
 G(2) = 64
 B(2) = 0
 R(3) = 0
 G(3) = 128
 B(3) = 0
 R(4) = 0
 G(4) = 255
 B(4) = 0
 FStep(1) = 0
 FStep(2) = pWind.Width / 2
 FStep(3) = pWind.Width / 2
 FStep(4) = pWind.Width
 ObjFade pWind, blHorizontal
 Select Case W
  Case Is > 0
  For I = 1 To pWind.Width / 2
   pWind.Line (I, 0)-(I, pWind.Height), RGB(255, 255, 255)
  Next I
  For I = Int(pWind.Width / 2 + (pWind.Width / 2 * W) / 0.035) To pWind.Width
   pWind.Line (I, 0)-(I, pWind.Height), RGB(255, 255, 255)
  Next I
  Case Is < 0
  For I = 1 To pWind.Width / 2 - (pWind.Width / 2 * Abs(W)) / 0.035
   pWind.Line (I, 0)-(I, pWind.Height), RGB(255, 255, 255)
  Next I
  For I = Int(pWind.Width / 2) To pWind.Width
   pWind.Line (I, 0)-(I, pWind.Height), RGB(255, 255, 255)
  Next I
  Case Else
 End Select
 frmShootCfg.Show 1, Me
 Sleep 750
End If
End Sub

Private Sub Form_Load()
StClear
pTree1src.Picture = LoadResPicture(101, 0)
pTree2src.Picture = LoadResPicture(103, 0)
pTree1msk.Picture = LoadResPicture(104, 0)
pTree2msk.Picture = LoadResPicture(102, 0)
picSrc.Picture = LoadResPicture(105, 0)
picMask.Picture = LoadResPicture(106, 0)
pExplmsk.Picture = LoadResPicture(108, 0)
pExplsrc.Picture = LoadResPicture(109, 0)
Load frmShootCfg
AllZero = True
InGame = False
imgP1.Picture = LoadResPicture(104, 1)
imgP2.Picture = LoadResPicture(105, 1)
lsInfo.Caption = "Bang Bang Clone 32-bit. 2-player game in turns. Here you have to shoot at the other cannon by giving angle and power. The game generates mountains, hills and valleys as obstacles to your ball, just click 'Regenerate' to see. Click 'Start game' to play in the current terrain. Good luck!"
Me.Caption = Version
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Not StopGame Then Cancel = 1 Else Unload Me: Unload frmShootCfg
End Sub

Private Sub Form_Resize()
Dim TPR&, TPG&, TPB&
If Me.WindowState = vbMinimized Then Exit Sub
If InGame Then
 If PWinState = vbMaximized Then
  Me.WindowState = PWinState
  Exit Sub
 End If
 If Me.WindowState = vbMaximized Then Me.WindowState = vbNormal
 frmShootCfg.WindowState = vbNormal
 Me.Width = PrevW
 Me.Height = PrevH
 Exit Sub
End If
Back.Top = 15
Back.Left = 0
Back.Width = Me.ScaleWidth
Back.Height = Me.ScaleHeight - 30
lsInfo.Left = 5
lsInfo.Width = Back.Width - 10
lP1Score.Left = 8
lP1Score.Width = Me.ScaleWidth
lP2Score.Left = 0
lP2Score.Width = Me.ScaleWidth - 8
lP2Score.Top = Me.ScaleHeight - 15
cmdRegenerate.Top = Back.Height - 49
cmdStart.Top = Back.Height - 49
cmdRegenerate.Left = 24
cmdStart.Left = Back.Width - 121
AllZero = True
PrevH = Me.Height: PrevW = Me.Width
PWinState = Me.WindowState
lAct.Left = 0
lAct.Top = 10
lAct.Width = Back.ScaleWidth
pWind.Left = Back.ScaleWidth / 2 - pWind.Width / 2
pWind.Top = Back.ScaleHeight - 49
lWind.Left = pWind.Left
lWind.Top = Back.ScaleHeight - 23
Back.Cls
End Sub

Sub GameLoop()
Dim DT!, T!, Axt!, Ayt!, X!, Y!, Vxt!, Xt!, Vyt!, Yt!, M!, Ka!, Grav!, I%, E&
Dim ProceedDrawHole As Boolean, HB!
DoEvents
DT = 0.01
T = 0
Yt = Back.ScaleHeight - StartY(TurnOf)
Xt = StartX(TurnOf)
Back.AutoRedraw = False
E = BitBlt(picCopy.hDC, 0, 0, 16, 16, Back.hDC, Xt - 8, Back.ScaleHeight - Yt - 8, SRCCOPY)
E = BitBlt(Back.hDC, Xt - 8, Back.ScaleHeight - Yt - 8, 16, 16, picMask.hDC, 0, 0, SRCAND)
E = BitBlt(Back.hDC, Xt - 8, Back.ScaleHeight - Yt - 8, 16, 16, picSrc.hDC, 0, 0, SRCINVERT)
DoEvents
M = 10
Ka = 0.5
Grav = 9.81
Axt = -Ka / M * PlrShoot(TurnOf).StartVX
Ayt = -Ka / M * PlrShoot(TurnOf).StartVY - Grav
Vxt = PlrShoot(TurnOf).StartVX + W
Vyt = PlrShoot(TurnOf).StartVY
WPlaySound "fire.wav"
CycleRestart:
X = 1 / 2 * Axt * DT ^ 2 + Vxt + Xt
Y = 1 / 2 * Ayt * DT ^ 2 + Vyt + Yt
E = BitBlt(Back.hDC, Xt - 8, Back.ScaleHeight - Yt - 8, 16, 16, picCopy.hDC, 0, 0, SRCCOPY)
E = BitBlt(picCopy.hDC, 0, 0, 16, 16, Back.hDC, X - 8, Back.ScaleHeight - Y - 8, SRCCOPY)
E = BitBlt(Back.hDC, X - 8, Back.ScaleHeight - Y - 8, 16, 16, picMask.hDC, 0, 0, SRCAND)
E = BitBlt(Back.hDC, X - 8, Back.ScaleHeight - Y - 8, 16, 16, picSrc.hDC, 0, 0, SRCINVERT)
Vxt = Axt * DT + Vxt + W
Vyt = Ayt * DT + Vyt
T = T + DT
Xt = X
Yt = Y
Axt = -Ka / M * Vxt
Ayt = -Ka / M * Vyt - Grav
If X > Back.ScaleWidth - 8 Then GoTo Collided
If X < 8 Then GoTo Collided
For I = 4 To 12
 If (X - 8 + I > P2ps + 15 And X - 8 + I < P2pe - 15) Or (X - 8 + I > P1ps + 15 And X - 8 + I < P1pe - 15) Then
  If Back.ScaleHeight - Y + 8 >= TerLvl(X - 8 + I) Then
   If (X - 8 + I > P2ps + 15 And X - 8 + I < P2pe - 15) Then CollidTo = 2
   If (X - 8 + I > P1ps + 15 And X - 8 + I < P1pe - 15) Then CollidTo = 1
   GoTo Collided
  End If
 End If
Next I
For I = 1 To 16
 If Back.ScaleHeight - Y + 6 >= TerLvl(X - 8 + I) Then
  CollidTo = 0
  GoTo Collided
 End If
Next I
Sleep 10
GoTo CycleRestart
Collided:
If CollidTo = 2 And TurnOf = 1 Then HasWon = 1: WPlaySound "destroy.wav"
If CollidTo = 1 And TurnOf = 2 Then HasWon = 2: WPlaySound "destroy.wav"
If CollidTo = 2 And TurnOf = 2 Then HasWon = 1: WPlaySound "destroy.wav"
If CollidTo = 1 And TurnOf = 1 Then HasWon = 2: WPlaySound "destroy.wav"
If CollidTo = 0 Then HasWon = 0: WPlaySound "blnull.wav"
Sleep 100
E = BitBlt(Back.hDC, Xt - 8, Back.ScaleHeight - Yt - 8, 16, 16, picCopy.hDC, 0, 0, SRCCOPY)
Back.AutoRedraw = True
If HasWon = 0 Then GoTo NoWins
If CollidTo = 1 Then
 DrwTranspSpriteBlt Back, P1ps - 4, TerLvl(P1ps + 1) - 48, pExplsrc, pExplmsk
 DoEvents
 imgP1.Picture = LoadResPicture(107, 1)
Else
 DrwTranspSpriteBlt Back, P2ps - 4, TerLvl(P2ps + 1) - 48, pExplsrc, pExplmsk
 DoEvents
 imgP2.Picture = LoadResPicture(108, 1)
End If
PlrScore(HasWon) = PlrScore(HasWon) + 10 + IIf(10 - PlrShoots(HasWon) > 0, (10 - PlrShoots(HasWon)) * 8, 0)
lP1Score.Caption = "Player 1 score: " & PlrScore(1)
lP2Score.Caption = "Player 2 score: " & PlrScore(2)
DoEvents
Sleep 2500
MsgBox "Player " & HasWon & " wons!", vbInformation
CollidTo = 0
If MsgBox("Another match?", vbQuestion + vbYesNo) = vbYes Then
 CanShoot = False
 imgP1.Picture = LoadResPicture(104, 1)
 imgP2.Picture = LoadResPicture(105, 1)
 ObjSet osLvSelect
 cmdRegenerate_Click
 Exit Sub
End If
InGame = False
FrmMain.Back.Cls
StClear
VarResetFade
ObjSet osWelcome
AllZero = True
CanShoot = False
PlrScore(1) = 0
PlrScore(2) = 0
cmdRegenerate_Click
Exit Sub

NoWins:
ProceedDrawHole = True
For I = X - 4 To X + 16 + 4
 If (I > P1ps And I < P1pe) Or (I > P2ps And I < P2pe) Then ProceedDrawHole = False: Exit For
Next I
If X < 8 Or X > Back.ScaleWidth - 8 Then ProceedDrawHole = False
If ProceedDrawHole = True Then
 For I = X - 12 To X + 12
  If Not (I > P1ps And I < P1pe) Or (I > P2ps And I < P2pe) Then
   HB = TerLvl(I)
   TerLvl(I) = Int(TerLvl(I) + (36 - (((I - (X)) / 2) ^ 2)))
   Back.Line (I, HB)-(I, TerLvl(I)), RGB(134, 69, 0)
   Back.Line (I - 1, TerLvl(I - 1))-(I, TerLvl(I)), RGB(0, 0, 0)
  End If
 Next I
End If
PlrShoots(TurnOf) = PlrShoots(TurnOf) + 1
If TurnOf = 1 Then TurnOf = 2 Else TurnOf = 1
lAct.Visible = True
DoEvents
End Sub

Sub VarResetFade()
FStep(1) = 0
FStep(2) = Back.ScaleHeight
Randomize (Timer)
R(1) = Int(Rnd * 255)
G(1) = Int(Rnd * 255)
B(1) = Int(Rnd * 255)
R(2) = Int(Rnd * 255)
G(2) = Int(Rnd * 255)
B(2) = Int(Rnd * 255)
ObjFade Back, blVertical
End Sub

Sub ObjSet(Status As Byte)
Select Case Status
 Case osPlaying
  lsInfo.Visible = False
  cmdStart.Visible = False
  cmdRegenerate.Visible = False
  lP1Score.Caption = "Player 1 score: " & PlrScore(1)
  lP2Score.Caption = "Player 2 score: " & PlrScore(2)
  lAct.Visible = True
  pWind.Visible = True
  lWind.Visible = True
  pWind.Cls
 Case osLvSelect
  lsInfo.Caption = "Click Regenerate until you see a terrain you like and then click Start Game to start another game."
  lsInfo.Visible = True
  cmdRegenerate.Visible = True
  cmdStart.Visible = True
  lWind.Visible = False
  pWind.Visible = False
 Case osWelcome
  FrmMain.lsInfo.Visible = True
  FrmMain.cmdStart.Visible = True
  FrmMain.cmdRegenerate.Visible = True
  FrmMain.imgP1.Visible = False
  FrmMain.imgP2.Visible = False
  FrmMain.lsInfo.Caption = "Bang Bang Clone 32-bit. 2-player game in turns. Here you have to shoot at the other cannon by giving angle and power. The game generates mountains, hills and valleys as obstacles to your ball, just click 'Regenerate' to see. Click 'Start game' to play in the current terrain. Good luck!"
  FrmMain.lP1Score.Caption = ""
  FrmMain.lP2Score.Caption = ""
  FrmMain.pWind.Visible = False
  FrmMain.lWind.Visible = False
  imgP1.Picture = LoadResPicture(104, 1)
  imgP2.Picture = LoadResPicture(105, 1)
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmEnd.Show 1, Me
End Sub

