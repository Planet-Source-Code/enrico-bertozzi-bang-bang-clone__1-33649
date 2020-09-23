VERSION 5.00
Begin VB.Form frmShootCfg 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Shoot parameters - player "
   ClientHeight    =   990
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4770
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   66
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   318
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStopGame 
      Caption         =   "End game"
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdFire 
      Caption         =   "Fire!"
      Default         =   -1  'True
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox txtPwr 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   840
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "100"
      Top             =   555
      Width           =   615
   End
   Begin VB.TextBox txtDeg 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1080
      MaxLength       =   2
      TabIndex        =   0
      Text            =   "45"
      Top             =   195
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Power:                    (1 -> 200)"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Angle:                          degrees (0 -> 90)"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "frmShootCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFire_Click()
On Error GoTo ExitSub
If CInt(txtDeg.Text) > 90 Or CInt(txtDeg.Text) < 0 Then Exit Sub
If CInt(txtPwr.Text) <= 0 Or CInt(txtPwr.Text) > 200 Then Exit Sub
PlrShoot(TurnOf).StartVX = IIf(TurnOf = 1, 1, -1) * Cos(CInt(txtDeg.Text) * 3.141 / 180) * CInt(txtPwr.Text) / 10
PlrShoot(TurnOf).StartVY = Sin(CInt(txtDeg.Text) * 3.141 / 180) * CInt(txtPwr.Text) / 10
PlrShoot(TurnOf).Power = CInt(txtPwr.Text)
PlrShoot(TurnOf).Angle = CInt(txtDeg.Text)
Me.Hide
FrmMain.GameLoop
ExitSub:
End Sub

Private Sub cmdStopGame_Click()
If StopGame Then
 InGame = False
 StClear
 AllZero = True
 CanShoot = False
 PlrScore(1) = 0
 PlrScore(2) = 0
 FrmMain.Back.Cls
 FrmMain.ObjSet osWelcome
 Unload Me
End If
End Sub

Private Sub Form_Activate()
Me.Caption = "Shoot parameters - player " & TurnOf
txtPwr.Text = PlrShoot(TurnOf).Power
txtDeg.Text = PlrShoot(TurnOf).Angle
txtDeg.SetFocus
If TurnOf = 1 Then
 Me.Left = FrmMain.Left + 600
 Me.Top = FrmMain.Top + 600
Else
 Me.Left = FrmMain.Left + FrmMain.Width - Me.Width - 600
 Me.Top = FrmMain.Top + 600
End If
End Sub

Private Sub txtDeg_GotFocus()
txtDeg.SelStart = 0
txtDeg.SelLength = Len(txtDeg.Text)
End Sub

Private Sub txtPwr_GotFocus()
txtPwr.SelStart = 0
txtPwr.SelLength = Len(txtPwr.Text)
End Sub
