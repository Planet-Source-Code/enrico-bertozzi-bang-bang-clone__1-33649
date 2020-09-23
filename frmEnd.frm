VERSION 5.00
Begin VB.Form frmEnd 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Vote me!!!"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton Ext 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Thanks to all who may find their code used here!"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1320
      Width           =   4695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Please vote me at PSC, leave feedback and suggestions: how can I improve this game? Tell me if you liked this!"
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   1680
      Width           =   4695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Game, graphics && all the stuff: Enrico Bertozzi"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   4695
   End
   Begin VB.Label Label2 
      Caption         =   "32-bit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Bang Bang Clone!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   300
      Width           =   3855
   End
End
Attribute VB_Name = "frmEnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128
End Type

Private Sub Ext_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim WinVer As OSVERSIONINFO
WinVer.dwOSVersionInfoSize = Len(WinVer)
E& = GetVersionEx(WinVer)
If WinVer.dwMajorVersion = 5 Then
 ODG = GetWindowLong(Me.hWnd, -20)
 ODG = ODG Or &H80000
 SetWindowLong Me.hWnd, -20, ODG
 SetLayeredWindowAttributes Me.hWnd, 0, 160, &H2
End If
End Sub
