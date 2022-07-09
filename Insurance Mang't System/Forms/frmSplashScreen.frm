VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSplashScreen 
   Caption         =   "Form1"
   ClientHeight    =   6870
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   800
      Left            =   120
      Top             =   600
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   6360
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   873
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label lblProgressBarCaption 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   5640
      Width           =   7935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "GEMINIA INSURANCE MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   5415
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7935
   End
   Begin VB.Image Image1 
      Height          =   4575
      Left            =   0
      Picture         =   "frmSplashScreen.frx":0000
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   7935
   End
End
Attribute VB_Name = "frmSplashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Timer1.Enabled = True
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2
End Sub
Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 5
If ProgressBar1.Value <= 40 Then
    lblProgressBarCaption.Caption = "WELCOME TO GEMINIA INSURANCE MANAGEMENT SYSTEM....KINDLY WAIT AS THE SYSTEM LOADS"
ElseIf ProgressBar1.Value >= 41 And ProgressBar1.Value <= 59 Then
    lblProgressBarCaption.Caption = "STARTING THE DATABASE SERVER....ALMOST THERE!"
ElseIf ProgressBar1.Value >= 60 And ProgressBar1.Value <= 84 Then
    lblProgressBarCaption.Caption = "UNPACKING APPROPRIATE FILES FOR YOU...!"
ElseIf ProgressBar1.Value >= 85 Then
    lblProgressBarCaption.Caption = "PREPARING ACCESS INTO THE SYSTEM...!"
End If
'lblProgressBarCaption.Caption = ProgressBar1.Value & "%"
If ProgressBar1.Value = ProgressBar1.Max Then
Unload Me
frmLogIn.Show
End If
End Sub
