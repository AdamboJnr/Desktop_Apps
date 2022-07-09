VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6615
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":000C
   ScaleHeight     =   4665
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      TabIndex        =   6
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdLogIn 
      Caption         =   "Log In"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   5
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox txtPassword 
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox txtUsername 
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   600
      Left            =   360
      Picture         =   "Form1.frx":48896
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   600
   End
   Begin VB.Shape Shape2 
      Height          =   735
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label lblUserName 
      Alignment       =   2  'Center
      Caption         =   "UserName"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   465
      Left            =   360
      Picture         =   "Form1.frx":49037
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   465
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label lblLogInCaption 
      Alignment       =   2  'Center
      Caption         =   "Account Log In"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Image3_Click()

End Sub

