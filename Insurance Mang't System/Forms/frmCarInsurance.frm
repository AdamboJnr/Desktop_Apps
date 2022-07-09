VERSION 5.00
Begin VB.Form frmCarInsurance 
   Caption         =   "Car Insurance"
   ClientHeight    =   5640
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   8250
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   7695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2055
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   3735
   End
   Begin VB.Image Image2 
      Height          =   1935
      Left            =   4320
      Top             =   1320
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   6000
      Picture         =   "frmCarInsurance.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Car Insurance"
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
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   5415
   End
End
Attribute VB_Name = "frmCarInsurance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
