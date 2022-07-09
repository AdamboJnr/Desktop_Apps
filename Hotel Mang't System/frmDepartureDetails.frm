VERSION 5.00
Begin VB.Form frmDepartureDetails 
   Caption         =   "Departure Details"
   ClientHeight    =   4305
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   Picture         =   "frmDepartureDetails.frx":0000
   ScaleHeight     =   4305
   ScaleWidth      =   6690
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   8
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   7
      Top             =   3120
      Width           =   1095
   End
   Begin VB.ComboBox cboRoomNumber 
      Height          =   315
      Left            =   2880
      TabIndex        =   6
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox txtCustomerName 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   2880
      TabIndex        =   4
      Top             =   1800
      Width           =   2175
   End
   Begin VB.ComboBox cboCustomerNumber 
      Height          =   315
      Left            =   2880
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   3720
      Picture         =   "frmDepartureDetails.frx":1082
      Stretch         =   -1  'True
      Top             =   240
      Width           =   615
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label llblRoomNumber 
      Alignment       =   1  'Right Justify
      Caption         =   "Room Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label lblCustomerName 
      Alignment       =   1  'Right Justify
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label lblCustomerNumber 
      Alignment       =   1  'Right Justify
      Caption         =   "Customer Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label lblDepartureDetails 
      Alignment       =   2  'Center
      Caption         =   "Departure Details"
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
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "frmDepartureDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()

End Sub

Private Sub cmdSave_Click()

End Sub

Private Sub Form_Load()

End Sub
