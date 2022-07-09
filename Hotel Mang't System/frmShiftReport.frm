VERSION 5.00
Begin VB.Form frmShiftReport 
   Caption         =   "Shift Report"
   ClientHeight    =   5385
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   5235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdViewAll 
      Caption         =   "View All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3840
      TabIndex        =   5
      Top             =   3840
      Width           =   1095
   End
   Begin VB.ListBox lstDetails 
      Height          =   840
      Left            =   2640
      TabIndex        =   4
      Top             =   2520
      Width           =   2055
   End
   Begin VB.ComboBox cboEmployeeNumber 
      Height          =   315
      Left            =   2760
      TabIndex        =   3
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   1680
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label lblDetails 
      Caption         =   "Details"
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
      Left            =   360
      TabIndex        =   2
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label lblEmployeeNumber 
      Caption         =   "Employee Number"
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
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label lblShiftReport 
      Alignment       =   2  'Center
      Caption         =   "Shift Report"
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
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "frmShiftReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label2_Click()

End Sub
