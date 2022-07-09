VERSION 5.00
Begin VB.Form frmRoomAllocation 
   Caption         =   "Room Allocation"
   ClientHeight    =   5715
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   5985
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
      Left            =   3480
      TabIndex        =   10
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdAllocate 
      Caption         =   "Allocate"
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
      Left            =   1920
      TabIndex        =   9
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox txtEmployeeName 
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox txtEmployeeNumber 
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   2880
      Width           =   1695
   End
   Begin VB.ComboBox cboBlock 
      Height          =   315
      Left            =   3000
      TabIndex        =   4
      Top             =   2160
      Width           =   1695
   End
   Begin VB.ComboBox cboRoomNumber 
      Height          =   315
      Left            =   3000
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Employee Name"
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
      Left            =   480
      TabIndex        =   7
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label lblEmployeeNumber 
      Alignment       =   1  'Right Justify
      Caption         =   "Employee number"
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
      Left            =   480
      TabIndex        =   5
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label lblBlock 
      Alignment       =   1  'Right Justify
      Caption         =   "Block"
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
      Left            =   600
      TabIndex        =   3
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label lblRoomNumber 
      Alignment       =   1  'Right Justify
      Caption         =   "Room number"
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
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblRoom 
      Alignment       =   2  'Center
      Caption         =   "Room Allocation"
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
      Width           =   2535
   End
End
Attribute VB_Name = "frmRoomAllocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
