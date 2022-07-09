VERSION 5.00
Begin VB.Form frmLifeAssuranceForm 
   Caption         =   "Life Assurance"
   ClientHeight    =   4545
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4545
   ScaleWidth      =   9000
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
      Left            =   7680
      TabIndex        =   7
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdProceed 
      Caption         =   "Proceed"
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
      Left            =   6240
      TabIndex        =   6
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Frame fraLifeAssuranceDetails 
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   8655
      Begin VB.TextBox txtPremium 
         Height          =   375
         Left            =   6240
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox cboAssuranceType 
         Height          =   315
         ItemData        =   "frmLifeAssuranceForm.frx":0000
         Left            =   1920
         List            =   "frmLifeAssuranceForm.frx":0002
         TabIndex        =   3
         Text            =   "Individual Life Assurance"
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblPremium 
         Alignment       =   1  'Right Justify
         Caption         =   "Premium"
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
         Left            =   4680
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblAssuranceType 
         Alignment       =   1  'Right Justify
         Caption         =   "Assurance Type"
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
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   6840
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblLifeAssurance 
      Alignment       =   2  'Center
      Caption         =   "Life Assurance Policy"
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
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   5655
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   960
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   6615
   End
End
Attribute VB_Name = "frmLifeAssuranceForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
