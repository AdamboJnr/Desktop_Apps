VERSION 5.00
Begin VB.Form frmPolicyDetails 
   Caption         =   "Policy Details"
   ClientHeight    =   7440
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10740
   LinkTopic       =   "Form2"
   ScaleHeight     =   7440
   ScaleWidth      =   10740
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
      Height          =   615
      Left            =   9000
      TabIndex        =   23
      Top             =   6480
      Width           =   1455
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
      Height          =   615
      Left            =   7200
      TabIndex        =   22
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   240
      TabIndex        =   12
      Top             =   4200
      Width           =   10215
      Begin VB.ComboBox cboNextOfKinResidence 
         Height          =   315
         Left            =   1320
         TabIndex        =   21
         Top             =   1440
         Width           =   1575
      End
      Begin VB.ComboBox cboRelation 
         Height          =   315
         ItemData        =   "Form2.frx":0000
         Left            =   7800
         List            =   "Form2.frx":0002
         TabIndex        =   19
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtNextOfKinPhoneNumber 
         Height          =   375
         Left            =   4680
         TabIndex        =   17
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtNextOfKinName 
         Height          =   375
         Left            =   1320
         TabIndex        =   15
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblNextOfKinResidence 
         Alignment       =   1  'Right Justify
         Caption         =   "Residence"
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
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblRelation 
         Alignment       =   1  'Right Justify
         Caption         =   "Relation"
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
         Left            =   6480
         TabIndex        =   18
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblNextOfKinPhoneNumber 
         Alignment       =   1  'Right Justify
         Caption         =   "Mobile No"
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
         Left            =   3240
         TabIndex        =   16
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblNextOfKinName 
         Alignment       =   1  'Right Justify
         Caption         =   "Name"
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
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblNextOfKin 
         Caption         =   "Next Of Kin"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Frame fraPolicyDetails 
      Height          =   2535
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   10215
      Begin VB.TextBox txtAgentNumber 
         Height          =   375
         Left            =   6600
         TabIndex        =   27
         Text            =   "N/A"
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txtAgentName 
         Height          =   375
         Left            =   2880
         TabIndex        =   25
         Text            =   "N/A"
         Top             =   1800
         Width           =   2055
      End
      Begin VB.ComboBox cboPolicyType 
         Height          =   315
         ItemData        =   "Form2.frx":0004
         Left            =   8520
         List            =   "Form2.frx":0006
         TabIndex        =   11
         Top             =   960
         Width           =   1695
      End
      Begin VB.ComboBox cboResidence 
         Height          =   315
         ItemData        =   "Form2.frx":0008
         Left            =   4800
         List            =   "Form2.frx":000A
         TabIndex        =   9
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtMobileNo 
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtIdNumber 
         Height          =   375
         Left            =   6840
         TabIndex        =   5
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox txtPolicyName 
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label lblAgentNumber 
         Alignment       =   1  'Right Justify
         Caption         =   "Agent Number"
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
         Left            =   5040
         TabIndex        =   26
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lblAgentName 
         Alignment       =   1  'Right Justify
         Caption         =   "Agent Name"
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
         Left            =   1320
         TabIndex        =   24
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblPolicyType 
         Alignment       =   1  'Right Justify
         Caption         =   "Policy Type"
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
         Left            =   6720
         TabIndex        =   10
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblResidence 
         Alignment       =   1  'Right Justify
         Caption         =   "Residence"
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
         Left            =   3360
         TabIndex        =   8
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblMobileNumber 
         Alignment       =   1  'Right Justify
         Caption         =   "Mobile No"
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
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblIdNumber 
         Alignment       =   1  'Right Justify
         Caption         =   "Id Number"
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
         Left            =   4800
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblPolicyName 
         Alignment       =   1  'Right Justify
         Caption         =   "Name"
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
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   7680
      Picture         =   "Form2.frx":000C
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblPolicyHeader 
      Alignment       =   2  'Center
      Caption         =   "Create Policy"
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
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   6015
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   1440
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   6975
   End
End
Attribute VB_Name = "frmPolicyDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text2_Change()

End Sub
