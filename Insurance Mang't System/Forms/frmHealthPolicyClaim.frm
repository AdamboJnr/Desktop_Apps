VERSION 5.00
Begin VB.Form frmHealthPolicyClaim 
   Caption         =   "Health Policy Claim"
   ClientHeight    =   4065
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7410
   LinkTopic       =   "Form2"
   ScaleHeight     =   4065
   ScaleWidth      =   7410
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
      Left            =   6120
      TabIndex        =   9
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
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
      Left            =   4800
      TabIndex        =   8
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Frame fraHealthPolicyClaimDetails 
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   7095
      Begin VB.TextBox txtReasonForClaim 
         Height          =   405
         Left            =   2040
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   960
         Width           =   1455
      End
      Begin VB.ComboBox cboPolicyType 
         Height          =   315
         ItemData        =   "frmHealthPolicyClaim.frx":0000
         Left            =   5160
         List            =   "frmHealthPolicyClaim.frx":000A
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtPolicyNumber 
         Height          =   375
         Left            =   2040
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblReasonForClaim 
         Alignment       =   1  'Right Justify
         Caption         =   "Reason For Claim "
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
         Top             =   960
         Width           =   1695
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
         Left            =   3720
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblPolicyNumber 
         Alignment       =   1  'Right Justify
         Caption         =   "Policy Number"
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
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   6000
      Picture         =   "frmHealthPolicyClaim.frx":001F
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblHealthInsuranceClaim 
      Alignment       =   2  'Center
      Caption         =   "Health Policy Claim"
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
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   5175
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "frmHealthPolicyClaim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
