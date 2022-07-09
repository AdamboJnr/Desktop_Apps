VERSION 5.00
Begin VB.Form frmLifeAssurancePolicyClaim 
   Caption         =   "Life Assurance Claim"
   ClientHeight    =   3915
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8250
   LinkTopic       =   "Form2"
   ScaleHeight     =   3915
   ScaleWidth      =   8250
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
      Left            =   6960
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
      Left            =   5640
      TabIndex        =   8
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Frame fraLifeAssurancePolicyClaimDetails 
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   7935
      Begin VB.TextBox txtReasonForClaim 
         Height          =   375
         Left            =   2160
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   960
         Width           =   1575
      End
      Begin VB.ComboBox cboAssuranceType 
         Height          =   315
         ItemData        =   "frmLifeAssurancePolicyClaim.frx":0000
         Left            =   6000
         List            =   "frmLifeAssurancePolicyClaim.frx":000A
         TabIndex        =   5
         Text            =   "Individual"
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtPolicyNumber 
         Height          =   375
         Left            =   2160
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblReasonForClaim 
         Alignment       =   1  'Right Justify
         Caption         =   "Reason For Claim"
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
         TabIndex        =   6
         Top             =   960
         Width           =   1695
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
         Height          =   375
         Left            =   4320
         TabIndex        =   4
         Top             =   240
         Width           =   1455
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
         Width           =   1695
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   6600
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblLifeAssurancePolicyClaim 
      Alignment       =   2  'Center
      Caption         =   "Life Assurance Policy Claim"
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
      TabIndex        =   0
      Top             =   360
      Width           =   5535
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   840
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   6495
   End
End
Attribute VB_Name = "frmLifeAssurancePolicyClaim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
