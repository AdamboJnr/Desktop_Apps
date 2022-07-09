VERSION 5.00
Begin VB.Form frmCarPolicyClaim 
   Caption         =   "Car Policy Claim"
   ClientHeight    =   4635
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7860
   LinkTopic       =   "Form2"
   ScaleHeight     =   4635
   ScaleWidth      =   7860
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
      Left            =   6600
      TabIndex        =   11
      Top             =   3360
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
      Left            =   5160
      Picture         =   "frmCarPolicyClaim.frx":0000
      TabIndex        =   10
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Frame fraCarPolicyClaimDetails 
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   7575
      Begin VB.TextBox txtStreet 
         Height          =   375
         Left            =   5520
         TabIndex        =   9
         Text            =   "N/A"
         Top             =   960
         Width           =   1575
      End
      Begin VB.ComboBox cboReason 
         Height          =   315
         Left            =   5520
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtDateOccured 
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtPolicyNumber 
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblStreet 
         Alignment       =   1  'Right Justify
         Caption         =   "Street"
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
         TabIndex        =   8
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblReason 
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
         Left            =   3720
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblDateOccured 
         Alignment       =   1  'Right Justify
         Caption         =   "Date Occured"
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
         TabIndex        =   4
         Top             =   960
         Width           =   1335
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
         Width           =   1335
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   6120
      Picture         =   "frmCarPolicyClaim.frx":1082
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblCarPolicyClaim 
      Alignment       =   2  'Center
      Caption         =   "Car Policy Claim"
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
      Width           =   5055
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   840
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   6015
   End
End
Attribute VB_Name = "frmCarPolicyClaim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
