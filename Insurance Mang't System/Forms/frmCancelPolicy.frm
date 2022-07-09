VERSION 5.00
Begin VB.Form frmCancelPolicy 
   Caption         =   "Cancel Policy"
   ClientHeight    =   5055
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8460
   LinkTopic       =   "Form2"
   ScaleHeight     =   5055
   ScaleWidth      =   8460
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
      TabIndex        =   13
      Top             =   3960
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
      TabIndex        =   12
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Frame fraCancelPolicyDetails 
      Height          =   2535
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   7815
      Begin VB.ComboBox cboReasonToCancel 
         Height          =   315
         ItemData        =   "frmCancelPolicy.frx":0000
         Left            =   3720
         List            =   "frmCancelPolicy.frx":0002
         TabIndex        =   11
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox txtPhoneNumber 
         Height          =   375
         Left            =   5760
         TabIndex        =   9
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtIdNumber 
         Height          =   375
         Left            =   2160
         TabIndex        =   7
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtName 
         Height          =   375
         Left            =   5760
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtPolicyNumber 
         Height          =   375
         Left            =   2160
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblReasonToCancel 
         Caption         =   "Reason for Cancelling"
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
         Left            =   1560
         TabIndex        =   10
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label lblPhoneNumber 
         Alignment       =   1  'Right Justify
         Caption         =   "Phone Number"
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
         Left            =   4080
         TabIndex        =   8
         Top             =   960
         Width           =   1335
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
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblName 
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
         Left            =   4080
         TabIndex        =   4
         Top             =   240
         Width           =   1095
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
         Width           =   1455
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6600
      Picture         =   "frmCancelPolicy.frx":0004
      Stretch         =   -1  'True
      Top             =   360
      Width           =   600
   End
   Begin VB.Label lblPolicyCancel 
      Alignment       =   2  'Center
      Caption         =   "Cancel Policy"
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
      Width           =   5415
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   960
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   6375
   End
End
Attribute VB_Name = "frmCancelPolicy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
