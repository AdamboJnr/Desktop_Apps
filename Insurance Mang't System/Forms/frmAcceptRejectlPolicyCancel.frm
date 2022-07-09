VERSION 5.00
Begin VB.Form frmAcceptRejectlPolicyCancel 
   Caption         =   "Accept/Reject Policy Cancel Claim"
   ClientHeight    =   5430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8625
   LinkTopic       =   "Form2"
   ScaleHeight     =   5430
   ScaleWidth      =   8625
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReject 
      Caption         =   "Reject"
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
      Left            =   7560
      Picture         =   "frmAcceptRejectlPolicyCancel.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Accept"
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
      Left            =   6360
      Picture         =   "frmAcceptRejectlPolicyCancel.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4440
      Width           =   855
   End
   Begin VB.Frame fraAcceptRejectPolicyCancelClaim 
      Height          =   3015
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   8295
      Begin VB.TextBox txtReasonToCancel 
         Height          =   375
         Left            =   6240
         TabIndex        =   11
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox txtPhoneNumber 
         Height          =   375
         Left            =   1920
         TabIndex        =   9
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox txtIdNumber 
         Height          =   375
         Left            =   6240
         TabIndex        =   7
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtName 
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   1440
         Width           =   1935
      End
      Begin VB.ListBox lstPolicyNumber 
         Height          =   840
         Left            =   1920
         TabIndex        =   3
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label lblReason 
         Alignment       =   1  'Right Justify
         Caption         =   "Reason To Cancel"
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
         Left            =   4200
         TabIndex        =   10
         Top             =   2160
         Width           =   1815
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
         Left            =   240
         TabIndex        =   8
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label IdNumber 
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
         Left            =   4200
         TabIndex        =   6
         Top             =   1440
         Width           =   1815
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
         Left            =   240
         TabIndex        =   4
         Top             =   1440
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
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Sum Assured"
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   7080
      Picture         =   "frmAcceptRejectlPolicyCancel.frx":0884
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblAcceptRejectPolicyCancel 
      Alignment       =   2  'Center
      Caption         =   "Accept/Reject Policy Cancel Claim"
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
      Width           =   6375
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   7335
   End
End
Attribute VB_Name = "frmAcceptRejectlPolicyCancel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
