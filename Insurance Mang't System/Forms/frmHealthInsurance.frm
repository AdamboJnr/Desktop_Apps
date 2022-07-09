VERSION 5.00
Begin VB.Form frmHealthInsurance 
   Caption         =   "Health Insurance"
   ClientHeight    =   8160
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11295
   LinkTopic       =   "Form2"
   ScaleHeight     =   8160
   ScaleWidth      =   11295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
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
      Left            =   8280
      TabIndex        =   18
      Top             =   6840
      Width           =   1335
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
      Height          =   615
      Left            =   6360
      TabIndex        =   17
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Frame fraHealthInsurancePolicy 
      Enabled         =   0   'False
      Height          =   2775
      Left            =   480
      TabIndex        =   3
      Top             =   3840
      Width           =   9255
      Begin VB.Frame fraDependantsBirthDetails 
         Height          =   735
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   8055
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   6000
            TabIndex        =   15
            Text            =   "1990"
            Top             =   240
            Width           =   1095
         End
         Begin VB.ComboBox cboDependantBirthMonth 
            Height          =   315
            Left            =   3360
            TabIndex        =   13
            Text            =   "January"
            Top             =   240
            Width           =   1095
         End
         Begin VB.ComboBox cboDependantBirthDay 
            Height          =   315
            Left            =   840
            TabIndex        =   11
            Text            =   "1"
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblDependantBirthYear 
            Alignment       =   1  'Right Justify
            Caption         =   "Year"
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
            Left            =   4800
            TabIndex        =   14
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lblDependantBirthMonth 
            Alignment       =   1  'Right Justify
            Caption         =   "Month"
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
            Left            =   2160
            TabIndex        =   12
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblDependantBirthDay 
            Caption         =   "Day"
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
            TabIndex        =   10
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   4920
         TabIndex        =   8
         Text            =   "Spouse"
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtDependantName 
         Height          =   375
         Left            =   1200
         TabIndex        =   6
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lblHealthInsuranceRealPrice 
         Caption         =   "Price: 8000/= Monthly"
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
         Left            =   5880
         TabIndex        =   16
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label lblDependantRelation 
         Alignment       =   1  'Right Justify
         Caption         =   "Relationship"
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
         TabIndex        =   7
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblDependantName 
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
         TabIndex        =   5
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblDependantsDetails 
         Caption         =   "Dependant Details"
         Enabled         =   0   'False
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
         TabIndex        =   4
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.ComboBox cboTypeOfInsurance 
      Height          =   315
      Left            =   2280
      TabIndex        =   2
      Text            =   "Self"
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   2175
      Left            =   5040
      Top             =   1320
      Width           =   4695
   End
   Begin VB.Label lblHealthInurancePrice 
      Caption         =   "Price: 6000/= Monthly"
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
      Left            =   2040
      TabIndex        =   19
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label lblTypeOfInsurance 
      Alignment       =   2  'Center
      Caption         =   "Insured Type"
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
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   7320
      Picture         =   "frmHealthInsurance.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   600
   End
   Begin VB.Label lblHealthInsuranceCaption 
      Alignment       =   2  'Center
      Caption         =   "Health Insurance Policy"
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
      Left            =   2880
      TabIndex        =   0
      Top             =   360
      Width           =   4215
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   2760
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   5295
   End
End
Attribute VB_Name = "frmHealthInsurance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
