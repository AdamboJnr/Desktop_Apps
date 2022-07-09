VERSION 5.00
Begin VB.Form frmCarPolicy 
   Caption         =   "Car Policy"
   ClientHeight    =   7950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10050
   LinkTopic       =   "Form2"
   ScaleHeight     =   7950
   ScaleWidth      =   10050
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
      Left            =   8520
      TabIndex        =   19
      Top             =   7080
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
      Left            =   6840
      TabIndex        =   18
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Frame fraCarAccidentDetails 
      Height          =   2415
      Left            =   360
      TabIndex        =   8
      Top             =   4560
      Width           =   9495
      Begin VB.TextBox txtStreet 
         Height          =   375
         Left            =   1440
         TabIndex        =   21
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Frame fraAccidentDate 
         Height          =   735
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   6855
         Begin VB.ComboBox Combo3 
            Height          =   315
            Left            =   5400
            TabIndex        =   15
            Text            =   "2010"
            Top             =   240
            Width           =   975
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   3000
            TabIndex        =   13
            Text            =   "January"
            Top             =   240
            Width           =   1095
         End
         Begin VB.ComboBox cboAccidentDate 
            Height          =   315
            Left            =   1080
            TabIndex        =   11
            Text            =   "1"
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblYearAccident 
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
            Left            =   4320
            TabIndex        =   14
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lblMonthAccident 
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
            Left            =   2040
            TabIndex        =   12
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblAccidentDate 
            Caption         =   "Date"
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
            Width           =   735
         End
      End
      Begin VB.Label lblCaption 
         Caption         =   "Car Accident Details"
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
         TabIndex        =   22
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label lblStreet 
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
         Left            =   240
         TabIndex        =   20
         Top             =   1680
         Width           =   975
      End
   End
   Begin VB.Frame fraCarDetails 
      Height          =   3015
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   4095
      Begin VB.TextBox txtPrice 
         Height          =   375
         Left            =   1680
         TabIndex        =   17
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox txtYearBought 
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox txtRegYear 
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txtVehicle 
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblPrice 
         Caption         =   "Price"
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
         TabIndex        =   16
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label lblYearBought 
         Caption         =   "Year Bought"
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
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lblRegYear 
         Caption         =   "Reg Year"
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
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblVehicle 
         Caption         =   "Vehicle"
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
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Label lblRealInsuranceCarPrice 
      Caption         =   "5000/= Monthly"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   24
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label lblInsuranceCarPrice 
      Alignment       =   1  'Right Justify
      Caption         =   "Price:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   23
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   2055
      Left            =   4800
      Top             =   1320
      Width           =   5055
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   7080
      Picture         =   "frmCarPolicy.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblCarInsuranceCaption 
      Alignment       =   2  'Center
      Caption         =   "Car Insurance"
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
      Width           =   5055
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   6015
   End
End
Attribute VB_Name = "frmCarPolicy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
