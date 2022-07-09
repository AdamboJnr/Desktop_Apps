VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAccidentInsurance 
   Caption         =   "Accident Insurance"
   ClientHeight    =   5310
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6465
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc dtaAccidentInsurance 
      Height          =   495
      Left            =   240
      Top             =   4440
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Insurance management database.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Insurance management database.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from tblAccidentInsurance"
      Caption         =   "Accident Insurance"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   855
      Left            =   5160
      Picture         =   "frmAccidentInsurance.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   855
      Left            =   3720
      Picture         =   "frmAccidentInsurance.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Frame fraAccidentInsuranceDetails 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   6015
      Begin VB.TextBox txtAge 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   9
         Top             =   2160
         Width           =   2775
      End
      Begin VB.ComboBox cboCoverType 
         DataSource      =   "dtaAccidentInsurance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmAccidentInsurance.frx":0884
         Left            =   2760
         List            =   "frmAccidentInsurance.frx":088E
         TabIndex        =   7
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtOccupation 
         DataSource      =   "dtaAccidentInsurance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   6
         Top             =   1560
         Width           =   2775
      End
      Begin VB.TextBox txtPolicyNumber 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   3
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label lblAge 
         Alignment       =   1  'Right Justify
         Caption         =   "Age"
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label lblOccupation 
         Alignment       =   1  'Right Justify
         Caption         =   "Occupation"
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label lblCoverType 
         Alignment       =   1  'Right Justify
         Caption         =   "Cover Type"
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label lblPolicyNumber 
         Alignment       =   1  'Right Justify
         Caption         =   "Policy Number"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   4920
      Picture         =   "frmAccidentInsurance.frx":08A3
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Accident Insurance"
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
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   4935
   End
End
Attribute VB_Name = "frmAccidentInsurance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    If txtAge.Text = "" And txtOccupation.Text = "" And txtPolicyNumber.Text = "" And cboCoverType.Text = "" Then
        Unload Me
        frmPolicyCreationForm.Show
    Else
        txtAge.Text = ""
        txtOccupation.Text = ""
        txtPolicyNumber.Text = ""
        cboCoverType.Text = ""
    End If
End Sub

Private Sub cmdSave_Click()
    Dim age As Long
    Dim PolicyNumber As Long
    
    If txtAge.Text = "" Or txtOccupation.Text = "" Or txtPolicyNumber.Text = "" Or cboCoverType.Text = "" Then
        MsgBox "Please Fill In All Inputs"
    Else
        age = txtAge.Text
        PolicyNumber = txtPolicyNumber.Text
        
        dtaAccidentInsurance.Recordset.AddNew
        dtaAccidentInsurance.Recordset.Fields(0).Value = PolicyNumber
        dtaAccidentInsurance.Recordset.Fields(1).Value = cboCoverType.Text
        dtaAccidentInsurance.Recordset.Fields(2).Value = txtOccupation.Text
        dtaAccidentInsurance.Recordset.Fields(3).Value = age
        dtaAccidentInsurance.Recordset.Update
        
        MsgBox "Updated Succesfully"
        
        txtAge.Text = ""
        txtOccupation.Text = ""
        txtPolicyNumber.Text = ""
        cboCoverType.Text = ""
        
    End If
End Sub
Private Sub Form_Load()
    'Positioning the form
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
    'Automatically Filling the Policy Number
    txtPolicyNumber.Text = lngPolicyNumber
    txtAge.Text = lngAge
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmPolicyCreationForm.Show
End Sub
