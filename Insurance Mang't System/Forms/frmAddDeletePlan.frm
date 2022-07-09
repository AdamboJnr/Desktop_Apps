VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAddPlan 
   Caption         =   "Add/Delete Plan"
   ClientHeight    =   5520
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6750
   LinkTopic       =   "Form2"
   ScaleHeight     =   5520
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc dtaAddPlan 
      Height          =   375
      Left            =   240
      Top             =   4680
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
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
      RecordSource    =   "select * from tblInsuranceType"
      Caption         =   "Add Plan"
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5400
      Picture         =   "frmAddDeletePlan.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdSaveAdd 
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
      Height          =   855
      Left            =   4080
      Picture         =   "frmAddDeletePlan.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4440
      Width           =   975
   End
   Begin VB.Frame fraAddPlan 
      Height          =   2895
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   6015
      Begin VB.TextBox txtInsuranceName 
         DataSource      =   "dtaAddPlan"
         Height          =   375
         Left            =   2400
         TabIndex        =   8
         Top             =   1200
         Width           =   2895
      End
      Begin VB.ComboBox cboInsuranceType 
         Height          =   315
         ItemData        =   "frmAddDeletePlan.frx":0884
         Left            =   2400
         List            =   "frmAddDeletePlan.frx":088E
         TabIndex        =   6
         Top             =   360
         Width           =   2895
      End
      Begin VB.TextBox txtPrice 
         Height          =   375
         Left            =   2400
         TabIndex        =   4
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label lblInsuranceName 
         Alignment       =   1  'Right Justify
         Caption         =   "Insurance Name"
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
         TabIndex        =   7
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lblPrice 
         Alignment       =   1  'Right Justify
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
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label lblInsurancePlan 
         Alignment       =   1  'Right Justify
         Caption         =   "Insurance Type"
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
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   5280
      Picture         =   "frmAddDeletePlan.frx":08AB
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblAddDelete 
      Alignment       =   2  'Center
      Caption         =   "Add Insurance Plan"
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
      Width           =   4215
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   5295
   End
End
Attribute VB_Name = "frmAddPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    If txtInsuranceName.Text = "" And txtPrice.Text = "" And cboInsuranceType.Text = "" Then
        Unload Me
        frmAdminDashboard.Show
    Else
        txtInsuranceName.Text = ""
        txtPrice.Text = ""
    End If
End Sub

Private Sub cmdSaveAdd_Click()
    Dim curPrice As Currency
    
    If cboInsuranceType.Text = "" Or txtInsuranceName.Text = "" Or txtPrice.Text = "" Then
        MsgBox "Please Fill In All Inputs", vbInformation
    Else
        curPrice = txtPrice.Text
        
        dtaAddPlan.Recordset.AddNew
        dtaAddPlan.Recordset.Fields(1).Value = cboInsuranceType.Text
        dtaAddPlan.Recordset.Fields(2).Value = txtInsuranceName.Text
        dtaAddPlan.Recordset.Fields(3).Value = curPrice
        dtaAddPlan.Recordset.Update
    End If
    MsgBox "Succesfully Added a New Plan", vbInformation
    txtInsuranceName.Text = ""
    txtPrice.Text = ""
End Sub
Private Sub Form_Load()
    'Positioning the form
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
    frmAdminDashboard.Show
End Sub
