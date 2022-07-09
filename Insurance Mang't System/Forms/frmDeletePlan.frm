VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmDeletePlan 
   Caption         =   "Form1"
   ClientHeight    =   4545
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc dtaDeletePolicy 
      Height          =   375
      Left            =   240
      Top             =   3480
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "Delete Policy"
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
      Left            =   4680
      Picture         =   "frmDeletePlan.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
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
      Left            =   3360
      Picture         =   "frmDeletePlan.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3480
      Width           =   975
   End
   Begin VB.Frame fraPlanDetails 
      Height          =   1935
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   5535
      Begin VB.TextBox txtInsuranceName 
         Height          =   375
         Left            =   2160
         TabIndex        =   5
         Top             =   1080
         Width           =   2655
      End
      Begin VB.ComboBox cboInsuranceNumber 
         DataSource      =   "dtaDeletePolicy"
         Height          =   315
         Left            =   2160
         TabIndex        =   3
         Top             =   360
         Width           =   2655
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
         TabIndex        =   4
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblInsuranceType 
         Alignment       =   1  'Right Justify
         Caption         =   "Insurance Number"
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
         Width           =   1575
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   4320
      Picture         =   "frmDeletePlan.frx":0884
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblDeletePlan 
      Alignment       =   2  'Center
      Caption         =   "Delete Plan"
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
      Width           =   3255
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   840
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "frmDeletePlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Populate_Combo()
    While dtaDeletePolicy.Recordset.EOF = False
        cboInsuranceNumber.AddItem dtaDeletePolicy.Recordset.Fields(0).Value
        dtaDeletePolicy.Recordset.MoveNext
    Wend
End Sub
Private Sub cboInsuranceNumber_Click()
    Dim searchvalue As Long
    
    searchvalue = cboInsuranceNumber.Text
    'Automatically filling in the Insurance Name
    dtaDeletePolicy.Recordset.MoveFirst
    dtaDeletePolicy.Recordset.Find "[InsuranceType Number]= " & searchvalue, 0, adSearchForward
    If dtaDeletePolicy.Recordset.EOF = True Then
        MsgBox "Record Not Found"
        dtaDeletePolicy.Recordset.MoveFirst
    ElseIf dtaDeletePolicy.Recordset.Fields(0).Value = searchvalue Then
        txtInsuranceName.Text = dtaDeletePolicy.Recordset.Fields(2).Value
    End If
End Sub
Private Sub cmdCancel_Click()
    If cboInsuranceNumber.Text = "" And txtInsuranceName.Text = "" Then
        Unload Me
        frmAdminDashboard.Show
    Else
        cboInsuranceNumber.Text = ""
        txtInsuranceName.Text = ""
    End If
End Sub
Private Sub cmdDelete_Click()
    Dim searchvalue As Long
    
    searchvalue = cboInsuranceNumber.Text
    
    dtaDeletePolicy.Recordset.MoveFirst
    dtaDeletePolicy.Recordset.Find "[InsuranceType Number]= " & searchvalue, 0, adSearchForward
    If dtaDeletePolicy.Recordset.EOF = True Then
        MsgBox "Record Not Found"
        dtaDeletePolicy.Recordset.MoveFirst
    ElseIf dtaDeletePolicy.Recordset.Fields(0).Value = searchvalue Then
        If MsgBox("Are you sure you want to delete Plan?", vbOKCancel + vbQuestion) = vbOK Then
            dtaDeletePolicy.Recordset.Delete
            MsgBox "Insurance Plan Deleted Succesfully", vbInformation
        End If
        cboInsuranceNumber.Clear
        txtInsuranceName.Text = ""
        dtaDeletePolicy.Refresh
        Call Populate_Combo
    End If
End Sub

Private Sub Form_Load()
    'Positioning the form
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
    Call Populate_Combo
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
    frmAdminDashboard.Show
End Sub
