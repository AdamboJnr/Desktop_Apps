VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmTypesOfAccount 
   Caption         =   "Types Of Account"
   ClientHeight    =   3450
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6360
   LinkTopic       =   "Form2"
   ScaleHeight     =   3450
   ScaleWidth      =   6360
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc dtaInsuranceAccount 
      Height          =   330
      Left            =   240
      Top             =   2640
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
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
      Caption         =   "Insurance Account"
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
   Begin VB.CommandButton cmdGenerateAll 
      Caption         =   "Generate"
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
      Left            =   5160
      Picture         =   "frmTypesOfAccount.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2400
      Width           =   975
   End
   Begin VB.Frame fraTypesOfAccount 
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   6015
      Begin VB.ComboBox cboTypeOfAccount 
         DataSource      =   "dtaInsuranceAccount"
         Height          =   315
         Left            =   2520
         TabIndex        =   3
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label lblTypeOfAccount 
         Alignment       =   1  'Right Justify
         Caption         =   "Insurance Account"
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
         Width           =   1815
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   4920
      Picture         =   "frmTypesOfAccount.frx":0442
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblTypesOfAccountReport 
      Alignment       =   2  'Center
      Caption         =   "Types Of Account"
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
      Width           =   4215
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   5295
   End
End
Attribute VB_Name = "frmTypesOfAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Populate_Combo()
    While dtaInsuranceAccount.Recordset.EOF = False
        cboTypeOfAccount.AddItem dtaInsuranceAccount.Recordset.Fields(2).Value
        dtaInsuranceAccount.Recordset.MoveNext
    Wend
End Sub
Private Sub cmdGenerateAll_Click()
    Dim PolicyType As String
    
    If cboTypeOfAccount.Text = "" Then
        MsgBox "Please Select an Insurance Type", vbInformation
    Else
        'Closing any opened reports
        If denInsuranceTypes.rsInsuranceTypes.State Then
            denInsuranceTypes.rsInsuranceTypes.Close
        End If
        PolicyType = cboTypeOfAccount.Text
        denInsuranceTypes.InsuranceTypes PolicyType
        rptInsuranceTypesPayment.Show
    End If
End Sub
Private Sub Form_Load()
    'Positioning the form
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
    Call Populate_Combo
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmAdminDashboard.Show
End Sub
