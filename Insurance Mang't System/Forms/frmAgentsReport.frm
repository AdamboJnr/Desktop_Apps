VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAgentsReport 
   Caption         =   "Agents Report"
   ClientHeight    =   4245
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6210
   LinkTopic       =   "Form2"
   ScaleHeight     =   4245
   ScaleWidth      =   6210
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc dtaAgentNumber 
      Height          =   375
      Left            =   120
      Top             =   3360
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      RecordSource    =   "select * from tblAgentDetails"
      Caption         =   "AgentsDetails"
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
   Begin VB.CommandButton cmdPayments 
      Caption         =   "Payments"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2640
      Picture         =   "frmAgentsReport.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2880
      Width           =   975
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
      Height          =   975
      Left            =   5040
      Picture         =   "frmAgentsReport.frx":079D
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdAgentsDetails 
      Caption         =   "Agents"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3840
      Picture         =   "frmAgentsReport.frx":0BDF
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   975
   End
   Begin VB.Frame fraAgentsReportDetails 
      Height          =   1335
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   5655
      Begin VB.ComboBox cboAgentNumber 
         DataSource      =   "dtaAgentNumber"
         Height          =   315
         Left            =   2640
         TabIndex        =   3
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label lblAgentNumber 
         Alignment       =   1  'Right Justify
         Caption         =   "Agent Number"
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
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   4800
      Picture         =   "frmAgentsReport.frx":1306
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblAgentsReport 
      Alignment       =   2  'Center
      Caption         =   "Agent's Report"
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
      Width           =   3615
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   840
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "frmAgentsReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Populate_Combo()
    While dtaAgentNumber.Recordset.EOF = False
        cboAgentNumber.AddItem dtaAgentNumber.Recordset.Fields(6).Value
        dtaAgentNumber.Recordset.MoveNext
    Wend
End Sub
Private Sub cmdAgentsDetails_Click()
    rptAgentDetails.Show
End Sub
Private Sub cmdCancel_Click()
    If cboAgentNumber.Text = "" Then
        Unload Me
        frmAdminDashboard.Show
    Else
        cboAgentNumber.Text = ""
    End If
End Sub
Private Sub cmdPayments_Click()
    Dim AgentNumber As Long
    If cboAgentNumber.Text = "" Then
        MsgBox "Please Select a valid Agent Number", vbCritical
        cboAgentNumber.SetFocus
    Else
        If denAgentPayments.rsAgentPayments.State Then
            denAgentPayments.rsAgentPayments.Close
        End If
        AgentNumber = cboAgentNumber.Text
        denAgentPayments.AgentPayments AgentNumber
        rptAgentPayments.Show
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
