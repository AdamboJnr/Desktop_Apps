VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPolicyPaymentsReport 
   Caption         =   "Policy Payments Report"
   ClientHeight    =   3720
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6840
   LinkTopic       =   "Form2"
   ScaleHeight     =   3720
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc dtaPolicyNumbers 
      Height          =   330
      Left            =   240
      Top             =   2760
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      RecordSource    =   "select * from tblAcceptedRejectedPolicies"
      Caption         =   "PolicyDetails"
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
      Left            =   5520
      Picture         =   "frmPolicyPaymentsReport.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdDisplayAll 
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
      Height          =   855
      Left            =   4320
      Picture         =   "frmPolicyPaymentsReport.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display"
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
      Left            =   3120
      Picture         =   "frmPolicyPaymentsReport.frx":0CB9
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   975
   End
   Begin VB.Frame fraPolicyPaymentReport 
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   6255
      Begin VB.ComboBox cboPolicyNumber 
         DataSource      =   "dtaPolicyNumbers"
         Height          =   315
         Left            =   2640
         TabIndex        =   2
         Top             =   360
         Width           =   2055
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
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   4680
      Picture         =   "frmPolicyPaymentsReport.frx":10FB
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblPolicyPaymentsReport 
      Alignment       =   2  'Center
      Caption         =   "Policy Payment Reports"
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
      Width           =   3495
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   960
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "frmPolicyPaymentsReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Populate_Combo()
    While dtaPolicyNumbers.Recordset.EOF = False
        cboPolicyNumber.AddItem dtaPolicyNumbers.Recordset.Fields(0).Value
        dtaPolicyNumbers.Recordset.MoveNext
    Wend
End Sub
Private Sub cmdCancel_Click()
    If cboPolicyNumber.Text = "" Then
        Unload Me
        frmAdminDashboard.Show
    Else
        cboPolicyNumber.Text = ""
    End If
End Sub

Private Sub cmdDisplay_Click()
    Dim PolicyNumber As Long
    If cboPolicyNumber.Text = "" Then
        MsgBox "Please Fill In A Valid Policy Number", vbCritical
    Else
        If denPremiumPayments.rsPremiumPayments.State Then
            denPremiumPayments.rsPremiumPayments.Close
        End If
        PolicyNumber = cboPolicyNumber.Text
        denPremiumPayments.PremiumPayments PolicyNumber
        rptPremiumPayments.Show
    End If
End Sub
Private Sub cmdDisplayAll_Click()
    rptAllPayments.Show
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
