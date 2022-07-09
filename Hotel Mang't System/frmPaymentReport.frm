VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPaymentReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Payment Report"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDatePaid 
      DataSource      =   "dtaCustomers"
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   3000
      Width           =   2535
   End
   Begin VB.TextBox txtAmount 
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   2160
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc dtaCustomers 
      Height          =   375
      Left            =   120
      Top             =   4920
      Width           =   2175
      _ExtentX        =   3836
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Ichaweri Hotel Management System.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Ichaweri Hotel Management System.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from tblPayment"
      Caption         =   "Customer Payment"
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
   Begin VB.CommandButton cmdViewAll 
      Caption         =   "View All"
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
      Left            =   4080
      TabIndex        =   4
      Top             =   3720
      Width           =   1095
   End
   Begin VB.ComboBox cboCustomerNumber 
      Height          =   315
      Left            =   2640
      TabIndex        =   3
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label lblDatePaid 
      Alignment       =   1  'Right Justify
      Caption         =   "Date Paid"
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
      TabIndex        =   6
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   1440
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label lblAmount 
      Alignment       =   1  'Right Justify
      Caption         =   "Amount Paid"
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
      TabIndex        =   2
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label lblCustomerNumber 
      Alignment       =   1  'Right Justify
      Caption         =   "Customer Number"
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
      Width           =   1935
   End
   Begin VB.Label lblPaymentReport 
      Caption         =   "Payment Report"
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
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "frmPaymentReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboCustomerNumber_Click()
    Dim SearchVal As Long
    Dim amount As Currency
   
    SearchVal = cboCustomerNumber.Text
    dtaCustomers.Recordset.MoveFirst
    dtaCustomers.Recordset.Find "[Customer ID]= " & SearchVal, 0, adSearchForward
    If dtaCustomers.Recordset.EOF = True Then
        MsgBox ("No Record Found")
        dtaCustomers.Recordset.MoveFirst
    ElseIf dtaCustomers.Recordset.Fields(2).Value = SearchVal Then
         'amount = dtaCustomers.Recordset.Fields(6).Value
         txtAmount.Text = dtaCustomers.Recordset.Fields(6).Value
         txtDatePaid.Text = dtaCustomers.Recordset.Fields(7).Value
    End If
End Sub
Private Sub cmdViewAll_Click()
    rptPayments.Show
End Sub
Private Sub Form_Load()
    'Positioning the Form
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
    'Populating Customer Combo
    dtaCustomers.Recordset.MoveFirst
    While dtaCustomers.Recordset.EOF = False
        cboCustomerNumber.AddItem dtaCustomers.Recordset.Fields(2).Value
        dtaCustomers.Recordset.MoveNext
    Wend
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmAdminDashboard.Show
End Sub
