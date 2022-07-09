VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPaymentsDetails 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Project Details"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAmount 
      DataSource      =   "dtaRooms"
      Height          =   375
      Left            =   2640
      TabIndex        =   12
      Top             =   4080
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc dtaRooms 
      Height          =   375
      Left            =   3720
      Top             =   5400
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      RecordSource    =   "select * from tblRoom"
      Caption         =   "Book Room"
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
   Begin MSAdodcLib.Adodc dtaEmployeeDetails 
      Height          =   375
      Left            =   240
      Top             =   4800
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Ichaweri Hotel Management System.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Ichaweri Hotel Management System.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from tblEmployee"
      Caption         =   "Employee"
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
   Begin MSAdodcLib.Adodc dtaCustomerDetails 
      Height          =   330
      Left            =   1920
      Top             =   5400
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Ichaweri Hotel Management System.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Ichaweri Hotel Management System.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from tblCheckIn"
      Caption         =   "Customer Details"
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
   Begin VB.TextBox txtRoomNumber 
      DataSource      =   "dtaPaymentDetails"
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   3360
      Width           =   2535
   End
   Begin VB.TextBox txtCustomerName 
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Top             =   1920
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc dtaPaymentDetails 
      Height          =   375
      Left            =   0
      Top             =   5400
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "Payment Details"
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
      Height          =   495
      Left            =   4080
      TabIndex        =   7
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdPay 
      Caption         =   "Pay"
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
      Left            =   2760
      TabIndex        =   6
      Top             =   4800
      Width           =   1095
   End
   Begin VB.ComboBox cboPaymentMode 
      Height          =   315
      ItemData        =   "frmPaymentsDetails.frx":0000
      Left            =   2640
      List            =   "frmPaymentsDetails.frx":000D
      TabIndex        =   5
      Top             =   2640
      Width           =   2535
   End
   Begin VB.TextBox txtCustomerID 
      DataSource      =   "dtaCustomerDetails"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label lblAmount 
      Alignment       =   1  'Right Justify
      Caption         =   "Amount"
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
      TabIndex        =   11
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label lblCustomerName 
      Alignment       =   1  'Right Justify
      Caption         =   "CustomerName"
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
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   3720
      Picture         =   "frmPaymentsDetails.frx":0026
      Stretch         =   -1  'True
      Top             =   240
      Width           =   615
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label lblRoomNumber 
      Alignment       =   1  'Right Justify
      Caption         =   "Room Number"
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
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label lblPaymentMode 
      Alignment       =   1  'Right Justify
      Caption         =   "Payment mode"
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
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblCustomeID 
      Alignment       =   1  'Right Justify
      Caption         =   "Customer ID"
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
      TabIndex        =   1
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblPaymentDetails 
      Alignment       =   2  'Center
      Caption         =   "Payment Details"
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
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmPaymentsDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    cboPaymentMode.Text = ""
    txtCustomerID.Text = ""
    txtCustomerName.Text = ""
    txtRoomNumber.Text = ""
    'cboEmployeeNumber.Text = ""
    txtAmount.Text = ""
End Sub
Private Sub cmdPay_Click()
    Dim lngRoomNumber As Long
    Dim curAmount As Currency
    Dim SearchValue As Long, PaymentId As Long
    
    If txtCustomerID.Text = "" Or txtCustomerName.Text = "" Or txtRoomNumber.Text = "" Or cboPaymentMode.Text = "" Or txtAmount.Text = "" Then
        MsgBox "Please Fill In All Inputs"
    Else
        SearchValue = txtRoomNumber.Text
        
        lngRoomNumber = txtRoomNumber.Text
        curAmount = txtAmount.Text
        
        'saving to the database
        dtaPaymentDetails.Recordset.AddNew
        dtaPaymentDetails.Recordset.Fields(1).Value = cboPaymentMode.Text
        dtaPaymentDetails.Recordset.Fields(2).Value = lngCustomerId
        dtaPaymentDetails.Recordset.Fields(3).Value = txtCustomerName.Text
        dtaPaymentDetails.Recordset.Fields(4).Value = lngRoomNumber
        dtaPaymentDetails.Recordset.Fields(5).Value = lngEmployeeNumber
        dtaPaymentDetails.Recordset.Fields(6).Value = curAmount
        dtaPaymentDetails.Recordset.Fields(7).Value = Format(Now, "mm/dd/yy hh:mm:ss")
        dtaPaymentDetails.Recordset.Update
        MsgBox "payment complete"
        
        dtaPaymentDetails.Recordset.MoveLast
        PaymentId = dtaPaymentDetails.Recordset.Fields(0).Value
        
        'Booked Room
        dtaRooms.Recordset.MoveFirst
        dtaRooms.Recordset.Find "[Room Number]= " & SearchValue, 0, adSearchForward
        If dtaRooms.Recordset.Fields(0).Value = SearchValue Then
            dtaRooms.Recordset.Fields(2).Value = "Unavailable"
            dtaRooms.Recordset.Update
        End If
        
        denCustomerReceipt.Receipt PaymentId
        rptCustomerReciept.Show
        
        'Clearing Inputs
        cboPaymentMode.Text = ""
        txtCustomerID.Text = ""
        txtCustomerName.Text = ""
        txtRoomNumber.Text = ""
        txtAmount.Text = ""
    End If
End Sub

Private Sub Form_Load()
    Dim SearchValue As Long
     
    'Positioning the Form
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
    
    txtCustomerID.Text = lngCustomerId
    dtaCustomerDetails.Recordset.MoveFirst
    While dtaCustomerDetails.Recordset.EOF = False
        If dtaCustomerDetails.Recordset.Fields(0).Value = lngCustomerId Then
           txtCustomerName.Text = dtaCustomerDetails.Recordset.Fields(1).Value
           txtRoomNumber.Text = dtaCustomerDetails.Recordset.Fields(5).Value
        End If
        dtaCustomerDetails.Recordset.MoveNext
    Wend
    
    'Automatically setting up the Amount
    SearchValue = txtRoomNumber.Text

    dtaRooms.Recordset.Find "[Room Number]= " & SearchValue, 0, adSearchForward
    If dtaRooms.Recordset.EOF = True Then
        MsgBox ("No Record Found")
        dtaRooms.Recordset.MoveFirst
    ElseIf dtaRooms.Recordset.Fields(0).Value = SearchValue Then
        txtAmount.Text = dtaRooms.Recordset.Fields(4).Value
    End If
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
    Unload frmRoomBooking
    frmMain.Show
End Sub


