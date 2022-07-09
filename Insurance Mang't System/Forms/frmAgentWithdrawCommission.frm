VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAgentWithdrawCommission 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Withdraw Commission"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9180
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc dtaAgentPayment 
      Height          =   450
      Left            =   1080
      Top             =   5880
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   794
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
      RecordSource    =   "select * from tblCommissionWithdraw"
      Caption         =   "Agent Payment"
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
   Begin MSAdodcLib.Adodc dtaAgentNumber 
      Height          =   375
      Left            =   2280
      Top             =   6480
      Visible         =   0   'False
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Insurance management database.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Insurance management database.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from tblAgentDetails"
      Caption         =   "Agent Number"
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
   Begin MSAdodcLib.Adodc dtaPolicyCreation 
      Height          =   375
      Left            =   120
      Top             =   6480
      Visible         =   0   'False
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
      RecordSource    =   "select * from tblPolicyCreation"
      Caption         =   "PolicyCreation"
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
      Height          =   1095
      Left            =   7560
      Picture         =   "frmAgentWithdrawCommission.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdCheckOut 
      Caption         =   "Check Out"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6000
      Picture         =   "frmAgentWithdrawCommission.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Frame fraCommissionDetails 
      Height          =   1935
      Left            =   240
      TabIndex        =   5
      Top             =   3840
      Width           =   8535
      Begin VB.ComboBox cboPaymentMode 
         Height          =   315
         ItemData        =   "frmAgentWithdrawCommission.frx":0BDF
         Left            =   6000
         List            =   "frmAgentWithdrawCommission.frx":0BEC
         TabIndex        =   13
         Text            =   "Cash"
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txtTotalCommission 
         Height          =   375
         Left            =   2160
         TabIndex        =   11
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtPrice 
         DataSource      =   "dtaAgentPayment"
         Height          =   375
         Left            =   6000
         TabIndex        =   9
         Text            =   "1000"
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtNumberOfCustomers 
         Height          =   375
         Left            =   2160
         TabIndex        =   7
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblPaymentMode 
         Alignment       =   1  'Right Justify
         Caption         =   "Payment Mode"
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
         Left            =   4080
         TabIndex        =   12
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblTotalCommission 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Commission"
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
         TabIndex        =   10
         Top             =   960
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
         Height          =   375
         Left            =   4320
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblNumberOfCustomers 
         Alignment       =   1  'Right Justify
         Caption         =   "No. Of Customers"
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
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame fraAgentWithdrawalDetails 
      Height          =   2295
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   8535
      Begin VB.ListBox lstCustomersName 
         Height          =   1035
         Left            =   6000
         TabIndex        =   20
         Top             =   960
         Width           =   1815
      End
      Begin VB.ComboBox cboMonth 
         DataSource      =   "dtaPolicyCreation"
         Height          =   315
         ItemData        =   "frmAgentWithdrawCommission.frx":0C05
         Left            =   2040
         List            =   "frmAgentWithdrawCommission.frx":0C2D
         TabIndex        =   18
         Top             =   1080
         Width           =   1695
      End
      Begin VB.ComboBox cboAgentNumber 
         DataSource      =   "dtaAgentNumber"
         Height          =   315
         Left            =   2040
         TabIndex        =   17
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtAgentName 
         Height          =   375
         Left            =   6000
         TabIndex        =   4
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblCustomersName 
         Alignment       =   1  'Right Justify
         Caption         =   "Customers"
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
         Left            =   4200
         TabIndex        =   19
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblMonth 
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
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblAgentName 
         Alignment       =   1  'Right Justify
         Caption         =   "Agent's Name"
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
         Left            =   4200
         TabIndex        =   3
         Top             =   360
         Width           =   1455
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
         Width           =   1455
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   6720
      Picture         =   "frmAgentWithdrawCommission.frx":0C93
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblCommissionWithdrawal 
      Alignment       =   2  'Center
      Caption         =   "Commission Withdraw"
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
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   5415
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   1080
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   6375
   End
End
Attribute VB_Name = "frmAgentWithdrawCommission"
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
Private Sub cboAgentNumber_Click()
    Dim lngAgentNumber As Long
    
    lngAgentNumber = cboAgentNumber.Text
    'Finding the Agent Number in db an automatically filling the name input
    dtaAgentNumber.Recordset.MoveFirst
    dtaAgentNumber.Recordset.Find "[Agent Number]= " & lngAgentNumber, 0, adSearchForward
    If dtaAgentNumber.Recordset.EOF = True Then
        MsgBox "No Record Found"
        dtaAgentNumber.Recordset.MoveFirst
    ElseIf dtaAgentNumber.Recordset.Fields(6).Value = lngAgentNumber Then
        txtAgentName.Text = dtaAgentNumber.Recordset.Fields(7).Value
    End If
End Sub
Private Sub cboMonth_Click()
    Dim creationDate As String, Uoption As String
    Dim UserChoice(1 To 12) As String
    Dim intChoice As Integer, output As Integer, x As Integer, count As Integer, lngAgentNumber As Integer
    Dim curTotalAmount As Currency, price As Currency
    
    lstCustomersName.Clear
    Uoption = cboMonth.Text
    UserChoice(1) = "January": UserChoice(2) = "February": UserChoice(3) = "March": UserChoice(4) = "April": UserChoice(5) = "May": UserChoice(6) = "June": UserChoice(7) = "July": UserChoice(8) = "August": UserChoice(9) = "September": UserChoice(10) = "October": UserChoice(11) = "November": UserChoice(12) = "December"
    lngAgentNumber = cboAgentNumber.Text
    
    'Comparing user input to the months in array
    For x = 1 To 12
        If Uoption = UserChoice(x) Then
            intChoice = x
            Exit For
        End If
    Next x
    
    'Finding the agent number and specific month of policy creation
    dtaPolicyCreation.Recordset.MoveFirst
    While dtaPolicyCreation.Recordset.EOF = False
        creationDate = dtaPolicyCreation.Recordset.Fields(8).Value
        output = Val(Mid(creationDate, 1, 1))
        If dtaPolicyCreation.Recordset.Fields(7).Value = lngAgentNumber And output = intChoice Then
            lstCustomersName.AddItem dtaPolicyCreation.Recordset.Fields(1).Value
        End If
        dtaPolicyCreation.Recordset.MoveNext
    Wend
    
    'Automatically filling in all inputs
    count = lstCustomersName.ListCount
    price = Val(txtPrice.Text)
    txtNumberOfCustomers.Text = count
    curTotalAmount = price * CCur(count)
    txtTotalCommission.Text = curTotalAmount
    
End Sub
Private Sub cmdCancel_Click()
    If txtAgentName.Text = "" And txtNumberOfCustomers.Text = "" And txtPrice.Text = "" And txtTotalCommission.Text = "" And cboAgentNumber.Text = "" And cboMonth.Text = "" And cboPaymentMode.Text = "" Then
        Unload Me
        frmMain.Show
    Else
        txtAgentName.Text = ""
        txtNumberOfCustomers.Text = ""
        txtPrice.Text = ""
        txtTotalCommission.Text = ""
        cboAgentNumber.Text = ""
        cboMonth.Text = ""
        cboPaymentMode.Text = ""
        lstCustomersName.Clear
    End If
End Sub
Private Sub cmdCheckOut_Click()
    Dim price As Currency
    Dim TotalCommission As Currency
    Dim lngAgentNumber As Long
    Dim NumberOfCustomer As Integer
    Dim AgentName As String, WithdrawDate As String
    Dim WithdrawId As Long
    
    If txtAgentName.Text = "" Or txtNumberOfCustomers.Text = "" Or txtPrice.Text = "" Or txtTotalCommission.Text = "" Or cboAgentNumber.Text = "" Or cboMonth.Text = "" Or cboPaymentMode.Text = "" Then
        MsgBox "Please Fill In All Inputs", vbCritical
    Else
        price = txtPrice.Text
        TotalCommission = txtTotalCommission.Text
        lngAgentNumber = cboAgentNumber.Text
        NumberOfCustomer = txtNumberOfCustomers.Text
        AgentName = txtAgentName.Text
        'Saving to Database
        dtaAgentPayment.Recordset.AddNew
        dtaAgentPayment.Recordset.Fields(0).Value = lngAgentNumber
        dtaAgentPayment.Recordset.Fields(1).Value = txtAgentName.Text
        dtaAgentPayment.Recordset.Fields(2).Value = NumberOfCustomer
        dtaAgentPayment.Recordset.Fields(3).Value = price
        dtaAgentPayment.Recordset.Fields(4).Value = TotalCommission
        dtaAgentPayment.Recordset.Fields(5).Value = cboMonth.Text
        dtaAgentPayment.Recordset.Fields(6).Value = cboPaymentMode.Text
        dtaAgentPayment.Recordset.Fields(7).Value = Format(Now, "mm/dd/yy hh:mm:ss")
        dtaAgentPayment.Recordset.Update
        dtaAgentPayment.Recordset.MoveLast
        WithdrawDate = dtaAgentPayment.Recordset.Fields(7).Value
        S = MsgBox(AgentName & " " & " has Succesfully Withdrawn " & TotalCommission & " at " & WithdrawDate, vbInformation)

        'Showing Reciept
        dtaAgentPayment.Recordset.MoveLast
        WithdrawId = dtaAgentPayment.Recordset.Fields(8).Value
        denAgentReciept.AgentReciept WithdrawId
        rptAgentReceipt.Show
        
        'Clearing Outputs
        txtAgentName.Text = ""
        txtNumberOfCustomers.Text = ""
        txtPrice.Text = ""
        txtTotalCommission.Text = ""
        cboAgentNumber.Text = ""
        cboMonth.Text = ""
        cboPaymentMode.Text = ""
        lstCustomersName.Clear
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
    frmMain.Show
End Sub
