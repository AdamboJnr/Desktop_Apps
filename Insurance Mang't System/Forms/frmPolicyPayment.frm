VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPolicyPayment 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Policy Payment"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9555
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   9555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc dtaPolicyAmoun 
      Height          =   375
      Left            =   0
      Top             =   6000
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
      RecordSource    =   "select * from tblInsuranceType"
      Caption         =   "Policy Amount"
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
   Begin MSAdodcLib.Adodc dtaPolicyPayment 
      Height          =   375
      Left            =   4320
      Top             =   5880
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
      RecordSource    =   "select * from tblPolicyPayment"
      Caption         =   "Policy Payment"
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
   Begin MSAdodcLib.Adodc dtaPolicyDetails 
      Height          =   375
      Left            =   2280
      Top             =   6000
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Insurance management database.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Insurance management database.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from tblAcceptedRejectedPolicies"
      Caption         =   "Policy Details"
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
      Left            =   8160
      Picture         =   "frmPolicyPayment.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5280
      Width           =   975
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
      Height          =   855
      Left            =   6840
      Picture         =   "frmPolicyPayment.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Frame fraPolicyPayment 
      Height          =   3735
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   9015
      Begin VB.TextBox txtHolderName 
         Height          =   375
         Left            =   1920
         TabIndex        =   21
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtRemainingBalance 
         Height          =   375
         Left            =   1920
         TabIndex        =   19
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox txtAmountPaid 
         DataSource      =   "dtaPolicyPayment"
         Height          =   375
         Left            =   1920
         TabIndex        =   17
         Top             =   3000
         Width           =   1815
      End
      Begin VB.ComboBox cboPaymentPlan 
         Height          =   315
         ItemData        =   "frmPolicyPayment.frx":0884
         Left            =   6000
         List            =   "frmPolicyPayment.frx":0891
         TabIndex        =   15
         Top             =   960
         Width           =   2055
      End
      Begin VB.ComboBox cboPolicyNumber 
         DataSource      =   "dtaPolicyDetails"
         Height          =   315
         Left            =   1920
         TabIndex        =   13
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtInsuranceType 
         Height          =   375
         Left            =   6000
         TabIndex        =   12
         Top             =   1680
         Width           =   2055
      End
      Begin VB.ComboBox cboPaymentMode 
         Height          =   315
         ItemData        =   "frmPolicyPayment.frx":08DD
         Left            =   6000
         List            =   "frmPolicyPayment.frx":08E7
         TabIndex        =   9
         Text            =   "Cash"
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtRecievedBy 
         DataSource      =   "dtaPolicyPayment"
         Height          =   375
         Left            =   6000
         TabIndex        =   7
         Top             =   2400
         Width           =   2055
      End
      Begin VB.TextBox txtPolicyPaymentAmount 
         DataSource      =   "dtaPolicyAmoun"
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lblHolderName 
         Alignment       =   1  'Right Justify
         Caption         =   "Holder Name:"
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
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblRemainingBalance 
         Alignment       =   1  'Right Justify
         Caption         =   "Rem Balance"
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
         TabIndex        =   18
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label lblAmountPaid 
         Alignment       =   1  'Right Justify
         Caption         =   "Amount Paid:"
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
         TabIndex        =   16
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label lblPaymentPlan 
         Alignment       =   1  'Right Justify
         Caption         =   "Payment Plan:"
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
         TabIndex        =   14
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblPaymentMode 
         Alignment       =   1  'Right Justify
         Caption         =   "Payment Mode:"
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
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblRecievedBy 
         Alignment       =   1  'Right Justify
         Caption         =   "Recieved By:"
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
         TabIndex        =   6
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label lblPolicyPaymentAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Amount:"
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
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label lblInsuranceType 
         Alignment       =   1  'Right Justify
         Caption         =   "Insurance Type:"
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
         TabIndex        =   3
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label lblPolicyNumber 
         Alignment       =   1  'Right Justify
         Caption         =   "Policy Number:"
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
         Width           =   1455
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   7560
      Picture         =   "frmPolicyPayment.frx":08F8
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblPolicyPaymentModule 
      Alignment       =   2  'Center
      Caption         =   "Policy Payment"
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
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   6135
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   7095
   End
End
Attribute VB_Name = "frmPolicyPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub PopulateCombo()
    While dtaPolicyDetails.Recordset.EOF = False
        If dtaPolicyDetails.Recordset.Fields(4).Value = "Accepted" Then
            cboPolicyNumber.AddItem dtaPolicyDetails.Recordset.Fields(0).Value
        End If
        dtaPolicyDetails.Recordset.MoveNext
    Wend
End Sub
Private Sub cboPolicyNumber_Click()
    Dim lngTempPolicyNumber As Long, PolicyAmount As Long
    Dim PolicyType As String, PaymentPlan As String
    Dim curAmount As Currency, curBalanceRem As Currency, curTotalBalanceRem As Currency
    
    lngTempPolicyNumber = cboPolicyNumber.Text
    'strPolicyName = txtInsuranceType.Text
    
    
    'Filling in type of insurance and user
    dtaPolicyDetails.Recordset.MoveFirst
    dtaPolicyDetails.Recordset.Find "[Policy Number]= " & lngTempPolicyNumber, 0, adSearchForward
    If dtaPolicyDetails.Recordset.EOF = True Then
        dtaPolicyDetails.Recordset.MoveFirst
    ElseIf dtaPolicyDetails.Recordset.Fields(0).Value = lngTempPolicyNumber Then
        txtInsuranceType.Text = dtaPolicyDetails.Recordset.Fields(2).Value
        txtHolderName.Text = dtaPolicyDetails.Recordset.Fields(1).Value
        txtRecievedBy.Text = strUser
    End If
    PolicyType = txtInsuranceType.Text
    'filling in the amount
    dtaPolicyAmoun.Recordset.MoveFirst
    dtaPolicyAmoun.Recordset.Find "InsuranceName='" & PolicyType & "'", 0, adSearchForward
    If dtaPolicyAmoun.Recordset.EOF = True Then
        dtaPolicyAmoun.Recordset.MoveFirst
    ElseIf dtaPolicyAmoun.Recordset.Fields(2).Value = PolicyType Then
        PolicyAmount = dtaPolicyAmoun.Recordset.Fields(3).Value
        txtPolicyPaymentAmount.Text = PolicyAmount
    End If
        
    'Finding if User had Paid
    dtaPolicyPayment.Recordset.MoveLast
    dtaPolicyPayment.Recordset.Find "[Policy Number]= " & lngTempPolicyNumber, 0, adSearchBackward
    'User who has Never Paid
    If dtaPolicyPayment.Recordset.BOF = True Then
       
        
        curRemBalance = txtPolicyPaymentAmount.Text
        txtRemainingBalance = curRemBalance
        dtaPolicyDetails.Recordset.MoveLast
    'User Who has paid at least Once
    ElseIf dtaPolicyPayment.Recordset.Fields(0).Value = lngTempPolicyNumber Then
        PaymentPlan = dtaPolicyPayment.Recordset.Fields(9).Value
        curBalanceRem = dtaPolicyPayment.Recordset.Fields(8).Value
        cboPaymentPlan.Text = PaymentPlan
        
        If curBalanceRem = 0 Then
            MsgBox "User has Already Finished the Premium Payment", vbCritical
            
            cboPaymentMode.Text = ""
            cboPaymentPlan.Text = ""
            cboPolicyNumber.Text = ""
            txtInsuranceType.Text = ""
            txtPolicyPaymentAmount.Text = ""
            txtRecievedBy.Text = ""
            txtAmountPaid.Text = ""
            txtRemainingBalance.Text = curBalanceRem
        ElseIf curBalanceRem < 0 Then
            S = MsgBox("User has an overdraft of" & curBalanceRem & "To Be carried forward Next Year", vbInformation)
            
            cboPaymentMode.Text = ""
            cboPaymentPlan.Text = ""
            cboPolicyNumber.Text = ""
            txtInsuranceType.Text = ""
            txtPolicyPaymentAmount.Text = ""
            txtRecievedBy.Text = ""
            txtAmountPaid.Text = ""
            txtRemainingBalance.Text = curBalanceRem
        Else
            txtRemainingBalance.Text = curBalanceRem
        End If
    End If
End Sub
Private Sub cmdCancel_Click()
    If cboPaymentMode.Text = "" And cboPaymentPlan.Text = "" And cboPolicyNumber.Text = "" And txtInsuranceType.Text = "" And txtPolicyPaymentAmount.Text = "" And txtRecievedBy.Text = "" Then
        Unload Me
        frmMain.Show
    Else
        cboPaymentMode.Text = ""
        cboPaymentPlan.Text = ""
        cboPolicyNumber.Text = ""
        txtInsuranceType.Text = ""
        txtPolicyPaymentAmount.Text = ""
        txtRecievedBy.Text = ""
        txtAmountPaid.Text = ""
        txtRemainingBalance.Text = ""
    End If
End Sub
Private Sub cmdSave_Click()
    Dim PolicyNumber As Long, PaymentId As Long
    Dim curAmount As Currency, curBalanceRem As Currency, curTotalBalanceRem As Currency, curTotalPremiumAmount As Currency
    
    If txtInsuranceType.Text = "" Or txtPolicyPaymentAmount.Text = "" Or txtRecievedBy.Text = "" Or cboPaymentMode.Text = "" Or cboPaymentPlan.Text = "" Or cboPolicyNumber.Text = "" Or txtAmountPaid.Text = "" Or txtRemainingBalance.Text = "" Then
        MsgBox "Please Fill In All Inputs", vbCritical
    Else
        PolicyNumber = cboPolicyNumber.Text
        curAmount = txtAmountPaid.Text
        curTotalPremiumAmount = txtPolicyPaymentAmount.Text
        
        'Finding if User had Paid
        dtaPolicyPayment.Recordset.MoveLast
        dtaPolicyPayment.Recordset.Find "[Policy Number]= " & PolicyNumber, 0, adSearchBackward
        'User Who hasn't Paid
        If dtaPolicyPayment.Recordset.BOF = True Then
            curTotalBalanceRem = curTotalPremiumAmount - curAmount
            dtaPolicyPayment.Recordset.AddNew
            dtaPolicyPayment.Recordset.Fields(0).Value = PolicyNumber
            dtaPolicyPayment.Recordset.Fields(1).Value = txtInsuranceType.Text
            dtaPolicyPayment.Recordset.Fields(2).Value = curAmount
            dtaPolicyPayment.Recordset.Fields(3).Value = txtRecievedBy.Text
            dtaPolicyPayment.Recordset.Fields(4).Value = cboPaymentMode.Text
            dtaPolicyPayment.Recordset.Fields(5).Value = Format(Now, "mm/dd/yy hh:mm:ss")
            dtaPolicyPayment.Recordset.Fields(7).Value = curAmount
            dtaPolicyPayment.Recordset.Fields(8).Value = curTotalBalanceRem
            dtaPolicyPayment.Recordset.Fields(9).Value = cboPaymentPlan.Text
            dtaPolicyPayment.Recordset.Fields(10).Value = txtHolderName.Text
            dtaPolicyPayment.Recordset.Update
            MsgBox "Payment Was Succesfull", vbInformation
        'User Who had Paid atleast once
        ElseIf dtaPolicyPayment.Recordset.Fields(0).Value = PolicyNumber Then
            curBalanceRem = dtaPolicyPayment.Recordset.Fields(8).Value
           
            curTotalBalanceRem = curBalanceRem - curAmount
            dtaPolicyPayment.Recordset.AddNew
            dtaPolicyPayment.Recordset.Fields(0).Value = PolicyNumber
            dtaPolicyPayment.Recordset.Fields(1).Value = txtInsuranceType.Text
            dtaPolicyPayment.Recordset.Fields(2).Value = curAmount
            dtaPolicyPayment.Recordset.Fields(3).Value = txtRecievedBy.Text
            dtaPolicyPayment.Recordset.Fields(4).Value = cboPaymentMode.Text
            dtaPolicyPayment.Recordset.Fields(5).Value = Format(Now, "mm/dd/yy hh:mm:ss")
            dtaPolicyPayment.Recordset.Fields(7).Value = curAmount
            dtaPolicyPayment.Recordset.Fields(8).Value = curTotalBalanceRem
            dtaPolicyPayment.Recordset.Fields(9).Value = cboPaymentPlan.Text
            dtaPolicyPayment.Recordset.Fields(10).Value = txtHolderName.Text
            dtaPolicyPayment.Recordset.Update
            MsgBox "Payment Succesfull", vbInformation
        End If
        'Showing Reciept
        dtaPolicyPayment.Recordset.MoveLast
        PaymentId = dtaPolicyPayment.Recordset.Fields(6).Value
        denPaymentReciept.PaymentReciept PaymentId
        rptPaymentReciept.Show
        
        'Clearing all Inputs
        cboPaymentPlan.Text = ""
        cboPolicyNumber.Text = ""
        txtInsuranceType.Text = ""
        txtPolicyPaymentAmount.Text = ""
        txtRecievedBy.Text = ""
        txtAmountPaid.Text = ""
        txtRemainingBalance = ""
    End If
End Sub
Private Sub Form_Load()
    
    'Positioning the form
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
    
    Call PopulateCombo
    
    txtRecievedBy.Text = strUser
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
    frmMain.Show
End Sub
