VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAcceptRejectClaim 
   Caption         =   "Accept/Reject Claim"
   ClientHeight    =   6615
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7380
   LinkTopic       =   "Form2"
   ScaleHeight     =   6615
   ScaleWidth      =   7380
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc dtaPremiumPayments 
      Height          =   375
      Left            =   3960
      Top             =   6240
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
      Caption         =   "Premium Payments"
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
   Begin MSAdodcLib.Adodc dtaAcceptRejectClaim 
      Height          =   375
      Left            =   1800
      Top             =   6240
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
      RecordSource    =   "select * from tblAcceptedRejectedClaims"
      Caption         =   "Accept/Reject Claim"
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
   Begin MSAdodcLib.Adodc dtaClaimNumber 
      Height          =   375
      Left            =   0
      Top             =   6240
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
      RecordSource    =   "select * from tblPolicyClaims"
      Caption         =   "Claim Number"
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
   Begin VB.CommandButton cmdDetails 
      Caption         =   "Details"
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
      Picture         =   "frmAcceptRejectClaim.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton cmdPremiumPayments 
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
      Left            =   1800
      Picture         =   "frmAcceptRejectClaim.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton cmdReject 
      Caption         =   "Reject"
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
      Left            =   5880
      Picture         =   "frmAcceptRejectClaim.frx":0BDF
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Accept"
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
      Left            =   4440
      Picture         =   "frmAcceptRejectClaim.frx":1021
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Frame fraAcceptRejectClaim 
      Height          =   3615
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   6855
      Begin VB.TextBox txtReport 
         Height          =   375
         Left            =   2520
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   2880
         Width           =   2775
      End
      Begin VB.TextBox txtPolicyNumber 
         Height          =   375
         Left            =   2520
         TabIndex        =   13
         Top             =   2280
         Width           =   2775
      End
      Begin VB.TextBox txtPolicyType 
         DataSource      =   "dtaAcceptRejectClaim"
         Height          =   375
         Left            =   2520
         TabIndex        =   11
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox txtClaimantName 
         DataSource      =   "dtaPremiumPayments"
         Height          =   375
         Left            =   2520
         TabIndex        =   9
         Top             =   1080
         Width           =   2775
      End
      Begin VB.ComboBox cboClaimNumber 
         DataSource      =   "dtaClaimNumber"
         Height          =   315
         Left            =   2520
         TabIndex        =   3
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label lblReport 
         Alignment       =   1  'Right Justify
         Caption         =   "Report"
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
         TabIndex        =   14
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label lblClaimNumber 
         Alignment       =   1  'Right Justify
         Caption         =   "Claim Number"
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
         TabIndex        =   12
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblPolicyType 
         Alignment       =   1  'Right Justify
         Caption         =   "Policy Type"
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
         TabIndex        =   10
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lblClaimantName 
         Alignment       =   1  'Right Justify
         Caption         =   "Claimant Name"
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
         TabIndex        =   4
         Top             =   1080
         Width           =   1695
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
         Left            =   360
         TabIndex        =   2
         Top             =   2280
         Width           =   1695
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   5760
      Picture         =   "frmAcceptRejectClaim.frx":1463
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblAcceptReject 
      Alignment       =   2  'Center
      Caption         =   "Accept /Reject Claim"
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
      Width           =   4455
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   960
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   5535
   End
End
Attribute VB_Name = "frmAcceptRejectClaim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SendEmail()
    On Error Resume Next 'Error Checking
    
    Set cdoMsg = CreateObject("CDO.Message")
    Set cdoConf = CreateObject("CDO.Configuration")
    Set cdoFields = cdoConf.Fields
    
    'Send One copy with Google SMTP server
    schema = "http://schemas.Microsoft.com/cdo/configuration/"
    cdoFields.Item(schema & "sendusing") = 2
    cdoFields.Item(schema & "smtpserver") = "smtp.gmail.com"
    cdoFields.Item(schema & "smtpserverport") = 465
    cdoFields.Item(schema & "smtpauthenticate") = 1
    cdoFields.Item(schema & "sendusername") = "AdamboAllan75@gmail.com"
    cdoFields.Item(schema & "sendpassword") = "A1d2am3b4o%^"
    cdoFields.Item(schema & "smtpusess1") = 1
    cdoFields.Update
    
    With cdoMsg
        .To = "GitauRyan@gmail.com"
        .From = "AdamboAllan75@gmail.com"
        .Subject = "Policy Claim"
        .HTMLBody = "The first Trial"
        .AddAttachment
        Set .Configuration = cdoConf
        .Send
    End With
    
    If Err.Number = 0 Then
        MsgBox "Email Sent Succesfully", , "Email"
    Else
        MsgBox "Email Error" & Err.Number, , "Email"
    End If
    
    Set cdoMsg = Nothing
    Set cdoConf = Nothing
    Set cdoFields = Nothing
    
End Sub
Private Sub Populate_Combo()
    While dtaClaimNumber.Recordset.EOF = False
        cboClaimNumber.AddItem dtaClaimNumber.Recordset.Fields(3).Value
        dtaClaimNumber.Recordset.MoveNext
    Wend
End Sub
Private Sub cboClaimNumber_Click()
    Dim searchvalue As Long, PolicyNumber As Long
    
    searchvalue = cboClaimNumber.Text
   
    'Automatically filling in the Claimant Details
    dtaClaimNumber.Recordset.MoveFirst
    dtaClaimNumber.Recordset.Find "[Claim Number]= " & searchvalue, 0, adSearchForward
    If dtaClaimNumber.Recordset.EOF = True Then
        MsgBox "Record Not Found"
        dtaClaimNumber.Recordset.MoveFirst
    ElseIf dtaClaimNumber.Recordset.Fields(3).Value = searchvalue Then
        dtaAcceptRejectClaim.Recordset.MoveFirst
        dtaAcceptRejectClaim.Recordset.Find "[Claim Number]= " & searchvalue, 0, adSearchForward
        If dtaAcceptRejectClaim.Recordset.EOF = True Then
            txtClaimantName.Text = dtaClaimNumber.Recordset.Fields(1).Value
            txtPolicyType.Text = dtaClaimNumber.Recordset.Fields(2).Value
            txtPolicyNumber.Text = dtaClaimNumber.Recordset.Fields(0).Value
        ElseIf dtaAcceptRejectClaim.Recordset.Fields(5).Value = searchvalue Then
            MsgBox "Claim Number Had Already been Used", vbCritical
            cboClaimNumber.Text = ""
            txtClaimantName.Text = ""
            txtPolicyType.Text = ""
            txtReport.Text = ""
            Exit Sub
        End If
    End If
End Sub
Private Sub cmdAccept_Click()
    Dim PolicyNumber As Long, ClaimNumberr As Long
    Dim Status As String, Reason As String
    
    If cboClaimNumber.Text = "" Or txtReport.Text = "" Then
        MsgBox "Please Fill In All Inputs", vbInformation
    Else
        PolicyNumber = txtPolicyNumber.Text
        Status = "Accepted"
        ClaimNumberr = cboClaimNumber.Text
        Reason = txtReport.Text
     

        dtaAcceptRejectClaim.Recordset.AddNew
        dtaAcceptRejectClaim.Recordset.Fields(0).Value = PolicyNumber
        dtaAcceptRejectClaim.Recordset.Fields(1).Value = txtReport.Text
        dtaAcceptRejectClaim.Recordset.Fields(2).Value = Status
        dtaAcceptRejectClaim.Recordset.Fields(3).Value = Format(Now, "mm/dd/yy hh:mm:ss")
        dtaAcceptRejectClaim.Recordset.Fields(4).Value = txtClaimantName.Text
        dtaAcceptRejectClaim.Recordset.Fields(5).Value = ClaimNumberr
        dtaAcceptRejectClaim.Recordset.Update
        S = MsgBox("Claim Number " & " " & ClaimNumberr & " has been Accepted due to" & " " & Reason, vbInformation)
                    
        txtPolicyNumber.Text = ""
        txtClaimantName.Text = ""
        txtPolicyType.Text = ""
        txtReport.Text = ""
        cboClaimNumber.Text = ""
    End If
End Sub
Private Sub cmdDetails_Click()
    Dim PolicyNumber As Long
    Dim PolicyType As String
    If txtPolicyNumber.Text = "" Then
        MsgBox "Please Select a Valid Policy Number", vbInformation
    Else
        PolicyType = txtPolicyType.Text
        If PolicyType = "Fire and Perils Insurance" Then
            'Closing any opened reports
            If denFireAndPerilsClaim.rsFireAndPerilsClaim.State Then
                denFireAndPerilsClaim.rsFireAndPerilsClaim.Close
            End If
            
            PolicyNumber = txtPolicyNumber.Text
            denFireAndPerilsClaim.FireAndPerilsClaim PolicyNumber
            rptFireAndPerilsClaim.Show
        ElseIf PolicyType = "Motor Insurance" Then
            'Closing any opened reports
            If denMotorInsuranceClaim.rsMotorInsuranceClaim.State Then
                denMotorInsuranceClaim.rsMotorInsuranceClaim.Close
            End If
            
            PolicyNumber = txtPolicyNumber.Text
            denMotorInsuranceClaim.MotorInsuranceClaim PolicyNumber
            rptMotorInsuranceClaim.Show
        ElseIf PolicyType = "Accident Insurance" Then
            'Closing any opened reports
            If denAccidentInsuranceClaim.rsAccidentInsurancClaim.State Then
                denAccidentInsuranceClaim.rsAccidentInsurancClaim.Close
            End If
            
            PolicyNumber = txtPolicyNumber.Text
            denAccidentInsuranceClaim.AccidentInsurancClaim PolicyNumber
            rptAccidentInsuranceClaim.Show
        End If
    End If
End Sub
Private Sub cmdPremiumPayments_Click()
    Dim PolicyNumber As Long
    If txtPolicyNumber.Text = "" Then
        MsgBox "Please Fill In A Valid Policy Number", vbCritical
    Else
        PolicyNumber = txtPolicyNumber.Text
        denPremiumPayments.PremiumPayments PolicyNumber
        rptPremiumPayments.Show
    End If
End Sub

Private Sub cmdReject_Click()
    Dim PolicyNumber As Long, ClaimNumberr As Long
    Dim Status As String, Reason As String
    
    If cboClaimNumber.Text = "" Or txtReport.Text = "" Then
        MsgBox "Please Fill In All Inputs", vbInformation
    Else
        PolicyNumber = txtPolicyNumber.Text
        Status = "Rejected"
        ClaimNumberr = cboClaimNumber.Text
        Reason = txtReport.Text
     

        dtaAcceptRejectClaim.Recordset.AddNew
        dtaAcceptRejectClaim.Recordset.Fields(0).Value = PolicyNumber
        dtaAcceptRejectClaim.Recordset.Fields(1).Value = txtReport.Text
        dtaAcceptRejectClaim.Recordset.Fields(2).Value = Status
        dtaAcceptRejectClaim.Recordset.Fields(3).Value = Format(Now, "mm/dd/yy hh:mm:ss")
        dtaAcceptRejectClaim.Recordset.Fields(4).Value = txtClaimantName.Text
        dtaAcceptRejectClaim.Recordset.Fields(5).Value = ClaimNumberr
        dtaAcceptRejectClaim.Recordset.Update
        S = MsgBox("Claim Number " & " " & ClaimNumberr & " has been Accepted due to" & " " & Reason, vbInformation)
                    
        txtPolicyNumber.Text = ""
        txtClaimantName.Text = ""
        txtPolicyType.Text = ""
        txtReport.Text = ""
        cboClaimNumber.Text = ""
    End If
End Sub

Private Sub Command1_Click()
Call SendEmail
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
