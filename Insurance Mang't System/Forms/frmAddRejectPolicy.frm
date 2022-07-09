VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAddRejectPolicy 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Policies"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc dtaAcceptedRejectedPolicies 
      Height          =   375
      Left            =   120
      Top             =   4920
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
      RecordSource    =   "Select * from tblAcceptedRejectedPolicies"
      Caption         =   "Accepted Rejected Policies"
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
      Left            =   0
      Top             =   5400
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
      Left            =   4680
      Picture         =   "frmAddRejectPolicy.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4920
      Width           =   855
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
      Left            =   3600
      Picture         =   "frmAddRejectPolicy.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton cmdView 
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
      Left            =   2520
      Picture         =   "frmAddRejectPolicy.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4920
      Width           =   855
   End
   Begin VB.Frame fraViewPolicy 
      Height          =   3495
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   5295
      Begin VB.TextBox txtReason 
         Height          =   375
         Left            =   2400
         TabIndex        =   12
         Top             =   2760
         Width           =   2415
      End
      Begin VB.TextBox txtHolderName 
         Height          =   375
         Left            =   2400
         TabIndex        =   10
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox txtPolicyType 
         DataSource      =   "dtaAcceptedRejectedPolicies"
         Height          =   375
         Left            =   2400
         TabIndex        =   8
         Top             =   2040
         Width           =   2415
      End
      Begin VB.ComboBox cboPolicyNumber 
         DataSource      =   "dtaPolicyDetails"
         Height          =   315
         Left            =   2400
         TabIndex        =   3
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label lblReason 
         Alignment       =   1  'Right Justify
         Caption         =   "Reason"
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
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label lblHolderName 
         Alignment       =   1  'Right Justify
         Caption         =   "Holder Name"
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
         TabIndex        =   9
         Top             =   1320
         Width           =   1575
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
         TabIndex        =   7
         Top             =   2040
         Width           =   1575
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
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   4440
      Picture         =   "frmAddRejectPolicy.frx":0CC6
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblPolicyConfirmation 
      Alignment       =   2  'Center
      Caption         =   "Confirm Policy"
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
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "frmAddRejectPolicy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Populate_Combo()
    While dtaPolicyDetails.Recordset.EOF = False
        cboPolicyNumber.AddItem dtaPolicyDetails.Recordset.Fields(0).Value
        dtaPolicyDetails.Recordset.MoveNext
    Wend
End Sub
Private Sub cboPolicyNumber_Click()
    Dim searchvalue As Long, PolicyNumber As Long
    
    searchvalue = cboPolicyNumber.Text
    PolicyNumber = cboPolicyNumber
    'Automatically filling in the Holders Name
    dtaPolicyDetails.Recordset.MoveFirst
    dtaPolicyDetails.Recordset.Find "[Policy Number]= " & searchvalue, 0, adSearchForward
    If dtaPolicyDetails.Recordset.EOF = True Then
        MsgBox "Record Not Found"
        dtaPolicyDetails.Recordset.MoveFirst
    ElseIf dtaPolicyDetails.Recordset.Fields(0).Value = searchvalue Then
        dtaAcceptedRejectedPolicies.Recordset.MoveFirst
        dtaAcceptedRejectedPolicies.Recordset.Find "[Policy Number]= " & PolicyNumber, 0, adSearchForward
        If dtaAcceptedRejectedPolicies.Recordset.EOF = True Then
            txtHolderName.Text = dtaPolicyDetails.Recordset.Fields(1).Value
            txtPolicyType.Text = dtaPolicyDetails.Recordset.Fields(5).Value
        ElseIf dtaAcceptedRejectedPolicies.Recordset.Fields(0).Value = PolicyNumber Then
            MsgBox "Policy Number Had Already been Used", vbCritical
            cboPolicyNumber.Text = ""
            txtHolderName.Text = ""
            txtPolicyType.Text = ""
            txtReason.Text = ""
            Exit Sub
        End If
    End If
End Sub
Private Sub cmdAccept_Click()
    Dim PolicyNumber As Long
    Dim Status As String
    
    If cboPolicyNumber.Text = "" Or txtReason.Text = "" Then
        MsgBox "Please Fill In All Inputs", vbInformation
    Else
        PolicyNumber = cboPolicyNumber.Text
        Status = "Accepted"
        Reason = txtReason.Text
     
        dtaAcceptedRejectedPolicies.Recordset.AddNew
        dtaAcceptedRejectedPolicies.Recordset.Fields(0).Value = PolicyNumber
        dtaAcceptedRejectedPolicies.Recordset.Fields(1).Value = txtHolderName.Text
        dtaAcceptedRejectedPolicies.Recordset.Fields(2).Value = txtPolicyType.Text
        dtaAcceptedRejectedPolicies.Recordset.Fields(3).Value = txtReason.Text
        dtaAcceptedRejectedPolicies.Recordset.Fields(4).Value = Status
        dtaAcceptedRejectedPolicies.Recordset.Fields(5).Value = Format(Now, "mm/dd/yy hh:mm:ss")
        dtaAcceptedRejectedPolicies.Recordset.Update
        S = MsgBox("Policy Number " & " " & PolicyNumber & " has been Accepted due to" & " " & Reason, vbInformation)
                    
        cboPolicyNumber.Text = ""
        txtHolderName.Text = ""
        txtPolicyType.Text = ""
        txtReason.Text = ""
    End If
End Sub
Private Sub cmdReject_Click()
    Dim PolicyNumber As Long
    Dim Status As String
    
    If cboPolicyNumber.Text = "" Or txtReason.Text = "" Then
        MsgBox "Please Fill In All Inputs", vbInformation
    Else
        PolicyNumber = cboPolicyNumber.Text
        Status = "Rejected"
        Reason = txtReason.Text
        
        dtaAcceptedRejectedPolicies.Recordset.AddNew
        dtaAcceptedRejectedPolicies.Recordset.Fields(0).Value = PolicyNumber
        dtaAcceptedRejectedPolicies.Recordset.Fields(1).Value = txtHolderName.Text
        dtaAcceptedRejectedPolicies.Recordset.Fields(2).Value = txtPolicyType.Text
        dtaAcceptedRejectedPolicies.Recordset.Fields(3).Value = txtReason.Text
        dtaAcceptedRejectedPolicies.Recordset.Fields(4).Value = Status
        dtaAcceptedRejectedPolicies.Recordset.Fields(5).Value = Format(Now, "mm/dd/yy hh:mm:ss")
        dtaAcceptedRejectedPolicies.Recordset.Update
        S = MsgBox("Policy Number " & " " & PolicyNumber & " has been Rejected due to" & " " & Reason, vbInformation)
        
        cboPolicyNumber.Text = ""
        txtHolderName.Text = ""
        txtPolicyType.Text = ""
        txtReason.Text = ""
    End If
End Sub
Private Sub cmdView_Click()
    Dim PolicyNumber As Long
    Dim PolicyType As String
    If cboPolicyNumber.Text = "" Then
        MsgBox "Please Select a Valid Policy Number", vbInformation
    Else
        PolicyType = txtPolicyType.Text
        If PolicyType = "Fire and Perils Insurance" Then
            PolicyNumber = cboPolicyNumber.Text
            denFireAndPerils.FireAndPerils PolicyNumber
            rptFireAndPerilsAcceptReject.Show
        ElseIf PolicyType = "Motor Insurance" Then
            PolicyNumber = cboPolicyNumber.Text
            denMotorInsurance.MotorInsurance PolicyNumber
            rptMotorInsuranceAcceptReject.Show
        ElseIf PolicyType = "Accident Insurance" Then
            PolicyNumber = cboPolicyNumber.Text
            denAccidentInsurance.AccidentInsurance PolicyNumber
            rptAccidentInsuranceAcceptReject.Show
        End If
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
