VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPolicyCreationForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Policy Creation Form"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   10680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc dtaAgentNumber 
      Height          =   330
      Left            =   960
      Top             =   7800
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
   Begin MSAdodcLib.Adodc dtaTypeOfAccount 
      Height          =   375
      Left            =   3960
      Top             =   7200
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
      Caption         =   "Account Type"
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
   Begin MSAdodcLib.Adodc dtaNextOfKin 
      Height          =   375
      Left            =   2040
      Top             =   7200
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
      RecordSource    =   "select * from tblNextOfKin"
      Caption         =   "Next Of Kin"
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
      Left            =   240
      Top             =   7200
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
      RecordSource    =   "select * from tblPolicyCreation"
      Caption         =   "Policy Creation"
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
      Left            =   9360
      Picture         =   "frmPolicyCreationForm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7200
      Width           =   975
   End
   Begin VB.CommandButton cmdProceed 
      Caption         =   "Proceed"
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
      Left            =   7920
      Picture         =   "frmPolicyCreationForm.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Frame fraNextOfKin 
      Height          =   2055
      Left            =   240
      TabIndex        =   3
      Top             =   4920
      Width           =   10095
      Begin VB.ComboBox cboNextOfKinRelation 
         Height          =   315
         ItemData        =   "frmPolicyCreationForm.frx":0884
         Left            =   7920
         List            =   "frmPolicyCreationForm.frx":0894
         TabIndex        =   25
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtNextOfKinResidence 
         DataSource      =   "dtaTypeOfAccount"
         Height          =   375
         Left            =   1680
         TabIndex        =   23
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtNextOfKinPhoneNo 
         Height          =   375
         Left            =   4800
         TabIndex        =   20
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtNextOfKinName 
         DataSource      =   "dtaNextOfKin"
         Height          =   375
         Left            =   1680
         TabIndex        =   17
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblNextOfKinResidence 
         Alignment       =   1  'Right Justify
         Caption         =   "Residence"
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
         TabIndex        =   22
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblNextOfKinRelation 
         Alignment       =   1  'Right Justify
         Caption         =   "Relation"
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
         Left            =   6480
         TabIndex        =   21
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblNextOfKinPhoneNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Mobile No"
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
         Left            =   3360
         TabIndex        =   19
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblNextOfKin 
         Caption         =   "Next Of Kin"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label lblNextOfKinName 
         Alignment       =   1  'Right Justify
         Caption         =   "Name"
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
         TabIndex        =   16
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame fraPolicyDetails 
      Height          =   3495
      Left            =   5520
      TabIndex        =   2
      Top             =   1200
      Width           =   4815
      Begin VB.ComboBox cboAgentNumber 
         DataSource      =   "dtaAgentNumber"
         Height          =   315
         Left            =   2040
         TabIndex        =   32
         Top             =   2520
         Width           =   2415
      End
      Begin VB.ComboBox cboGender 
         Height          =   315
         ItemData        =   "frmPolicyCreationForm.frx":08B9
         Left            =   2040
         List            =   "frmPolicyCreationForm.frx":08C3
         TabIndex        =   29
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox txtResidence 
         Height          =   375
         Left            =   2040
         TabIndex        =   27
         Top             =   1800
         Width           =   2415
      End
      Begin VB.ComboBox cboPolicyType 
         Height          =   315
         ItemData        =   "frmPolicyCreationForm.frx":08D5
         Left            =   2040
         List            =   "frmPolicyCreationForm.frx":08D7
         TabIndex        =   24
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label lblGender 
         Alignment       =   1  'Right Justify
         Caption         =   "Gender"
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
         TabIndex        =   28
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblResidence 
         Alignment       =   1  'Right Justify
         Caption         =   "Residence"
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
         TabIndex        =   26
         Top             =   1800
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
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   2520
         Width           =   1455
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
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.Frame fraPersonalDetails 
      Height          =   3495
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   4935
      Begin VB.TextBox txtEmail 
         Height          =   375
         Left            =   2040
         TabIndex        =   31
         Top             =   2760
         Width           =   2415
      End
      Begin VB.TextBox txtDateOfBirth 
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox txtMobileNumber 
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox txtIdNumber 
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox txtName 
         DataSource      =   "dtaPolicyDetails"
         Height          =   375
         Left            =   2040
         TabIndex        =   8
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lblEmail 
         Alignment       =   1  'Right Justify
         Caption         =   "Email"
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
         TabIndex        =   30
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label lblDateOfBirth 
         Alignment       =   1  'Right Justify
         Caption         =   "D.O.B"
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
         TabIndex        =   7
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label lblMobileNumber 
         Alignment       =   1  'Right Justify
         Caption         =   "Mobile Number"
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
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label lblIdNumber 
         Alignment       =   1  'Right Justify
         Caption         =   "Id Number"
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
         TabIndex        =   5
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Caption         =   "Name"
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
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   8280
      Picture         =   "frmPolicyCreationForm.frx":08D9
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblCreatePolicy 
      Alignment       =   2  'Center
      Caption         =   "CreatePolicy"
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
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   6735
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   7695
   End
End
Attribute VB_Name = "frmPolicyCreationForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check_Dateinput(inp As Control)
    Dim strlen As Integer
    strlen = Len(inp.Text)
    If strlen > 4 Then
        MsgBox "Date Should Not Exceed 4", vbCritical
        inp.Text = ""
        inp.SetFocus
    ElseIf strlen <= 3 Then
        MsgBox "Date Should not be Below 4 Characters", vbCritical
        inp.Text = ""
        inp.SetFocus
    End If
    If Not IsNumeric(inp.Text) Then
        MsgBox "Please input Numbers", vbCritical
        inp.Text = ""
        inp.SetFocus
    End If
End Sub
Private Sub AgentCombo()
    While dtaAgentNumber.Recordset.EOF = False
        cboAgentNumber.AddItem dtaAgentNumber.Recordset.Fields(6).Value
        dtaAgentNumber.Recordset.MoveNext
    Wend
End Sub
Private Sub Check_Email(inp As Control)
    Dim email As String
    email = inp.Text
    If InStr(1, email, "@") = 0 Then
        MsgBox "Key In Valid Email"
        inp.Text = ""
        inp.SetFocus
    End If
End Sub
Private Sub PopulateCombo()
    While dtaTypeOfAccount.Recordset.EOF = False
        cboPolicyType.AddItem dtaTypeOfAccount.Recordset.Fields(2).Value
        dtaTypeOfAccount.Recordset.MoveNext
    Wend
End Sub
Private Sub Check_input(inp As Control)
    Dim strlen As Integer
    strlen = Len(inp.Text)
    If strlen > 10 Then
        MsgBox "Number Should Not Exceed 10", vbCritical
        inp.Text = ""
        inp.SetFocus
    ElseIf strlen <= 6 Then
        MsgBox "Number Should not be Below Six Characters", vbCritical
        inp.Text = ""
        inp.SetFocus
    End If
    If Not IsNumeric(inp.Text) Then
        MsgBox "Please input Numbers", vbCritical
        inp.Text = ""
        inp.SetFocus
    End If
End Sub
Private Sub cmdCancel_Click()
    If cboAgentNumber.Text = "" And txtDateOfBirth.Text = "" And txtIdNumber.Text = "" And txtMobileNumber.Text = "" And txtName.Text = "" And txtNextOfKinName.Text = "" And txtNextOfKinPhoneNo.Text = "" And txtNextOfKinResidence.Text = "" And txtNextOfKinResidence.Text = "" And txtResidence.Text = "" And cboNextOfKinRelation.Text = "" And cboPolicyType.Text = "" And cboGender.Text = "" Then
        Unload Me
        frmMain.Show
    Else
        txtAgentNumber.Text = ""
        txtDateOfBirth.Text = ""
        txtIdNumber.Text = ""
        txtMobileNumber.Text = ""
        txtName.Text = ""
        txtNextOfKinName.Text = ""
        txtMobileNumber.Text = ""
        txtNextOfKinPhoneNo.Text = ""
        txtNextOfKinResidence.Text = ""
        txtResidence.Text = ""
        cboGender.Text = ""
        cboNextOfKinRelation.Text = ""
        cboPolicyType.Text = ""
    End If
End Sub
Private Sub cmdProceed_Click()
    Dim lngPhoneNumber As Long
    Dim lngIdNumber As Long
    Dim lngAgentNumber As Long, DateOfBirth As String
    Dim lngKinPolicyNumber As Long
    Dim lngkinPhoneNumber As Long
    Dim strTypeOfPolicy As String
    
    'Checking if all Inputs are entered
    If cboAgentNumber.Text = "" Or txtDateOfBirth.Text = "" Or txtIdNumber.Text = "" Or txtMobileNumber.Text = "" Or txtName.Text = "" Or txtNextOfKinName.Text = "" Or txtNextOfKinPhoneNo.Text = "" Or txtNextOfKinResidence.Text = "" Or txtNextOfKinResidence.Text = "" Or txtResidence.Text = "" Or cboNextOfKinRelation.Text = "" Or cboPolicyType.Text = "" Or cboGender.Text = "" Then
        MsgBox "Please Fill In All Inputs", vbCritical
    Else
        'Saving details of policy applicant
        lngPhoneNumber = txtMobileNumber.Text
        lngIdNumber = txtIdNumber.Text
        lngAgentNumber = cboAgentNumber.Text
        strTypeOfPolicy = cboPolicyType.Text
        
        dtaPolicyDetails.Recordset.AddNew
        dtaPolicyDetails.Recordset.Fields(1).Value = txtName.Text
        dtaPolicyDetails.Recordset.Fields(2).Value = lngIdNumber
        dtaPolicyDetails.Recordset.Fields(3).Value = lngPhoneNumber
        dtaPolicyDetails.Recordset.Fields(4).Value = txtResidence.Text
        dtaPolicyDetails.Recordset.Fields(5).Value = cboPolicyType.Text
        dtaPolicyDetails.Recordset.Fields(6).Value = cboGender.Text
        dtaPolicyDetails.Recordset.Fields(7).Value = lngAgentNumber
        dtaPolicyDetails.Recordset.Fields(8).Value = Format(Now, "mm/dd/yy hh:mm:ss")
        dtaPolicyDetails.Recordset.Fields(9).Value = txtDateOfBirth.Text
        dtaPolicyDetails.Recordset.Fields(10).Value = txtEmail.Text
        dtaPolicyDetails.Recordset.Update
        
        'Automating Inputs on the next Forms
        dtaPolicyDetails.Recordset.MoveLast
        lngPolicyNumber = dtaPolicyDetails.Recordset.Fields(0).Value
        DateOfBirth = dtaPolicyDetails.Recordset.Fields(9).Value
        lngAge = 2021 - Val(DateOfBirth)
        
        'saving details of next of kin
        lngkinPhoneNumber = txtNextOfKinPhoneNo
        dtaNextOfKin.Recordset.AddNew
        dtaNextOfKin.Recordset.Fields(0).Value = lngPolicyNumber
        dtaNextOfKin.Recordset.Fields(1).Value = txtNextOfKinName.Text
        dtaNextOfKin.Recordset.Fields(2).Value = lngkinPhoneNumber
        dtaNextOfKin.Recordset.Fields(3).Value = cboNextOfKinRelation.Text
        dtaNextOfKin.Recordset.Fields(4).Value = txtNextOfKinResidence.Text
        dtaNextOfKin.Recordset.Update
        
        MsgBox "Record Updated Succesfully", vbInformation
        'Clearing Inputs
        cboAgentNumber.Text = ""
        txtDateOfBirth.Text = ""
        txtIdNumber.Text = ""
        txtMobileNumber.Text = ""
        txtName.Text = ""
        txtNextOfKinName.Text = ""
        txtNextOfKinPhoneNo.Text = ""
        txtNextOfKinResidence.Text = ""
        txtResidence.Text = ""
        cboGender.Text = ""
        cboNextOfKinRelation.Text = ""
        cboPolicyType.Text = ""
        
        'Choosing Page
        If strTypeOfPolicy = "Fire and Perils Insurance" Then
            Unload Me
            Unload frmMain
            frmFirePerils.Show
        ElseIf strTypeOfPolicy = "Motor Insurance" Then
            Unload Me
            Unload frmMain
            frmMotorInsurance.Show
        ElseIf strTypeOfPolicy = "Accident Insurance" Then
            Unload Me
            Unload frmMain
            frmAccidentInsurance.Show
        Else
            Unload Me
            Unload frmMain
            frmPolicyPayment.Show
        End If
    End If
End Sub
Private Sub Form_Load()
    'Positioning the form
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
    Call PopulateCombo
    Call AgentCombo
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
    frmMain.Show
End Sub
Private Sub txtDateOfBirth_Validate(Cancel As Boolean)
    Call Check_Dateinput(txtDateOfBirth)
End Sub

Private Sub txtEmail_Validate(Cancel As Boolean)
    Call Check_Email(txtEmail)
End Sub
Private Sub txtIdNumber_Validate(Cancel As Boolean)
    Call Check_input(txtIdNumber)
End Sub
Private Sub txtMobileNumber_Validate(Cancel As Boolean)
    Call Check_input(txtMobileNumber)
End Sub
Private Sub txtNextOfKinPhoneNo_Validate(Cancel As Boolean)
    Call Check_input(txtNextOfKinPhoneNo)
End Sub
