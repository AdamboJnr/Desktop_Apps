VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPolicyClaimModule 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Policy Claim Module"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6255
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc dtaPolicyClaim 
      Height          =   375
      Left            =   480
      Top             =   4560
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
      RecordSource    =   "select * from tblPolicyClaims"
      Caption         =   "PolicyClaim"
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
   Begin MSAdodcLib.Adodc dtaAcceptedPolicies 
      Height          =   375
      Left            =   480
      Top             =   4080
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
      Caption         =   "Policies"
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
      Left            =   5040
      Picture         =   "Form2a.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4080
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
      Left            =   3720
      Picture         =   "Form2a.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4080
      Width           =   975
   End
   Begin VB.Frame fraPolicyClaimDetaiils 
      Height          =   2655
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   5775
      Begin VB.TextBox txtClaimType 
         DataSource      =   "dtaPolicyClaim"
         Height          =   375
         Left            =   2280
         TabIndex        =   9
         Top             =   1800
         Width           =   2535
      End
      Begin VB.ComboBox cboPolicyNumber 
         DataSource      =   "dtaAcceptedPolicies"
         Height          =   315
         Left            =   2280
         TabIndex        =   8
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox txtClaimantName 
         Height          =   375
         Left            =   2280
         TabIndex        =   7
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label lblClaimantName 
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
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblClaimantClaim 
         Alignment       =   1  'Right Justify
         Caption         =   "Claim Type"
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
         TabIndex        =   3
         Top             =   1800
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
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   4800
      Picture         =   "Form2a.frx":0884
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblPolicyClaimModuleCaption 
      Alignment       =   2  'Center
      Caption         =   "Policy Claim"
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
      Width           =   3975
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "frmPolicyClaimModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub PopulateCombo()
    While dtaAcceptedPolicies.Recordset.EOF = False
        If dtaAcceptedPolicies.Recordset.Fields(4).Value = "Accepted" Then
            cboPolicyNumber.AddItem dtaAcceptedPolicies.Recordset.Fields(0).Value
        End If
        dtaAcceptedPolicies.Recordset.MoveNext
    Wend
End Sub
Private Sub cboPolicyNumber_Click()
    Dim lngTempPolicyNumber As Long
    
    lngTempPolicyNumber = cboPolicyNumber.Text
    
    'Filling in type of insurance and user
    dtaAcceptedPolicies.Recordset.MoveFirst
    dtaAcceptedPolicies.Recordset.Find "[Policy Number]= " & lngTempPolicyNumber, 0, adSearchForward
    If dtaAcceptedPolicies.Recordset.EOF = True Then
        dtaAcceptedPolicies.Recordset.MoveFirst
    ElseIf dtaAcceptedPolicies.Recordset.Fields(0).Value = lngTempPolicyNumber Then
        txtClaimantName.Text = dtaAcceptedPolicies.Recordset.Fields(1).Value
        txtClaimType.Text = dtaAcceptedPolicies.Recordset.Fields(2).Value
    End If
End Sub

Private Sub cmdCancel_Click()
    If txtClaimantName.Text = "" And txtClaimType.Text = "" And cboPolicyNumber.Text Then
        Unload Me
        frmMain.Show
    Else
        txtClaimantName.Text = ""
        txtClaimType.Text = ""
        cboPolicyNumber.Text = ""
    End If
End Sub
Private Sub cmdProceed_Click()
    Dim ClaimType As String
    Dim PolicyNumber As Long
    
    ClaimType = txtClaimType.Text
    
    If txtClaimantName.Text = "" Or txtClaimType.Text = "" Or cboPolicyNumber.Text = "" Then
        MsgBox "Please Fill In All Inputs", vbCritical
    Else
        PolicyNumber = cboPolicyNumber.Text
        'Saving to Database
        dtaPolicyClaim.Recordset.AddNew
        dtaPolicyClaim.Recordset.Fields(0).Value = PolicyNumber
        dtaPolicyClaim.Recordset.Fields(1).Value = txtClaimantName.Text
        dtaPolicyClaim.Recordset.Fields(2).Value = txtClaimType.Text
        dtaPolicyClaim.Recordset.Update
        
        ClaimantPolicyNumber = cboPolicyNumber.Text
        ClaimantName = txtClaimantName.Text
        
        txtClaimantName.Text = ""
        txtClaimType.Text = ""
        cboPolicyNumber.Text = ""
        
        'Saving Claim Number
        dtaPolicyClaim.Recordset.MoveLast
        ClaimNumber = dtaPolicyClaim.Recordset.Fields(3).Value
        
        If ClaimType = "Motor Insurance" Then
            Unload Me
            frmMotorClaim.Show
        ElseIf ClaimType = "Fire and Perils Insurance" Then
            Unload Me
            frmFireAndPerilsClaim.Show
        ElseIf ClaimType = "Accident Insurance" Then
            Unload Me
            frmAccidentClaim.Show
        End If
    End If
End Sub

Private Sub Form_Load()
    'Positioning the form
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
    Call PopulateCombo
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
    frmMain.Show
End Sub
