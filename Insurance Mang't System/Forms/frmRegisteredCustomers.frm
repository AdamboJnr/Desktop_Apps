VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmRegisteredCustomers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registered Customers"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6600
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Left            =   2760
      Picture         =   "frmRegisteredCustomers.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3600
      Width           =   975
   End
   Begin MSAdodcLib.Adodc dtaPolicyDetails 
      Height          =   330
      Left            =   240
      Top             =   3840
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
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
      Left            =   5280
      Picture         =   "frmRegisteredCustomers.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdAllCustomers 
      Caption         =   "Policies"
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
      Left            =   4080
      Picture         =   "frmRegisteredCustomers.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3600
      Width           =   975
   End
   Begin VB.Frame fraRegisteredCustomersDetails 
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   6135
      Begin VB.TextBox txtPolicyType 
         Height          =   375
         Left            =   3000
         TabIndex        =   10
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox txtHolderName 
         Height          =   375
         Left            =   3000
         TabIndex        =   8
         Top             =   960
         Width           =   2175
      End
      Begin VB.ComboBox cboPolicyNumber 
         DataSource      =   "dtaPolicyDetails"
         Height          =   315
         Left            =   3000
         TabIndex        =   3
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label lblTypeOfPolicy 
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
         Left            =   840
         TabIndex        =   9
         Top             =   1560
         Width           =   1695
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
         Left            =   840
         TabIndex        =   7
         Top             =   960
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
         Left            =   840
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   4800
      Picture         =   "frmRegisteredCustomers.frx":0CC6
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblViewRegisteredCustomers 
      Alignment       =   2  'Center
      Caption         =   "Registered Customers"
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
      Width           =   3855
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "frmRegisteredCustomers"
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
    Dim PolicyNumber As Long
    
    PolicyNumber = cboPolicyNumber.Text
    'Automatically filling in the Holder Name
    dtaPolicyDetails.Recordset.MoveFirst
    dtaPolicyDetails.Recordset.Find "[Policy Number]= " & PolicyNumber, 0, adSearchForward
    If dtaPolicyDetails.Recordset.EOF = True Then
        MsgBox "Record Not Found"
        dtaPolicyDetails.Recordset.MoveFirst
    ElseIf dtaPolicyDetails.Recordset.Fields(0).Value = PolicyNumber Then
        txtHolderName.Text = dtaPolicyDetails.Recordset.Fields(1).Value
        txtPolicyType.Text = dtaPolicyDetails.Recordset.Fields(2).Value
    End If
End Sub
Private Sub cmdAllCustomers_Click()
    rptPolicyHolders.Show
End Sub
Private Sub cmdDetails_Click()
    Dim PolicyNumber As Long
    Dim PolicyType As String
    If cboPolicyNumber.Text = "" Then
        MsgBox "Please Select a Valid Policy Number", vbInformation
    Else
        PolicyType = txtPolicyType.Text
        If PolicyType = "Fire and Perils Insurance" Then
            'Closing any opened reports
            If denFireAndPerils.rsFireAndPerils.State Then
                denFireAndPerils.rsFireAndPerils.Close
            End If
            
            PolicyNumber = cboPolicyNumber.Text
            denFireAndPerils.FireAndPerils PolicyNumber
            rptFireAndPerilsAcceptReject.Show
        ElseIf PolicyType = "Motor Insurance" Then
            'Closing any opened reports
            If denMotorInsurance.rsMotorInsurance.State Then
                denMotorInsurance.rsMotorInsurance.Close
            End If
            
            PolicyNumber = cboPolicyNumber.Text
            denMotorInsurance.MotorInsurance PolicyNumber
            rptMotorInsuranceAcceptReject.Show
        ElseIf PolicyType = "Accident Insurance" Then
            'Closing any opened reports
            If denAccidentInsurance.rsAccidentInsurance.State Then
                denAccidentInsurance.rsAccidentInsurance.Close
            End If
            
            PolicyNumber = cboPolicyNumber.Text
            denAccidentInsurance.AccidentInsurance PolicyNumber
            rptAccidentInsuranceAcceptReject.Show
        End If
    End If
End Sub

Private Sub cmdExit_Click()
    If cboPolicyNumber.Text = "" Or txtHolderName.Text = "" Or txtPolicyType.Text = "" Then
        Unload Me
        frmAdminDashboard.Show
    Else
        cboPolicyNumber.Text = ""
        txtHolderName.Text = ""
        txtPolicyType.Text = ""
    End If
End Sub

Private Sub Form_Load()
    'Positioning the form
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
    Call PopulateCombo
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmAdminDashboard.Show
End Sub
