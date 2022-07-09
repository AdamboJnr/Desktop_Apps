VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAgentRegistration 
   Caption         =   "Agent Registration"
   ClientHeight    =   6465
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11370
   LinkTopic       =   "Form2"
   ScaleHeight     =   6465
   ScaleWidth      =   11370
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc dtaAgentRegistration 
      Height          =   375
      Left            =   480
      Top             =   5400
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      RecordSource    =   "SELECT * FROM tblAgentDetails"
      Caption         =   "Agent Registration"
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
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9840
      Picture         =   "frmAgentRegistration.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8280
      Picture         =   "frmAgentRegistration.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Frame fraAgentResidentualDetails 
      Height          =   2775
      Left            =   5880
      TabIndex        =   10
      Top             =   1440
      Width           =   5055
      Begin VB.TextBox txtResidence 
         DataSource      =   "dtaAgentRegistration"
         Height          =   375
         Left            =   2280
         TabIndex        =   18
         Top             =   1080
         Width           =   2175
      End
      Begin VB.ComboBox cboCounty 
         Height          =   315
         ItemData        =   "frmAgentRegistration.frx":0884
         Left            =   2280
         List            =   "frmAgentRegistration.frx":0894
         TabIndex        =   14
         Text            =   "Mombasa"
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox txtEmail 
         Height          =   375
         Left            =   2280
         TabIndex        =   13
         Top             =   360
         Width           =   2175
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
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblCounty 
         Alignment       =   1  'Right Justify
         Caption         =   "County"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label lblEmail 
         Alignment       =   1  'Right Justify
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame fraPersonalDetails 
      Height          =   3615
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   5055
      Begin VB.ComboBox cboGender 
         Height          =   315
         ItemData        =   "frmAgentRegistration.frx":08B9
         Left            =   2280
         List            =   "frmAgentRegistration.frx":08C6
         TabIndex        =   9
         Top             =   3000
         Width           =   2175
      End
      Begin VB.TextBox txtPhoneNumber 
         Height          =   375
         Left            =   2280
         TabIndex        =   7
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox txtIdNumber 
         Height          =   375
         Left            =   2280
         TabIndex        =   6
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox txtFullName 
         Height          =   375
         Left            =   2280
         TabIndex        =   5
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblGender 
         Alignment       =   1  'Right Justify
         Caption         =   "Gender"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label lblPhoneNumber 
         Caption         =   "Phone Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label lblIdNumber 
         Alignment       =   1  'Right Justify
         Caption         =   "ID Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblFullName 
         Alignment       =   1  'Right Justify
         Caption         =   "FullName"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   8160
      Picture         =   "frmAgentRegistration.frx":08DF
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblAgentRegCaption 
      Alignment       =   2  'Center
      Caption         =   "Agent Registration Form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
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
      Width           =   6495
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   1440
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   7455
   End
End
Attribute VB_Name = "frmAgentRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check_Email(inp As Control)
    Dim email As String
    email = inp.Text
    If InStr(1, email, "@") = 0 Then
        MsgBox "Key In Valid Email"
        inp.Text = ""
        inp.SetFocus
    End If
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

Private Sub cmdDelete_Click()
    If txtEmail.Text = "" And txtFullName.Text = "" And txtIdNumber.Text = "" And txtPhoneNumber.Text = "" And txtResidence.Text = "" And cboCounty.Text = "" And cboGender.Text = "" Then
        Unload Me
        frmMain.Show
    Else
        txtEmail.Text = ""
        txtFullName.Text = ""
        txtIdNumber.Text = ""
        txtPhoneNumber.Text = ""
        txtResidence.Text = ""
        cboCounty.Text = ""
        cboGender.Text = ""
    End If
End Sub

Private Sub cmdSave_Click()
    If txtEmail.Text = "" Or txtFullName.Text = "" Or txtIdNumber.Text = "" Or txtPhoneNumber.Text = "" Or txtResidence.Text = "" Or cboCounty.Text = "" Or cboGender.Text = "" Then
        MsgBox "Please Fill In all Inputs"
    Else
        'save to database
         dtaAgentRegistration.Recordset.AddNew
         dtaAgentRegistration.Recordset.Fields(0).Value = txtIdNumber.Text
         dtaAgentRegistration.Recordset.Fields(1).Value = txtPhoneNumber.Text
         dtaAgentRegistration.Recordset.Fields(2).Value = cboGender.Text
         dtaAgentRegistration.Recordset.Fields(3).Value = txtEmail.Text
         dtaAgentRegistration.Recordset.Fields(4).Value = txtResidence.Text
         dtaAgentRegistration.Recordset.Fields(5).Value = cboCounty.Text
         dtaAgentRegistration.Recordset.Fields(7).Value = txtFullName.Text
         dtaAgentRegistration.Recordset.Update
         'Update User
         MsgBox "Record updated Succesfully", vbInformation
         'Clear the Inputs
         txtEmail.Text = ""
         txtFullName.Text = ""
         txtIdNumber.Text = ""
         txtPhoneNumber.Text = ""
         txtResidence.Text = ""
         txtFullName.SetFocus
    End If
End Sub
Private Sub Form_Load()
    'Positioning the form
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
    'Ensuring Inputs Are Empty
    txtEmail.Text = ""
    txtFullName.Text = ""
    txtIdNumber.Text = ""
    txtPhoneNumber.Text = ""
    txtResidence.Text = ""
    cboCounty.Text = ""
    cboGender.Text = ""
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
    frmMain.Show
End Sub
Private Sub txtEmail_Validate(Cancel As Boolean)
    Call Check_Email(txtEmail)
End Sub
Private Sub txtIdNumber_Validate(Cancel As Boolean)
    Call Check_input(txtIdNumber)
End Sub
Private Sub txtPhoneNumber_Validate(Cancel As Boolean)
    Call Check_input(txtPhoneNumber)
End Sub
