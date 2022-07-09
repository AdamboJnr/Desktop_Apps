VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCreateUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create User"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5715
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc dtaEmployeeNumber 
      Height          =   495
      Left            =   2400
      Top             =   5040
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
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
      Connect         =   $"frmCreateUser.frx":0000
      OLEDBString     =   $"frmCreateUser.frx":008F
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from tblEmployee"
      Caption         =   "Employee Number"
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
   Begin VB.ComboBox cboEmployeeNumber 
      DataSource      =   "dtaEmployeeNumber"
      Height          =   315
      Left            =   2760
      TabIndex        =   8
      Top             =   3000
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc dtaCreateAccount 
      Height          =   495
      Left            =   240
      Top             =   5040
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
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
      RecordSource    =   "select * from tblLogIn"
      Caption         =   "Create Account"
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
   Begin VB.CheckBox chkAdmin 
      Caption         =   "Admin Priviledges"
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
      Left            =   1560
      TabIndex        =   6
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00000000&
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
      Height          =   615
      Left            =   3720
      TabIndex        =   5
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
      DataSource      =   "dtaCreateAccount"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox txtUsername 
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label lblEmployeeNumber 
      Alignment       =   1  'Right Justify
      Caption         =   "Employee Number"
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
      Left            =   480
      TabIndex        =   7
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   3840
      Picture         =   "frmCreateUser.frx":011E
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   840
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label lblPassword 
      Alignment       =   1  'Right Justify
      Caption         =   "Password"
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
      Left            =   480
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label lblUsername 
      Alignment       =   1  'Right Justify
      Caption         =   "Username"
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
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lblCreateUser 
      Alignment       =   2  'Center
      Caption         =   "Create User"
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
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
End
Attribute VB_Name = "frmCreateUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSave_Click()
    Dim strTypeOfUser As String
    If txtPassword.Text = "" Or txtPassword.Text = "" Or cboEmployeeNumber.Text = "" Then
        MsgBox "Please Fill in all Inputs"
    Else
        lngEmployeeNumber = cboEmployeeNumber.Text
        'Checking for Admin Priviledges
        If chkAdmin.Value = Checked Then
            strTypeOfUser = "Admin"
            dtaCreateAccount.Recordset.AddNew
            dtaCreateAccount.Recordset.Fields(1).Value = txtUsername.Text
            dtaCreateAccount.Recordset.Fields(2).Value = txtPassword.Text
            dtaCreateAccount.Recordset.Fields(3).Value = strTypeOfUser
            dtaCreateAccount.Recordset.Fields(4).Value = lngEmployeeNumber
            dtaCreateAccount.Recordset.Update
        Else
            strTypeOfUser = "User"
            dtaCreateAccount.Recordset.AddNew
            dtaCreateAccount.Recordset.Fields(1).Value = txtUsername.Text
            dtaCreateAccount.Recordset.Fields(2).Value = txtPassword.Text
            dtaCreateAccount.Recordset.Fields(3).Value = strTypeOfUser
            dtaCreateAccount.Recordset.Fields(4).Value = lngEmployeeNumber
            dtaCreateAccount.Recordset.Update
        End If
        txtUsername.Text = ""
        txtPassword.Text = ""
        chkAdmin.Value = Unchecked
    End If
End Sub

Private Sub Form_Load()
    'Positioning the Form
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
    'Loading Employee Numbers
    While dtaEmployeeNumber.Recordset.EOF = False
        cboEmployeeNumber.AddItem dtaEmployeeNumber.Recordset.Fields(0).Value
        dtaEmployeeNumber.Recordset.MoveNext
    Wend
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmAdminDashboard.Show
End Sub
