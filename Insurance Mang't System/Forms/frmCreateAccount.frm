VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCreateAccount 
   Caption         =   "Create Account"
   ClientHeight    =   4785
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6120
   LinkTopic       =   "Form2"
   ScaleHeight     =   4785
   ScaleWidth      =   6120
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc dtaCreateAccount 
      Height          =   375
      Left            =   480
      Top             =   3960
      Width           =   2535
      _ExtentX        =   4471
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
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create"
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
      Left            =   4560
      TabIndex        =   7
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Frame fraAccountCreationDetails 
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   5775
      Begin VB.CheckBox chkAdminPrivileges 
         Caption         =   "Admin Privileges"
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
         Left            =   1920
         TabIndex        =   6
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txtPassword 
         DataSource      =   "dtaCreateAccount"
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2400
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txtUserName 
         DataSource      =   "dtaCreateAccount"
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   360
         Width           =   2295
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
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblUserName 
         Alignment       =   1  'Right Justify
         Caption         =   "UserName"
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
         Width           =   1815
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   4560
      Picture         =   "frmCreateAccount.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblCreateAccount 
      Alignment       =   2  'Center
      Caption         =   "Create Account"
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
      Width           =   3855
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "frmCreateAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCreate_Click()
    Dim strTypeOfUser As String
    'Checking for Admin Priviledges
    If chkAdminPrivileges.Value = Checked Then
        strTypeOfUser = "Admin"
        dtaCreateAccount.Recordset.AddNew
        dtaCreateAccount.Recordset.Fields(0).Value = txtUserName.Text
        dtaCreateAccount.Recordset.Fields(1).Value = txtPassword.Text
        dtaCreateAccount.Recordset.Fields(2).Value = strTypeOfUser
        dtaCreateAccount.Recordset.Update
    Else
        strTypeOfUser = "User"
        dtaCreateAccount.Recordset.AddNew
        dtaCreateAccount.Recordset.Fields(0).Value = txtUserName.Text
        dtaCreateAccount.Recordset.Fields(1).Value = txtPassword.Text
        dtaCreateAccount.Recordset.Fields(2).Value = strTypeOfUser
        dtaCreateAccount.Recordset.Update
    End If
    txtUserName.Text = ""
    txtPassword.Text = ""
    chkAdminPrivileges.Value = Unchecked
End Sub

Private Sub Form_Load()
    'Positioning the form
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
    'Clearing the inputs
    txtUserName.Text = ""
    txtPassword.Text = ""
    dtaCreateAccount.Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
    frmAdminDashboard.Show
End Sub
