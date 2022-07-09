VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Log in form"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc dtaUserLogs 
      Height          =   375
      Left            =   360
      Top             =   3480
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Ichaweri Hotel Management System.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Ichaweri Hotel Management System.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from tblUserLogs"
      Caption         =   "UserLogs"
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
   Begin MSAdodcLib.Adodc dtaLogIn 
      Height          =   375
      Left            =   240
      Top             =   3120
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Ichaweri Hotel Management System.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Ichaweri Hotel Management System.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from tblLogIn"
      Caption         =   "Log In"
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
      Caption         =   "cancel"
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
      Left            =   3840
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Log in"
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
      Left            =   2400
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
      DataSource      =   "dtaUserLogs"
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox txtUsername 
      DataSource      =   "dtaLogIn"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   1560
      Width           =   2655
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
      Left            =   240
      TabIndex        =   7
      Top             =   2400
      Width           =   1575
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
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   3600
      Picture         =   "login form.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Shape Shape1 
      Height          =   855
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label lblLogin 
      Alignment       =   2  'Center
      Caption         =   "Accout Log in"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   2
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   15
      Left            =   2880
      TabIndex        =   1
      Top             =   1440
      Width           =   735
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLogIn_Click()
    Dim strUserName As String
    dtaLogIn.RecordSource = " select * from tblLogIn where Username = '" + txtUsername.Text + "' and Password = '" + txtPassword.Text + "' "
    dtaLogIn.Refresh
    If txtUsername.Text = "" Or txtPassword.Text = "" Then
        MsgBox "Please Fill in All Inputs"
    Else
        strUserName = txtUsername.Text
        
        If dtaLogIn.Recordset.EOF = True Then
            MsgBox "Wrong Username Or Password...!", vbCritical
            txtUsername.SetFocus
        Else
            strUserType = dtaLogIn.Recordset.Fields(3).Value
            frmMain.cmdAdminDashboard.Visible = True
            'Saving User Activity
            dtaUserLogs.Recordset.AddNew
            dtaUserLogs.Recordset.Fields(1).Value = strUserName
            dtaUserLogs.Recordset.Fields(2).Value = Format(Now, "mm/dd/yy hh:mm:ss")
            dtaUserLogs.Recordset.Update
            Unload Me
            frmMain.Show
            lngEmployeeNumber = dtaLogIn.Recordset.Fields(4).Value
        End If
    End If
End Sub
Private Sub Form_Load()
    'Positioning the Form
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
    'Clearing the inputs
    txtUsername.Text = ""
    txtPassword.Text = ""
    dtaLogIn.Visible = False
End Sub

