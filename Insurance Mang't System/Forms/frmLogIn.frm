VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLogIn 
   Caption         =   "Log In Form"
   ClientHeight    =   5010
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6585
   LinkTopic       =   "Form2"
   ScaleHeight     =   5010
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc dtaUserLogs 
      Height          =   375
      Left            =   240
      Top             =   4200
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
      RecordSource    =   "SELECT *FROM tblUserLogs"
      Caption         =   "User Logs"
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
   Begin MSAdodcLib.Adodc dtaLogin 
      Height          =   495
      Left            =   240
      Top             =   3360
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Insurance management database.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Insurance management database.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM tblLogin"
      Caption         =   "Login"
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
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   6
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdLogIn 
      Caption         =   "Log In"
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
      Left            =   3000
      TabIndex        =   5
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2760
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2520
      Width           =   2895
   End
   Begin VB.TextBox txtUserName 
      DataField       =   "UserName"
      DataSource      =   "dtaUserLogs"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   1800
      Picture         =   "frmLogIn.frx":0000
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   435
      Left            =   1800
      Picture         =   "frmLogIn.frx":07A1
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   480
   End
   Begin VB.Shape Shape3 
      Height          =   615
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label lblUserName 
      Caption         =   "UserName"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   4560
      Picture         =   "frmLogIn.frx":0EC8
      Stretch         =   -1  'True
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lblFormCaption 
      Alignment       =   1  'Right Justify
      Caption         =   "Account Log In"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
   Begin VB.Shape Shape2 
      Height          =   855
      Left            =   1440
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   3975
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   2175
   End
End
Attribute VB_Name = "frmLogIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLogIn_Click()
'Dim strSearch As String
Dim strUserName As String
dtaLogin.RecordSource = "select * from tblLogin where Username= '" + txtUserName.Text + "' and Password = '" + txtPassword.Text + "'"
dtaLogin.Refresh

strUserName = txtUserName.Text

If dtaLogin.Recordset.EOF Then
    MsgBox "Sign In Failed, Try Again.!!", vbCritical
Else
    frmLogIn.Hide
    strUser = dtaLogin.Recordset.Fields(0).Value
    strTypeUser = dtaLogin.Recordset.Fields(2).Value
    
    'Saving User Activity
    dtaUserLogs.Recordset.AddNew
    dtaUserLogs.Recordset.Fields(0).Value = strUserName
    dtaUserLogs.Recordset.Fields(1).Value = Format(Now, "mm/dd/yy hh:mm:ss")
    dtaUserLogs.Recordset.Fields(3).Value = "N/A"
    dtaUserLogs.Recordset.Update
    
    dtaUserLogs.Recordset.MoveLast
    LogId = dtaUserLogs.Recordset.Fields(2).Value
    frmMain.Show
End If
End Sub
Private Sub Form_Load()
    txtUserName.Text = ""
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim logout As Date
    logout = Format(Now, "mm/dd/yy hh:mm:ss")
    strTimeLoggedOut = CStr(logout)
    
    dtaUserLogs.Recordset.MoveFirst
    dtaUserLogs.Recordset.Find "[Log Id]= " & LogId, 0, adSearchForward
    If dtaUserLogs.Recordset.EOF = True Then
        dtaUserLogs.Recordset.MoveFirst
    ElseIf dtaUserLogs.Recordset.Fields(2).Value = LogId Then
        dtaUserLogs.Recordset.Fields(3).Value = strTimeLoggedOut
        dtaUserLogs.Recordset.Update
    End If
    Unload Me
End Sub
