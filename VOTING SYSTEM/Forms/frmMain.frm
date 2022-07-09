VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmMain 
   Caption         =   "Home"
   ClientHeight    =   5895
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      DataSource      =   "dtaAdmissionNumber"
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   3600
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSAdodcLib.Adodc dtaAdmissionNumber 
      Height          =   330
      Left            =   480
      Top             =   5520
      Width           =   1695
      _ExtentX        =   2990
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Voting System.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Voting System.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from tblStudentsDetails"
      Caption         =   "Admission Number"
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
   Begin VB.TextBox Text1 
      DataSource      =   "dtaVotingStatus"
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3360
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSAdodcLib.Adodc dtaVotingStatus 
      Height          =   375
      Left            =   3120
      Top             =   5520
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Voting System.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Voting System.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from tblVotingPositions"
      Caption         =   "Check Status"
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
   Begin VB.CommandButton cmdLogOut 
      Caption         =   "Log Out"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4320
      Picture         =   "frmMain.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton cmdVote 
      Caption         =   "Vote"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2880
      Picture         =   "frmMain.frx":4888A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton cmdRegisterAspirants 
      Caption         =   "Register Aspirants"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1440
      Picture         =   "frmMain.frx":490A1
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton cmdRegisterStudent 
      Caption         =   "Register Student"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      Picture         =   "frmMain.frx":4A123
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   4935
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      Caption         =   "KCNP VOTING SYSTEM"
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
      TabIndex        =   4
      Top             =   4800
      Width           =   4695
   End
   Begin VB.Image Image1 
      Height          =   2655
      Left            =   1440
      Picture         =   "frmMain.frx":4A84A
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Menu mnuStudents 
      Caption         =   "Students"
      Begin VB.Menu mnuRegisterStudent 
         Caption         =   "Register Student"
      End
   End
   Begin VB.Menu mnuAspirants 
      Caption         =   "Aspirants"
      Begin VB.Menu mnuRegisterAspirant 
         Caption         =   "Register Aspirant"
      End
   End
   Begin VB.Menu mnuVote 
      Caption         =   "Vote"
      Begin VB.Menu mnuCastVote 
         Caption         =   "Cast Vote"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Reports"
      Begin VB.Menu mnuPresidentElection 
         Caption         =   "President Report"
      End
      Begin VB.Menu mnuDeputyPresident 
         Caption         =   "Deputy President Report"
      End
      Begin VB.Menu mnuSecGen 
         Caption         =   "Secretary General Report"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLogOut_Click()
    Unload Me
    frmLogIn.Show
End Sub

Private Sub cmdRegisterAspirants_Click()
    Unload Me
    frmAspirantRegistration.Show
End Sub

Private Sub cmdRegisterStudent_Click()
    Unload Me
    frmStudentRegistration.Show
End Sub
Private Sub cmdVote_Click()
    Dim lngSearchValue As Long
    lngAdmissionNumber = Val(InputBox("Please Input Student Admission Number"))
        dtaVotingStatus.Recordset.MoveFirst
        dtaVotingStatus.Recordset.Find "[Admission Number]= " & lngAdmissionNumber, 0, adSearchForward
        If dtaVotingStatus.Recordset.EOF = True Then
            dtaAdmissionNumber.Recordset.MoveFirst
            dtaAdmissionNumber.Recordset.Find "[Admission Number]= " & lngAdmissionNumber, 0, adSearchForward
            If dtaAdmissionNumber.Recordset.EOF = True Then
                MsgBox "Student Record not found. Please Register Student"
            ElseIf dtaAdmissionNumber.Recordset.Fields(0).Value = lngAdmissionNumber Then
                frmCastVote.Show
            End If
        ElseIf dtaVotingStatus.Recordset.Fields(4).Value = lngAdmissionNumber Then
            MsgBox "User Already voted..!!!", vbCritical
            frmMain.Show
        End If
End Sub
Private Sub Form_Load()
    'Positioning the Form
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        frmLogIn.Show
End Sub

Private Sub mnuCastVote_Click()
    Dim lngSearchValue As Long
    lngAdmissionNumber = Val(InputBox("Please Input Student Admission Number"))
    dtaVotingStatus.Recordset.MoveFirst
    dtaVotingStatus.Recordset.Find "[Admission Number]= " & lngAdmissionNumber, 0, adSearchForward
    If dtaVotingStatus.Recordset.EOF = True Then
        dtaAdmissionNumber.Recordset.MoveFirst
        dtaAdmissionNumber.Recordset.Find "[Admission Number]= " & lngAdmissionNumber, 0, adSearchForward
        If dtaAdmissionNumber.Recordset.EOF = True Then
            MsgBox "Student Record not found. Please Register Student"
        ElseIf dtaAdmissionNumber.Recordset.Fields(0).Value = lngAdmissionNumber Then
            frmCastVote.Show
        End If
    ElseIf dtaVotingStatus.Recordset.Fields(4).Value = lngAdmissionNumber Then
        MsgBox "User Already voted..!!!", vbCritical
        frmMain.Show
    End If
End Sub

Private Sub mnuDeputyPresident_Click()
    Unload Me
    frmDepPresidentElection.Show
End Sub
Private Sub mnuPresidentElection_Click()
    Unload Me
    frmPresidentialReportForm.Show
End Sub

Private Sub mnuRegisterAspirant_Click()
    Unload Me
    frmAspirantRegistration.Show
End Sub

Private Sub mnuRegisterStudent_Click()
    Unload Me
    frmStudentRegistration.Show
End Sub

Private Sub mnuSecGen_Click()
    Unload Me
    frmSecGeneralElection.Show
End Sub
