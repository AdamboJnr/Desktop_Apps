VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPresidentialReportForm 
   Caption         =   "Presidential Report"
   ClientHeight    =   5535
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   5445
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc dtaVotingPositions 
      Height          =   375
      Left            =   120
      Top             =   5040
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "Votes"
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
   Begin MSAdodcLib.Adodc dtaPresidency 
      Height          =   375
      Left            =   120
      Top             =   4440
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      RecordSource    =   "select * from tblPresident"
      Caption         =   "Presidency"
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
   Begin VB.ListBox lstVoters 
      DataSource      =   "dtaVotingPositions"
      Height          =   1035
      Left            =   2280
      TabIndex        =   7
      Top             =   2040
      Width           =   2775
   End
   Begin VB.TextBox txtAspirantTotal 
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   3480
      Width           =   2775
   End
   Begin VB.CommandButton cmdTotalVotes 
      Caption         =   "Total Votes"
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
      Left            =   3960
      TabIndex        =   3
      Top             =   4200
      Width           =   1095
   End
   Begin VB.ComboBox cboAspirant 
      DataSource      =   "dtaPresidency"
      Height          =   315
      Left            =   2280
      TabIndex        =   2
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label lblVoters 
      Alignment       =   1  'Right Justify
      Caption         =   "Voters"
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
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label lblAspirantTotal 
      Alignment       =   1  'Right Justify
      Caption         =   "Aspirant Total Votes"
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
      TabIndex        =   5
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label lblAspirant 
      Alignment       =   1  'Right Justify
      Caption         =   "Aspirant"
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
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   3960
      Picture         =   "frmPresidentialReportForm.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      Caption         =   "President Election Report"
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
      Width           =   3135
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "frmPresidentialReportForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub PopulateCombo()
    While dtaPresidency.Recordset.EOF = False
        cboAspirant.AddItem dtaPresidency.Recordset.Fields(1).Value
        dtaPresidency.Recordset.MoveNext
    Wend
End Sub
Private Sub cboAspirant_Click()
    Dim searchvalue As String
    searchvalue = cboAspirant.Text
    dtaVotingPositions.Recordset.MoveFirst
    While dtaVotingPositions.Recordset.EOF = False
        If dtaVotingPositions.Recordset.Fields(0).Value = searchvalue Then
            lstVoters.AddItem dtaVotingPositions.Recordset.Fields(4).Value
        End If
        dtaVotingPositions.Recordset.MoveNext
    Wend
    txtAspirantTotal.Text = lstVoters.ListCount
End Sub

Private Sub cmdTotalVotes_Click()
    rptPresidency.Show
End Sub

Private Sub Form_Load()
    'Positioning the Form
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
    Call PopulateCombo
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmMain.Show
End Sub
