VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCastVote 
   Caption         =   "Cast Vote"
   ClientHeight    =   6780
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   7740
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc dtaSecGeneral 
      Height          =   375
      Left            =   2400
      Top             =   6240
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
      RecordSource    =   "select * from tblSecretaryGen"
      Caption         =   "Sec Gen"
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
   Begin MSAdodcLib.Adodc dtaDepPresident 
      Height          =   330
      Left            =   240
      Top             =   6240
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      RecordSource    =   "select * from tblDeputyPresident"
      Caption         =   "DepPresident"
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
      Left            =   240
      Top             =   5640
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Voting System.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Voting System.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from tblPresident"
      Caption         =   "President"
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
   Begin MSAdodcLib.Adodc dtaVote 
      Height          =   375
      Left            =   2400
      Top             =   5640
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
      Caption         =   "Vote"
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
      Height          =   615
      Left            =   6120
      TabIndex        =   11
      Top             =   5400
      Width           =   1095
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
      Height          =   615
      Left            =   4680
      TabIndex        =   10
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Frame fraPresident 
      Height          =   1095
      Left            =   480
      TabIndex        =   7
      Top             =   4080
      Width           =   6735
      Begin VB.ComboBox cboPresident 
         DataSource      =   "dtaPresidency"
         Height          =   315
         Left            =   3120
         TabIndex        =   9
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label lblPresident 
         Alignment       =   1  'Right Justify
         Caption         =   "President"
         DataSource      =   "dtaVote"
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
         TabIndex        =   8
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame fraDeputyPresident 
      Height          =   1095
      Left            =   480
      TabIndex        =   4
      Top             =   2640
      Width           =   6735
      Begin VB.ComboBox cboDeputyPresident 
         DataSource      =   "dtaDepPresident"
         Height          =   315
         Left            =   3120
         TabIndex        =   6
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label lblDeputyPresident 
         Alignment       =   1  'Right Justify
         Caption         =   "Deputy President"
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
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame fraSecretaryGeneral 
      Height          =   1095
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   6735
      Begin VB.ComboBox cboSecretaryGeneral 
         DataSource      =   "dtaSecGeneral"
         Height          =   315
         Left            =   3120
         TabIndex        =   3
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label lblSecretaryGeneral 
         Alignment       =   1  'Right Justify
         Caption         =   "Secretary General"
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
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   5520
      Picture         =   "frmCastVote.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   615
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   1080
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label lblCastVote 
      Alignment       =   2  'Center
      Caption         =   "Cast Vote(s)"
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
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmCastVote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub PopulateCombobox()
        'president
        While dtaPresidency.Recordset.EOF = False
            cboPresident.AddItem dtaPresidency.Recordset.Fields(1).Value
            dtaPresidency.Recordset.MoveNext
        Wend
        'deputy president
        While dtaDepPresident.Recordset.EOF = False
            cboDeputyPresident.AddItem dtaDepPresident.Recordset.Fields(1).Value
            dtaDepPresident.Recordset.MoveNext
        Wend
        'secretary general
        While dtaSecGeneral.Recordset.EOF = False
            cboSecretaryGeneral.AddItem dtaSecGeneral.Recordset.Fields(1).Value
            dtaSecGeneral.Recordset.MoveNext
        Wend
End Sub

Private Sub cmdCancel_Click()
    cboDeputyPresident.Text = ""
    cboPresident.Text = ""
    cboSecretaryGeneral.Text = ""
End Sub

Private Sub cmdVote_Click()
    dtaVote.Recordset.AddNew
    dtaVote.Recordset.Fields(0).Value = cboPresident.Text
    dtaVote.Recordset.Fields(1).Value = cboDeputyPresident.Text
    dtaVote.Recordset.Fields(2).Value = cboSecretaryGeneral.Text
    dtaVote.Recordset.Fields(4).Value = lngAdmissionNumber
    dtaVote.Recordset.Update
    MsgBox "Voted Succesfully"
    frmMain.Show
End Sub

Private Sub Form_Load()
    'Positioning the Form
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
    'Populating the combo boxes
    Call PopulateCombobox
End Sub
