VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAspirantRegistration 
   Caption         =   "Aspirant Registration"
   ClientHeight    =   7500
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   7320
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc dtaDepPresident 
      Height          =   330
      Left            =   240
      Top             =   6960
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "Dep President"
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
   Begin MSAdodcLib.Adodc dtaSecGeneral 
      Height          =   375
      Left            =   2520
      Top             =   6960
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
      RecordSource    =   "select * from tblSecretaryGen"
      Caption         =   "Sec General"
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
      Top             =   6240
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
      Left            =   5640
      TabIndex        =   12
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
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
      Left            =   4080
      TabIndex        =   11
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Frame fraAspirantsDetails 
      Height          =   4815
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   6495
      Begin VB.ComboBox cboPositionVying 
         Height          =   315
         ItemData        =   "frmAspirantRegistration.frx":0000
         Left            =   2640
         List            =   "frmAspirantRegistration.frx":000D
         TabIndex        =   15
         Top             =   4080
         Width           =   3015
      End
      Begin VB.ComboBox cboDepartment 
         Height          =   315
         ItemData        =   "frmAspirantRegistration.frx":0041
         Left            =   2640
         List            =   "frmAspirantRegistration.frx":0057
         TabIndex        =   14
         Top             =   2520
         Width           =   3015
      End
      Begin VB.TextBox txtPhoneNumber 
         Height          =   405
         Left            =   2640
         TabIndex        =   10
         Top             =   3240
         Width           =   3015
      End
      Begin VB.TextBox txtClass 
         DataSource      =   "dtaSecGeneral"
         Height          =   375
         Left            =   2640
         TabIndex        =   9
         Top             =   1800
         Width           =   3015
      End
      Begin VB.TextBox txtAdmissionNumber 
         DataSource      =   "dtaDepPresident"
         Height          =   375
         Left            =   2640
         TabIndex        =   8
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox txtAspirantName 
         DataSource      =   "dtaPresidency"
         Height          =   375
         Left            =   2640
         TabIndex        =   7
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label lblPositionVying 
         Alignment       =   1  'Right Justify
         Caption         =   "Position Vying"
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
         TabIndex        =   13
         Top             =   4080
         Width           =   2055
      End
      Begin VB.Label lblPhoneNumber 
         Alignment       =   1  'Right Justify
         Caption         =   "Phone Number"
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
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label lblDepartment 
         Alignment       =   1  'Right Justify
         Caption         =   "Department"
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
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label lblClass 
         Alignment       =   1  'Right Justify
         Caption         =   "Class"
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
         TabIndex        =   4
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label lblAdmissionNumber 
         Alignment       =   1  'Right Justify
         Caption         =   "Admission Number"
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
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label lblAspirantName 
         Alignment       =   1  'Right Justify
         Caption         =   "Aspirant Name"
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
         Width           =   2055
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   5160
      Picture         =   "frmAspirantRegistration.frx":00C6
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   840
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   5055
   End
   Begin VB.Label lblAspirantRegistration 
      Alignment       =   2  'Center
      Caption         =   "Aspirants Registration"
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
      Width           =   3975
   End
End
Attribute VB_Name = "frmAspirantRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    'Clearing the Inputs
    txtAdmissionNumber.Text = ""
    txtAspirantName.Text = ""
    txtClass.Text = ""
    txtPhoneNumber.Text = ""
    cboDepartment.Text = ""
    cboPositionVying.Text = ""
End Sub

Private Sub cmdSave_Click()
    Dim lngAdmission As Long
    Dim lngPhoneNumber As Long
    If txtAdmissionNumber.Text = "" Or txtAspirantName.Text = "" Or txtClass.Text = "" Or txtPhoneNumber.Text = "" Or cboDepartment.Text = "" Or cboPositionVying.Text = "" Then
        MsgBox "Please Fill in All the Inputs"
    Else
        lngAdmission = txtAdmissionNumber.Text
        lngPhoneNumber = txtPhoneNumber.Text
        'Saving to the database
        If cboPositionVying.Text = "PRESIDENT" Then
            dtaPresidency.Recordset.AddNew
            dtaPresidency.Recordset.Fields(1).Value = txtAspirantName.Text
            dtaPresidency.Recordset.Fields(2).Value = txtClass.Text
            dtaPresidency.Recordset.Fields(3).Value = cboDepartment.Text
            dtaPresidency.Recordset.Fields(4).Value = lngPhoneNumber
            dtaPresidency.Recordset.Fields(5).Value = cboPositionVying.Text
            dtaPresidency.Recordset.Fields(6).Value = lngAdmission
            dtaPresidency.Recordset.Update
            MsgBox "Record Updated Succesfully"
        ElseIf cboPositionVying.Text = "DEPUTY PRESIDENT" Then
            dtaDepPresident.Recordset.AddNew
            dtaDepPresident.Recordset.Fields(1).Value = txtAspirantName.Text
            dtaDepPresident.Recordset.Fields(2).Value = lngAdmission
            dtaDepPresident.Recordset.Fields(3).Value = txtClass.Text
            dtaDepPresident.Recordset.Fields(4).Value = cboDepartment.Text
            dtaDepPresident.Recordset.Fields(5).Value = lngPhoneNumber
            dtaDepPresident.Recordset.Fields(6).Value = cboPositionVying
            dtaDepPresident.Recordset.Update
            MsgBox "Record Updated Succesfully"
        ElseIf cboPositionVying.Text = "SECRETARY GENERAL" Then
            dtaSecGeneral.Recordset.AddNew
            dtaSecGeneral.Recordset.Fields(1).Value = txtAspirantName.Text
            dtaSecGeneral.Recordset.Fields(2).Value = lngAdmission
            dtaSecGeneral.Recordset.Fields(3).Value = txtClass.Text
            dtaSecGeneral.Recordset.Fields(4).Value = cboDepartment.Text
            dtaSecGeneral.Recordset.Fields(5).Value = lngPhoneNumber
            dtaSecGeneral.Recordset.Fields(6).Value = cboPositionVying
            dtaSecGeneral.Recordset.Update
            MsgBox "Record Updated Succesfully"
        End If
        'Clearing the Inputs
        txtAdmissionNumber.Text = ""
        txtAspirantName.Text = ""
        txtClass.Text = ""
        txtPhoneNumber.Text = ""
        cboDepartment.Text = ""
        cboPositionVying.Text = ""
    End If
End Sub

Private Sub Form_Load()
     'Positioning the Form
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmMain.Show
End Sub
