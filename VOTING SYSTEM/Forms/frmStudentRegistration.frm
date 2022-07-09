VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmStudentRegistration 
   Caption         =   "Student Voter Registration"
   ClientHeight    =   6345
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc dtaStudentRegistration 
      Height          =   375
      Left            =   840
      Top             =   5520
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Voting System.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Voting System.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from tblStudentsDetails"
      Caption         =   "Student Registration"
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
      Left            =   5520
      TabIndex        =   12
      Top             =   5400
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
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Frame fraStudentDetails 
      Height          =   3975
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   6375
      Begin VB.ComboBox cboDepartment 
         Height          =   315
         ItemData        =   "frmStudentRegistration.frx":0000
         Left            =   2640
         List            =   "frmStudentRegistration.frx":0019
         TabIndex        =   13
         Top             =   2520
         Width           =   2895
      End
      Begin VB.TextBox txtPhoneNumber 
         Height          =   405
         Left            =   2640
         TabIndex        =   10
         Top             =   3240
         Width           =   2895
      End
      Begin VB.TextBox txtClass 
         Height          =   375
         Left            =   2640
         TabIndex        =   9
         Top             =   1800
         Width           =   2895
      End
      Begin VB.TextBox txtAdmissionNumber 
         Height          =   375
         Left            =   2640
         TabIndex        =   8
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox txtStudentName 
         DataSource      =   "dtaStudentRegistration"
         Height          =   375
         Left            =   2640
         TabIndex        =   7
         Top             =   360
         Width           =   2895
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
         Left            =   360
         TabIndex        =   6
         Top             =   3240
         Width           =   1815
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
         Left            =   360
         TabIndex        =   5
         Top             =   2520
         Width           =   1815
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
         Left            =   360
         TabIndex        =   4
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label2 
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
         Left            =   360
         TabIndex        =   3
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Caption         =   "Student Name"
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
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   5280
      Picture         =   "frmStudentRegistration.frx":008D
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   5295
   End
   Begin VB.Label lblStudentRegistrationCaption 
      Alignment       =   2  'Center
      Caption         =   "Student Voter Registration"
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
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   4215
   End
End
Attribute VB_Name = "frmStudentRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    'Clearing The Inputs
    txtAdmissionNumber.Text = ""
    txtClass.Text = ""
    txtPhoneNumber.Text = ""
    txtStudentName.Text = ""
    cboDepartment.Text = ""
End Sub

Private Sub cmdSave_Click()
    'variables to hold the type long inputs
    Dim lngAdmission As Long
    Dim lngNumber As Long
    'Checking For Empty Inputs
    If txtAdmissionNumber.Text = "" Or txtClass.Text = "" Or txtPhoneNumber.Text = "" Or txtStudentName.Text = "" Then
        MsgBox "Please Fill In all Inputs"
    Else
        'Assigning the variables
        lngAdmission = txtAdmissionNumber
        lngNumber = txtPhoneNumber
        'Saving the data in the Database
        dtaStudentRegistration.Recordset.AddNew
        dtaStudentRegistration.Recordset.Fields(0).Value = lngAdmission
        dtaStudentRegistration.Recordset.Fields(1).Value = txtStudentName.Text
        dtaStudentRegistration.Recordset.Fields(2).Value = txtClass.Text
        dtaStudentRegistration.Recordset.Fields(3).Value = cboDepartment.Text
        dtaStudentRegistration.Recordset.Fields(4).Value = lngNumber
        dtaStudentRegistration.Recordset.Update
        'Updating the user on the saved record
        MsgBox "Record Updated Succesfully"
        
        txtAdmissionNumber.Text = ""
        txtClass.Text = ""
        txtPhoneNumber.Text = ""
        txtStudentName.Text = ""
    End If
End Sub

Private Sub Form_Load()
    'Positioning the Form
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'If MsgBox("Do You Want to Exit?", vbOKCancel) Then
    frmMain.Show
    'End If
End Sub
