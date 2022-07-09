VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmDeletePolicy 
   Caption         =   "Delete Customer Policy"
   ClientHeight    =   5535
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6585
   LinkTopic       =   "Form2"
   ScaleHeight     =   5535
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc dtaAllPoliciesDeleted 
      Height          =   375
      Left            =   1920
      Top             =   4560
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      RecordSource    =   "select * from tblDeletedPolicies"
      Caption         =   "Policies Deleted"
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
   Begin MSAdodcLib.Adodc dtaPolicyDelete 
      Height          =   375
      Left            =   1920
      Top             =   5160
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      RecordSource    =   "select * from tblAcceptedRejectedPolicies"
      Caption         =   "Policy Delete"
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
   Begin MSAdodcLib.Adodc dtaDeletePolicy 
      Height          =   450
      Left            =   120
      Top             =   4920
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   794
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
      RecordSource    =   "select * from tblPolicyCreation"
      Caption         =   "Delete Policy"
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
   Begin MSAdodcLib.Adodc dtaPolicyNumber 
      Height          =   375
      Left            =   0
      Top             =   4440
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Insurance management database.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Insurance management database.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from tblPolicyCreation"
      Caption         =   "Policy Number"
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
      Height          =   855
      Left            =   5160
      Picture         =   "frmDeletePolicy.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3720
      Picture         =   "frmDeletePolicy.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Frame fraCustomerDeletion 
      Height          =   2895
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   6015
      Begin VB.TextBox txtReasonForDeletion 
         DataSource      =   "dtaAllPoliciesDeleted"
         Height          =   375
         Left            =   2400
         TabIndex        =   11
         Top             =   2040
         Width           =   2415
      End
      Begin VB.TextBox txtPolicyType 
         DataSource      =   "dtaDeletePolicy"
         Height          =   375
         Left            =   2400
         TabIndex        =   7
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox txtCustomerName 
         DataSource      =   "dtaPolicyDelete"
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   1440
         Width           =   2415
      End
      Begin VB.ComboBox cboPolicyNumber 
         DataSource      =   "dtaPolicyNumber"
         Height          =   315
         Left            =   2400
         TabIndex        =   3
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lblReasonForDeletion 
         Alignment       =   1  'Right Justify
         Caption         =   "Reason"
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
         Left            =   120
         TabIndex        =   10
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label lblPolicyType 
         Alignment       =   1  'Right Justify
         Caption         =   "Policy Type"
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
         TabIndex        =   6
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblCustomerName 
         Alignment       =   1  'Right Justify
         Caption         =   "Customer Name"
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
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label lblPolicyNumber 
         Alignment       =   1  'Right Justify
         Caption         =   "Policy Number"
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
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   4800
      Picture         =   "frmDeletePolicy.frx":0884
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblDeleteCustomerPolicy 
      Alignment       =   2  'Center
      Caption         =   "Delete Customer Policy"
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
      Top             =   360
      Width           =   3375
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   1080
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "frmDeletePolicy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Populate_Combo()
    While dtaPolicyNumber.Recordset.EOF = False
        cboPolicyNumber.AddItem dtaPolicyNumber.Recordset.Fields(0).Value
        dtaPolicyNumber.Recordset.MoveNext
    Wend
End Sub
Private Sub cboPolicyNumber_Click()
    Dim searchvalue As Long
    
    searchvalue = cboPolicyNumber.Text
    'Automatically filling in the Insurance Name
    dtaPolicyNumber.Recordset.MoveFirst
    dtaPolicyNumber.Recordset.Find "[Policy Number]= " & searchvalue, 0, adSearchForward
    If dtaPolicyNumber.Recordset.EOF = True Then
        MsgBox "Record Not Found"
        dtaPolicyNumber.Recordset.MoveFirst
    ElseIf dtaPolicyNumber.Recordset.Fields(0).Value = searchvalue Then
        txtPolicyType.Text = dtaPolicyNumber.Recordset.Fields(5).Value
        txtCustomerName.Text = dtaPolicyNumber.Recordset.Fields(1).Value
    End If
End Sub
Private Sub cmdCancel_Click()
    If cboPolicyNumber.Text = "" And txtCustomerName.Text = "" And txtPolicyType.Text = "" And txtReasonForDeletion.Text = "" Then
        Unload Me
        frmAdminDashboard.Show
    Else
        cboPolicyNumber.Text = ""
        txtCustomerName.Text = ""
        txtPolicyType.Text = ""
        txtReasonForDeletion.Text = ""
    End If
End Sub
Private Sub cmdDelete_Click()
    Dim searchvalue As Long
    
    If cboPolicyNumber.Text = "" Or txtReasonForDeletion.Text = "" Or txtCustomerName.Text = "" Or txtPolicyType.Text = "" Then
        MsgBox "Please Fill In All Inputs", vbCritical
    Else
        searchvalue = cboPolicyNumber.Text
        'Saving to Database First
        dtaAllPoliciesDeleted.Recordset.AddNew
        dtaAllPoliciesDeleted.Recordset.Fields(0).Value = searchvalue
        dtaAllPoliciesDeleted.Recordset.Fields(1).Value = txtCustomerName.Text
        dtaAllPoliciesDeleted.Recordset.Fields(2).Value = txtPolicyType.Text
        dtaAllPoliciesDeleted.Recordset.Fields(3).Value = txtReasonForDeletion.Text
        dtaAllPoliciesDeleted.Recordset.Fields(4).Value = Format(Now, "mm/dd/yy hh:mm:ss")
        dtaAllPoliciesDeleted.Recordset.Update
        
        'Deleting the policies
        dtaDeletePolicy.Recordset.MoveFirst
        dtaPolicyDelete.Recordset.MoveFirst
        dtaDeletePolicy.Recordset.Find "[Policy Number]= " & searchvalue, 0, adSearchForward
        dtaPolicyDelete.Recordset.Find "[Policy Number]= " & searchvalue, 0, adSearchForward
        If dtaDeletePolicy.Recordset.EOF = True Then
            MsgBox "Record Not Found"
            dtaDeletePolicy.Recordset.MoveFirst
            dtaPolicyDelete.Recordset.MoveFirst
        ElseIf dtaDeletePolicy.Recordset.Fields(0).Value = searchvalue Then
            If MsgBox("Are you sure you want to delete Plan?", vbOKCancel + vbQuestion) = vbOK Then
                dtaDeletePolicy.Recordset.Delete
                dtaPolicyDelete.Recordset.Delete
                MsgBox "Insurance Policy Deleted Succesfully", vbInformation
            End If
            cboPolicyNumber.Clear
            txtCustomerName.Text = ""
            txtPolicyType.Text = ""
            txtReasonForDeletion.Text = ""
            dtaDeletePolicy.Refresh
            Call Populate_Combo
        End If
    End If
End Sub
Private Sub Form_Load()
    'Positioning the form
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
    Call Populate_Combo
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
    frmAdminDashboard.Show
End Sub
