VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmcreatemployee 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc dtaCreateEmployee 
      Height          =   855
      Left            =   600
      Top             =   5160
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1508
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
      RecordSource    =   "select * from tblEmployee"
      Caption         =   "Create Employee"
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
      Height          =   495
      Left            =   3480
      TabIndex        =   14
      Top             =   5040
      Width           =   1215
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
      Height          =   495
      Left            =   5040
      TabIndex        =   13
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Frame fraEmployeeDetails 
      Height          =   3735
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   6015
      Begin VB.TextBox txtIDNumber 
         Height          =   375
         Left            =   2520
         TabIndex        =   15
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox txtResidence 
         DataSource      =   "dtaCreateEmployee"
         Height          =   375
         Left            =   2520
         TabIndex        =   12
         Top             =   2760
         Width           =   2655
      End
      Begin VB.TextBox txtPhoneNumber 
         Height          =   375
         Left            =   2520
         TabIndex        =   10
         Top             =   2160
         Width           =   2655
      End
      Begin VB.TextBox txtAge 
         Height          =   375
         Left            =   2520
         TabIndex        =   7
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox txtEmployeeName 
         Height          =   375
         Left            =   2520
         TabIndex        =   6
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label lblResidence 
         Alignment       =   1  'Right Justify
         Caption         =   "Residence"
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
         TabIndex        =   11
         Top             =   2760
         Width           =   1695
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
         TabIndex        =   9
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label lblIDNumber 
         Alignment       =   1  'Right Justify
         Caption         =   "ID Number"
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
         TabIndex        =   8
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label lblAge 
         Alignment       =   1  'Right Justify
         Caption         =   "Age"
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
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblEmployeeName 
         Alignment       =   1  'Right Justify
         Caption         =   "Employee Name"
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
         Width           =   1695
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   3960
      Picture         =   "frmcreatemployee.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   615
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label lblCreateEmployee 
      Alignment       =   2  'Center
      Caption         =   "Create Employee"
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
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmcreatemployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSave_Click()
    Dim lngIDNumber As Long
    Dim lngPhoneNumber As Long
    Dim lngAge As Long
    
    'Checking for inputs
    If txtAge.Text = "" Or txtEmployeeName.Text = "" Or txtIDNumber.Text = "" Or txtPhoneNumber.Text = "" Or txtResidence.Text = "" Then
        MsgBox "Please Fill in All Inputs"
    Else
        lngIDNumber = txtIDNumber.Text
        lngPhoneNumber = txtPhoneNumber.Text
        lngAge = txtAge.Text
        dtaCreateEmployee.Recordset.AddNew
        dtaCreateEmployee.Recordset.Fields(1).Value = txtEmployeeName.Text
        dtaCreateEmployee.Recordset.Fields(2).Value = lngIDNumber
        dtaCreateEmployee.Recordset.Fields(3).Value = lngPhoneNumber
        dtaCreateEmployee.Recordset.Fields(4).Value = lngAge
        dtaCreateEmployee.Recordset.Fields(5).Value = txtResidence.Text
        dtaCreateEmployee.Recordset.Update
        MsgBox "record saved"
        txtAge.Text = ""
        txtEmployeeName.Text = ""
        txtIDNumber.Text = ""
        txtPhoneNumber.Text = ""
        txtResidence.Text = ""
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


