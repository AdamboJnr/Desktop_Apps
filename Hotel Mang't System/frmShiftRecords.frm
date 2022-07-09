VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmShiftRecords 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Shift Records"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5835
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSAdodcLib.Adodc dtaEmployee 
      Height          =   375
      Left            =   120
      Top             =   3600
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
      RecordSource    =   "select * from tblEmployee"
      Caption         =   "Employee Details"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc dtaShiftRecords 
      Height          =   735
      Left            =   240
      Top             =   4320
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
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
      RecordSource    =   "select * from  tblShift"
      Caption         =   "Shift Records"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3600
      TabIndex        =   10
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   2040
      TabIndex        =   9
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Frame frmAttendance 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2520
      TabIndex        =   6
      Top             =   2520
      Width           =   2535
      Begin VB.OptionButton optAbsent 
         Caption         =   "Absent"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   1695
      End
      Begin VB.OptionButton optPresent 
         Caption         =   "Present"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.TextBox txtEmployeeName 
      DataSource      =   "dtaShiftRecords"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   1920
      Width           =   2535
   End
   Begin VB.ComboBox cboEmployeeNumber 
      DataSource      =   "dtaEmployee"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2520
      TabIndex        =   2
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   3720
      Picture         =   "frmShiftRecords.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   615
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label lblAttendance 
      Alignment       =   1  'Right Justify
      Caption         =   "Attendance"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label lblEmployeeName 
      Alignment       =   1  'Right Justify
      Caption         =   "Employee Name"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label lblEmployeeNumber 
      Alignment       =   1  'Right Justify
      Caption         =   "Employee Number"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblShiftRecords 
      Alignment       =   2  'Center
      Caption         =   "Shift Records"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmShiftRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public UserClocked As Boolean
Private Sub CheckUser()
    Dim EmployeeNumber As Long
    Dim TodayDate As String, dbDate As String, FinalDate As String, TodayFinalDate As String
    
    UserClocked = False
    
    EmployeeNumber = cboEmployeeNumber.Text
    dtaShiftRecords.Recordset.MoveFirst
    While dtaShiftRecords.Recordset.EOF = False
        If dtaShiftRecords.Recordset.Fields(3).Value = EmployeeNumber Then
            dbDate = dtaShiftRecords.Recordset.Fields(2).Value
            FinalDate = Mid(dbDate, 1, 5)
            TodayDate = Format(Now, "mm/dd/yy hh:mm:ss")
            TodayFinalDate = Mid(TodayDate, 1, 5)
            If FinalDate = TodayFinalDate Then
                'MsgBox "User has already been Clocked", vbCritical
                UserClocked = True
                Exit Sub
            'Else

            End If
        End If
        dtaShiftRecords.Recordset.MoveNext
    Wend
End Sub

Private Sub cboEmployeeNumber_Click()
    Dim SearchValue As Long
    SearchValue = cboEmployeeNumber.Text
    dtaEmployee.Recordset.MoveFirst
    dtaEmployee.Recordset.Find "[Employee Number]= " & SearchValue, 0, adSearchForward
    If dtaEmployee.Recordset.EOF = True Then
        MsgBox ("No Record Found")
        dtaEmployee.Recordset.MoveFirst
    ElseIf dtaEmployee.Recordset.Fields(0).Value = SearchValue Then
        txtEmployeeName.Text = dtaEmployee.Recordset.Fields(1).Value
    End If
 
End Sub
Private Sub cmdCancel_Click()
    cboEmployeeNumber.Text = ""
    txtEmployeeName.Text = ""
    optPresent.Value = False
    optAbsent.Value = False
End Sub
Private Sub cmdSave_Click()
    Dim lngEmployeeNumber As Long

    
    If txtEmployeeName.Text = "" Or cboEmployeeNumber.Text = "" Then
        MsgBox "Please Fill In All Inputs", vbCritical
    Else
        lngEmployeeNumber = cboEmployeeNumber.Text
        
        Call CheckUser
        
        If UserClocked = True Then
            MsgBox "User Already Clocked", vbCritical
            cboEmployeeNumber.Text = ""
            txtEmployeeName.Text = ""
            optAbsent.Value = False
            optPresent.Value = False
            Exit Sub
        Else
            dtaShiftRecords.Recordset.AddNew
            If optAbsent.Value = True Then
                dtaShiftRecords.Recordset.Fields(1).Value = "Absent"
            ElseIf optPresent.Value = True Then
                dtaShiftRecords.Recordset.Fields(1).Value = "Present"
            End If
            dtaShiftRecords.Recordset.Fields(2).Value = Format(Now, "mm/dd/yy hh:mm:ss")
            dtaShiftRecords.Recordset.Fields(3).Value = lngEmployeeNumber
            dtaShiftRecords.Recordset.Fields(4).Value = txtEmployeeName.Text
            dtaShiftRecords.Recordset.Update
            MsgBox "Record Updated Succesfully"
            cboEmployeeNumber.Text = ""
            txtEmployeeName.Text = ""
            optPresent.Value = False
            optAbsent.Value = False
        End If
    End If
End Sub
Private Sub Form_Load()
    While dtaEmployee.Recordset.EOF = False
        cboEmployeeNumber.AddItem dtaEmployee.Recordset.Fields(0).Value
        dtaEmployee.Recordset.MoveNext
    Wend
    'Positioning the Form
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmMain.Show
End Sub
