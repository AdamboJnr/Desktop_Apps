VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCheckOut 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Check Out"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc dtaRooms 
      Height          =   375
      Left            =   240
      Top             =   4800
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      RecordSource    =   "select * from tblRoom"
      Caption         =   "Rooms"
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
   Begin VB.TextBox txtRoomNumber 
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   3120
      Width           =   2415
   End
   Begin MSAdodcLib.Adodc dtaCustomerNumber 
      Height          =   495
      Left            =   240
      Top             =   4080
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Ichaweri Hotel Management System.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Ichaweri Hotel Management System.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from tblCheckIn"
      Caption         =   "Customer Number"
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
      Left            =   4080
      TabIndex        =   7
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdCheckOut 
      Caption         =   "Check Out"
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
      Left            =   2640
      TabIndex        =   6
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox txtCustomerName 
      DataSource      =   "dtaRooms"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   2280
      Width           =   2415
   End
   Begin VB.ComboBox cboCustomerNumber 
      DataSource      =   "dtaCustomerNumber"
      Height          =   315
      Left            =   2880
      TabIndex        =   4
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   3480
      Picture         =   "frmCheckOut.frx":0000
      Stretch         =   -1  'True
      Top             =   480
      Width           =   615
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label lblRoomNumber 
      Alignment       =   1  'Right Justify
      Caption         =   "Room Number"
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
      TabIndex        =   3
      Top             =   3120
      Width           =   2055
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
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label lblCustomerNumber 
      Alignment       =   1  'Right Justify
      Caption         =   "Customer Number"
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
      TabIndex        =   1
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label lblCheckOut 
      Alignment       =   2  'Center
      Caption         =   "Check Out"
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
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "frmCheckOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboCustomerNumber_Click()
    Dim SearchValue As Long
    SearchValue = cboCustomerNumber.Text
    dtaCustomerNumber.Recordset.MoveFirst
    dtaCustomerNumber.Recordset.Find "[Customer ID]= " & SearchValue, 0, adSearchForward
    If dtaCustomerNumber.Recordset.EOF = True Then
        MsgBox ("No Record Found")
        dtaCustomerNumber.Recordset.MoveFirst
    ElseIf dtaCustomerNumber.Recordset.Fields(0).Value = SearchValue Then
        txtCustomerName.Text = dtaCustomerNumber.Recordset.Fields(1).Value
        txtRoomNumber.Text = dtaCustomerNumber.Recordset.Fields(5).Value
    End If
End Sub
Private Sub cmdCheckOut_Click()
    Dim SearchValue As Long
    
    If txtCustomerName.Text = "" Or txtRoomNumber.Text = "" Or cboCustomerNumber.Text = "" Then
        MsgBox "Please Fill In All Inputs"
    Else
        SearchValue = txtRoomNumber.Text
        dtaRooms.Recordset.MoveFirst
        dtaRooms.Recordset.Find "[Room Number]= " & SearchValue, 0, adSearchForward
        If dtaRooms.Recordset.Fields(0).Value = SearchValue Then
            dtaRooms.Recordset.Fields(2).Value = "Available"
            dtaRooms.Recordset.Update
            MsgBox "Room Cleared Succesfully"
        End If
    End If
    txtCustomerName.Text = ""
    txtRoomNumber.Text = ""
    cboCustomerNumber.Text = ""
End Sub
Private Sub Form_Load()
    'Positioning the Form
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
    While dtaCustomerNumber.Recordset.EOF = False
        cboCustomerNumber.AddItem dtaCustomerNumber.Recordset.Fields(0).Value
        dtaCustomerNumber.Recordset.MoveNext
    Wend
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmMain.Show
End Sub
