VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmRommBookingReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Room Booking Report"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdView 
      Caption         =   "view"
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
      Left            =   2880
      TabIndex        =   10
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox txtRoomAmount 
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   4200
      Width           =   2415
   End
   Begin MSAdodcLib.Adodc dtaRooms 
      Height          =   495
      Left            =   120
      Top             =   5040
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
   Begin MSAdodcLib.Adodc dtaRoomBooking 
      Height          =   495
      Left            =   240
      Top             =   5640
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "Room Booking"
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
   Begin VB.CommandButton cmdViewAll 
      Caption         =   "View All"
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
      Left            =   4200
      TabIndex        =   7
      Top             =   4920
      Width           =   1095
   End
   Begin VB.ListBox lstDetails 
      Height          =   840
      Left            =   2880
      TabIndex        =   6
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox txtAccounts 
      DataSource      =   "dtaRoomBooking"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   3360
      Width           =   2415
   End
   Begin VB.ComboBox cboRoomNumber 
      DataSource      =   "dtaRooms"
      Height          =   315
      Left            =   2880
      TabIndex        =   2
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label lblRoomAmount 
      Alignment       =   1  'Right Justify
      Caption         =   "Room Amount"
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
      TabIndex        =   8
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   1080
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   3375
   End
   Begin VB.Label lblAccounts 
      Alignment       =   1  'Right Justify
      Caption         =   "count"
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
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label lblDetails 
      Alignment       =   1  'Right Justify
      Caption         =   "Details"
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
      Top             =   2160
      Width           =   2055
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
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label lblRoomBookingReport 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Room Booking Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1200
      TabIndex        =   0
      Top             =   600
      Width           =   3165
   End
End
Attribute VB_Name = "frmRommBookingReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboRoomNumber_Click()
    lstDetails.Clear
    Dim amount As Currency, TotalAmount As Currency
    Dim roomnumber As Long
    'cboRoomNumber.Text = ""
    Dim SearchValue As String
    SearchValue = cboRoomNumber.Text
    dtaRoomBooking.Recordset.MoveFirst
    While dtaRoomBooking.Recordset.EOF = False
        If dtaRoomBooking.Recordset.Fields(5).Value = SearchValue Then
            lstDetails.AddItem dtaRoomBooking.Recordset.Fields(1).Value
        End If
        dtaRoomBooking.Recordset.MoveNext
    Wend
    txtAccounts.Text = lstDetails.ListCount
    'Calculating Total Room Amount
    dtaRooms.Recordset.MoveFirst
    roomnumber = cboRoomNumber.Text
    dtaRooms.Recordset.Find "[Room Number]= " & roomnumber, 0, adSearchForward
    If dtaRooms.Recordset.EOF = True Then
        MsgBox ("No Record Found")
        dtaRooms.Recordset.MoveFirst
    ElseIf dtaRooms.Recordset.Fields(0).Value = roomnumber Then
        amount = dtaRooms.Recordset.Fields(4).Value
        TotalAmount = amount * Val(txtAccounts.Text)
        txtRoomAmount.Text = TotalAmount
    End If
End Sub
Private Sub cmdViewAll_Click()
    rptRoomReport2.Show
End Sub
Private Sub Command1_Click()
    Dim lngRoomNumber As Long
    lngRoomNumber = cboRoomNumber.Text
    denExample.Example lngRoomNumber
    rptExample.Show
End Sub
Private Sub Form_Load()
    'Positioning the Form
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
    
    dtaRooms.Recordset.MoveFirst
    While dtaRooms.Recordset.EOF = False
        cboRoomNumber.AddItem dtaRooms.Recordset.Fields(0).Value
        dtaRooms.Recordset.MoveNext
    Wend
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmAdminDashboard.Show
End Sub
