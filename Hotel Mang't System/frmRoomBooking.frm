VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmRoomBooking 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Room Booking"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc dtaRooms 
      Height          =   375
      Left            =   3360
      Top             =   5520
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Ichaweri Hotel Management System.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Ichaweri Hotel Management System.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from tblRoom"
      Caption         =   "Room Number"
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
      Left            =   480
      Top             =   5400
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      RecordSource    =   "select *  from tblCheckIn"
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
      Left            =   8400
      TabIndex        =   14
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdProceed 
      Caption         =   "Proceed"
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
      Left            =   6960
      TabIndex        =   13
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Frame fraRoomBooking 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   5040
      TabIndex        =   2
      Top             =   1320
      Width           =   4455
      Begin VB.TextBox txtTypeOfRoom 
         Height          =   375
         Left            =   2160
         TabIndex        =   18
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtNumberOfCustomers 
         Height          =   375
         Left            =   2160
         TabIndex        =   16
         Top             =   2040
         Width           =   1935
      End
      Begin VB.ComboBox cboRoomNumber 
         DataSource      =   "dtaRooms"
         Height          =   315
         Left            =   2160
         TabIndex        =   12
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblTypeOfRoom 
         Alignment       =   1  'Right Justify
         Caption         =   "Type Of Room"
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
         TabIndex        =   17
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lblNumberOfCustomers 
         Alignment       =   1  'Right Justify
         Caption         =   "Number of customers"
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
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   1815
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
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame fraCustomerDetails 
      Height          =   3375
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   4455
      Begin VB.TextBox txtResidence 
         Height          =   375
         Left            =   2160
         TabIndex        =   10
         Top             =   2520
         Width           =   1935
      End
      Begin VB.TextBox txtPhoneNumber 
         Height          =   375
         Left            =   2160
         TabIndex        =   9
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox txtIDNumber 
         Height          =   375
         Left            =   2160
         TabIndex        =   6
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtCustomerName 
         DataSource      =   "dtaRoomBooking"
         Height          =   375
         Left            =   2160
         TabIndex        =   4
         Top             =   360
         Width           =   1935
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
         TabIndex        =   8
         Top             =   2520
         Width           =   1575
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
         TabIndex        =   7
         Top             =   1800
         Width           =   1575
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
         TabIndex        =   5
         Top             =   1080
         Width           =   1575
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
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   6480
      Picture         =   "frmRoomBooking.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   2160
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   5055
   End
   Begin VB.Label lblRoomBooking 
      Alignment       =   2  'Center
      Caption         =   "Room Booking"
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
      Left            =   2280
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
End
Attribute VB_Name = "frmRoomBooking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboRoomNumber_Click()
    Dim lngRoomNumber As Long
    lngRoomNumber = cboRoomNumber.Text
    dtaRooms.Recordset.MoveFirst
    dtaRooms.Recordset.Find "[Room Number]= " & lngRoomNumber, 0, adSearchForward
    If dtaRooms.Recordset.EOF = True Then
        MsgBox ("No Record Found")
        dtaRooms.Recordset.MoveFirst
    ElseIf dtaRooms.Recordset.Fields(0).Value = lngRoomNumber Then
       txtTypeOfRoom.Text = dtaRooms.Recordset.Fields(3).Value
    End If
End Sub

Private Sub cmdProceed_Click()
    Dim lngIDNumber As Long
    Dim lngPhoneNumber As Long
    Dim intNumberOfCustomers As Integer
    If txtCustomerName.Text = "" Or txtIDNumber.Text = "" Or txtNumberOfCustomers.Text = "" Or txtPhoneNumber.Text = "" Or txtResidence.Text = "" Or txtTypeOfRoom.Text = "" Or cboRoomNumber.Text = "" Then
        MsgBox "Please Fill In All Inputs"
    Else
        lngIDNumber = txtIDNumber.Text
        lngPhoneNumber = txtPhoneNumber.Text
        intNumberOfCustomers = txtNumberOfCustomers.Text
        'Saving to the Database
        dtaRoomBooking.Recordset.AddNew
        dtaRoomBooking.Recordset.Fields(1).Value = txtCustomerName.Text
        dtaRoomBooking.Recordset.Fields(2).Value = lngIDNumber
        dtaRoomBooking.Recordset.Fields(3).Value = lngPhoneNumber
        dtaRoomBooking.Recordset.Fields(4).Value = txtResidence.Text
        dtaRoomBooking.Recordset.Fields(5).Value = cboRoomNumber.Text
        dtaRoomBooking.Recordset.Fields(6).Value = intNumberOfCustomers
        dtaRoomBooking.Recordset.Fields(7).Value = txtTypeOfRoom.Text
        dtaRoomBooking.Recordset.Update
        
        lngCustomerId = dtaRoomBooking.Recordset.Fields(0).Value
        txtCustomerName.Text = ""
        txtIDNumber.Text = ""
        txtPhoneNumber.Text = ""
        txtResidence.Text = ""
        cboRoomNumber.Text = ""
        txtTypeOfRoom.Text = ""
        txtNumberOfCustomers.Text = ""
        Me.Hide
        frmPaymentsDetails.Show
    End If
End Sub
Private Sub Form_Load()
    'Positioning the Form
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
    
    dtaRooms.Recordset.MoveFirst
    While dtaRooms.Recordset.EOF = False
        If dtaRooms.Recordset.Fields(2).Value = "Available" Then
            cboRoomNumber.AddItem dtaRooms.Recordset.Fields(0).Value
        End If
        dtaRooms.Recordset.MoveNext
    Wend
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmMain.Show
End Sub


