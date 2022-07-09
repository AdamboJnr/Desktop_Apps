VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Home"
   ClientHeight    =   7005
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdminDashboard 
      Caption         =   "Dashboard"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5400
      Picture         =   "frmMain.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton cmdLogOut 
      Caption         =   "Log Out"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6720
      Picture         =   "frmMain.frx":1082
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton cmdClearRoom 
      Caption         =   "Clear Room"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4080
      Picture         =   "frmMain.frx":4990C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton cmdShiftRecords 
      Caption         =   "Shift Records"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2760
      Picture         =   "frmMain.frx":4A98E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton cmdAddEmployee 
      Caption         =   " Add  Employee"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1440
      Picture         =   "frmMain.frx":4B2C0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton cmdBookRoom 
      Caption         =   "Book Room"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      Picture         =   "frmMain.frx":4B9E7
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   5760
      Width           =   7335
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      Caption         =   "ICHAWERI MANAGEMENT SYSTEM"
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
      TabIndex        =   5
      Top             =   5880
      Width           =   7095
   End
   Begin VB.Image Image1 
      Height          =   3735
      Left            =   1320
      Picture         =   "frmMain.frx":4CA69
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   5655
   End
   Begin VB.Menu mnuRoom 
      Caption         =   "Room"
      Begin VB.Menu mnuRoomBooking 
         Caption         =   "Room Booking"
      End
   End
   Begin VB.Menu mnuEmployee 
      Caption         =   "Employee"
      Begin VB.Menu mnuCreateEmployee 
         Caption         =   "Create Employee"
      End
   End
   Begin VB.Menu mnuShift 
      Caption         =   "Shift"
      Begin VB.Menu mnuShiftRecords 
         Caption         =   "ShiftRecords"
      End
   End
   Begin VB.Menu mnuCheckOut 
      Caption         =   "Check Out"
      Begin VB.Menu mnuClearRoom 
         Caption         =   "Clear Room"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAddEmployee_Click()
    Unload Me
    frmcreatemployee.Show
End Sub

Private Sub cmdAdminDashboard_Click()
    Unload Me
    'Unload frmLogin
    If strUserType = "Admin" Then
        frmAdminDashboard.Show
    Else
        MsgBox "Only Admins can Access This Page", vbCritical
        frmMain.Show
    End If
End Sub

Private Sub cmdBookRoom_Click()
    Unload Me
    frmRoomBooking.Show
End Sub

Private Sub cmdClearRoom_Click()
    Unload Me
    frmCheckOut.Show
End Sub

Private Sub cmdLogOut_Click()
    If MsgBox("Are You Sure You Want to Log Out?", vbOKCancel + vbQuestion) = vbOK Then
    Unload Me
    frmLogin.Show
    'Clearing Inputs
    frmLogin.txtPassword.Text = ""
    frmLogin.txtUsername.Text = ""
    End If
End Sub

Private Sub cmdShiftRecords_Click()
    Unload Me
    frmShiftRecords.Show
End Sub

Private Sub Form_Load()
    'Positioning the Form
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'frmLogin.Show
End Sub

Private Sub mnuClearRoom_Click()
    Unload Me
    frmCheckOut.Show
End Sub

Private Sub mnuCreateEmployee_Click()
    Unload Me
    frmcreatemployee.Show
End Sub

Private Sub mnuRoomBooking_Click()
    Unload Me
    frmRoomBooking.Show
End Sub

Private Sub mnuShiftRecords_Click()
    Unload Me
    frmShiftRecords.Show
End Sub
