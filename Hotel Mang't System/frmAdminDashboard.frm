VERSION 5.00
Begin VB.Form frmAdminDashboard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Admin Dashboard"
   ClientHeight    =   5055
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCreateUser 
      Caption         =   "Create User"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton cmdPaymentReport 
      Caption         =   "Payment Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2760
      TabIndex        =   2
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton cmdUserLog 
      Caption         =   "User Log"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1440
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create User"
      Height          =   1095
      Left            =   -2160
      TabIndex        =   0
      Top             =   -2640
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   2175
      Left            =   840
      Picture         =   "frmAdminDashboard.frx":0000
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Ichaweri Hotel Management System"
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
      Left            =   360
      TabIndex        =   3
      Top             =   4080
      Width           =   4335
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   240
      Top             =   3960
      Width           =   4575
   End
   Begin VB.Menu mnuCreate 
      Caption         =   "&Create"
      Begin VB.Menu mnuCreateNewUser 
         Caption         =   "Create New User"
      End
   End
   Begin VB.Menu mnuLogs 
      Caption         =   "L&ogs"
      Begin VB.Menu mnuViewUserLogs 
         Caption         =   "View User Logs"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuPaymentReport 
         Caption         =   "Payment Report"
      End
      Begin VB.Menu mnuRoomBookingReports 
         Caption         =   "Room Booking Reports"
      End
      Begin VB.Menu mnuEmployeeRecords 
         Caption         =   "Emloyee Records Report"
      End
   End
End
Attribute VB_Name = "frmAdminDashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCreateUser_Click()
    frmAdminDashboard.Hide
    'Unload frmMain
    frmCreateUser.Show
End Sub
Private Sub cmdPaymentReport_Click()
    frmAdminDashboard.Hide
    frmPaymentReport.Show
End Sub
Private Sub cmdUserLog_Click()
    frmAdminDashboard.Hide
    frmUserLogs.Show
End Sub
Private Sub Form_Load()
    'Positioning the Form
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmMain.Show
End Sub
Private Sub mnuCreateNewUser_Click()
    frmAdminDashboard.Hide
    frmCreateUser.Show
End Sub
Private Sub mnuEmployeeRecords_Click()
    rptEmployeeRecords.Show
End Sub

Private Sub mnuPaymentReport_Click()
    frmAdminDashboard.Hide
    frmPaymentReport.Show
End Sub
Private Sub mnuRoomBookingReports_Click()
    frmAdminDashboard.Hide
    frmRommBookingReport.Show
End Sub
Private Sub mnuViewUserLogs_Click()
    frmAdminDashboard.Hide
    frmUserLogs.Show
End Sub
