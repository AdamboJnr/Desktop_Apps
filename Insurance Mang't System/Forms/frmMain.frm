VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   7335
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   9585
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   9585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Left            =   7800
      Picture         =   "frmMain.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton cmdAdmin 
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
      Left            =   6240
      Picture         =   "frmMain.frx":4888A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton cmdCommission 
      Caption         =   "Commission"
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
      Left            =   4680
      Picture         =   "frmMain.frx":490D2
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton cmdAgentRegister 
      Caption         =   "Register Agent"
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
      Left            =   3120
      Picture         =   "frmMain.frx":49949
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton cmdPolicyPayment 
      Caption         =   "Policy Payment"
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
      Left            =   1560
      Picture         =   "frmMain.frx":4A27B
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton cmdCreatePolicy 
      Caption         =   "Create policy"
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
      Left            =   0
      Picture         =   "frmMain.frx":4AA18
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label lblCompanyCaption 
      Alignment       =   2  'Center
      Caption         =   "GEMINIA INSURANCE MANAGEMENT SYSTEM"
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
      Left            =   720
      TabIndex        =   5
      Top             =   6360
      Width           =   8295
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   6240
      Width           =   8535
   End
   Begin VB.Image Image1 
      Height          =   4455
      Left            =   1440
      Picture         =   "frmMain.frx":4B18E
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   6495
   End
   Begin VB.Menu mnuPolicy 
      Caption         =   "Policy"
      Begin VB.Menu mnuPolicyCreation 
         Caption         =   "Policy Creation"
      End
      Begin VB.Menu mnuPolicyClaim 
         Caption         =   "Policy Claim"
      End
   End
   Begin VB.Menu mnuPayment 
      Caption         =   "Payment"
      Begin VB.Menu mnuPolicyPayment 
         Caption         =   "Policy Payment"
      End
   End
   Begin VB.Menu mnuAgents 
      Caption         =   "Agents"
      Begin VB.Menu mnuAgentRegistration 
         Caption         =   "Agent Registration"
      End
      Begin VB.Menu mnuAgentCommissionWithdraw 
         Caption         =   "Agent Commission Withdraw"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdmin_Click()
    If strTypeUser = "Admin" Then
        frmMain.Hide
        frmAdminDashboard.Show
    Else
        If MsgBox("Only Admins can Access this page, Would you like to Log In as an Admin", vbOKCancel + vbQuestion) = vbOK Then
            Me.Hide
            frmLogIn.Show
        End If
    End If
End Sub
Private Sub cmdAgentRegister_Click()
    frmMain.Hide
    frmAgentRegistration.Show
End Sub
Private Sub cmdCommission_Click()
    frmMain.Hide
    frmAgentWithdrawCommission.Show
End Sub
Private Sub cmdCreatePolicy_Click()
    frmMain.Hide
    frmPolicyCreationForm.Show
End Sub
Private Sub cmdLogOut_Click()
    frmMain.Hide
    frmLogIn.Show
End Sub
Private Sub cmdPolicyPayment_Click()
    frmMain.Hide
    frmPolicyPayment.Show
End Sub

Private Sub Form_Load()
    'Positioning the form
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmLogIn.Show
End Sub

Private Sub mnuAgentCommissionWithdraw_Click()
    frmMain.Hide
    frmAgentWithdrawCommission.Show
End Sub
Private Sub mnuAgentRegistration_Click()
    frmMain.Hide
    frmAgentRegistration.Show
End Sub

Private Sub mnuPolicyClaim_Click()
    frmMain.Hide
    frmPolicyClaimModule.Show
End Sub
Private Sub mnuPolicyCreation_Click()
    frmMain.Hide
    frmPolicyCreationForm.Show
End Sub
Private Sub mnuPolicyPayment_Click()
    frmMain.Hide
    frmPolicyPayment.Show
End Sub
