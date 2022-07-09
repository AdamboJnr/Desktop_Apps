VERSION 5.00
Begin VB.Form frmAdminDashboard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Admin Dashboard"
   ClientHeight    =   6855
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   8070
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPolicyPaymentReport 
      Caption         =   "Policy Payment Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5400
      Picture         =   "frmAdminDashboard.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton cmdAcceptRejectClaim 
      Caption         =   "Accept/Reject Claim"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3600
      Picture         =   "frmAdminDashboard.frx":079D
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton cmdAddDeletePlan 
      Caption         =   "Add Plan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1800
      Picture         =   "frmAdminDashboard.frx":49027
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton cmdDeleteCustomerPolicy 
      Caption         =   "Delete Customer Policy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      Picture         =   "frmAdminDashboard.frx":918B1
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   3855
      Left            =   1080
      Picture         =   "frmAdminDashboard.frx":DA13B
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   5295
   End
   Begin VB.Label lblAdminDashboard 
      Alignment       =   2  'Center
      Caption         =   "Admin Dashboard"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   5880
      Width           =   7095
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   5760
      Width           =   7335
   End
   Begin VB.Menu mnuCustomers 
      Caption         =   "Customers"
      Begin VB.Menu mnuDeleteCustomerPolicy 
         Caption         =   "Delete Customer Policy"
      End
      Begin VB.Menu mnuRegistered 
         Caption         =   "View Registered Customers"
      End
      Begin VB.Menu mnuUpdateDetails 
         Caption         =   "Update Policy details"
      End
      Begin VB.Menu mnuPolicyClaim 
         Caption         =   "PolicyClaimReport"
      End
   End
   Begin VB.Menu mnuInsurancePlans 
      Caption         =   "Insurance Plans"
      Begin VB.Menu mnuAddnsurancePlan 
         Caption         =   "Add Insurance Plan"
      End
      Begin VB.Menu mnuDeleteInsurancePlan 
         Caption         =   "Delete Insurance Plan"
      End
   End
   Begin VB.Menu mnuPolicies 
      Caption         =   "Policies"
      Begin VB.Menu mnuAcceptRejectPolicy 
         Caption         =   "Accept/Reject Policy"
      End
      Begin VB.Menu mnuAcceptRejectClaim 
         Caption         =   "Accept/Reject Claim"
      End
   End
   Begin VB.Menu mnuCreate 
      Caption         =   "Create"
      Begin VB.Menu mnuCreateAccount 
         Caption         =   "Create Account"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "Reports"
      Begin VB.Menu mnuAgentReport 
         Caption         =   "Agents Report"
      End
      Begin VB.Menu mnuPolicyPaymentReport 
         Caption         =   "Policy Payment Report"
      End
      Begin VB.Menu mnuTypesOfAccountReport 
         Caption         =   "Types Of Account Report"
      End
      Begin VB.Menu mnuUserLogs 
         Caption         =   "UserLogs Report"
      End
   End
End
Attribute VB_Name = "frmAdminDashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAcceptRejectClaim_Click()
    Me.Hide
    frmAcceptRejectClaim.Show
End Sub
Private Sub cmdAddDeletePlan_Click()
    Me.Hide
    frmAddPlan.Show
End Sub
Private Sub cmdDeleteCustomerPolicy_Click()
    Me.Hide
    frmDeletePolicy.Show
End Sub
Private Sub cmdPolicyPaymentReport_Click()
    Me.Hide
    frmPolicyPaymentsReport.Show
End Sub
Private Sub Form_Load()
    'Positioning the form
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmMain.Show
End Sub
Private Sub mnuAcceptRejectClaim_Click()
    Me.Hide
    frmAcceptRejectClaim.Show
End Sub
Private Sub mnuAcceptRejectPolicyCancelClaim_Click()
    Me.Hide
    frmAcceptRejectlPolicyCancel.Show
End Sub
Private Sub mnuAddDeleteInsurancePlan_Click()
   Me.Hide
   frmAddDeletePlan.Show
End Sub
Private Sub mnuAcceptRejectPolicy_Click()
    Me.Hide
    frmAddRejectPolicy.Show
End Sub
Private Sub mnuAddnsurancePlan_Click()
    Me.Hide
    frmAddPlan.Show
End Sub
Private Sub mnuAgentReport_Click()
    Me.Hide
    frmAgentsReport.Show
End Sub
Private Sub mnuCreateAccount_Click()
    Me.Hide
    frmCreateAccount.Show
End Sub
Private Sub mnuDeleteCustomerPolicy_Click()
    Me.Hide
    frmDeletePolicy.Show
End Sub
Private Sub mnuDeleteInsurancePlan_Click()
    Me.Hide
    frmDeletePlan.Show
End Sub
Private Sub mnuPolicyClaim_Click()
    Me.Hide
    rptPolicyClaims.Show
End Sub
Private Sub mnuPolicyPaymentReport_Click()
    Me.Hide
    frmPolicyPaymentsReport.Show
End Sub
Private Sub mnuRegistered_Click()
    Me.Hide
    frmRegisteredCustomers.Show
End Sub
Private Sub mnuTypesOfAccountReport_Click()
    Me.Hide
    frmTypesOfAccount.Show
End Sub

Private Sub mnuUpdateDetails_Click()
    Me.Hide
    frmUpdateDetails.Show
End Sub

Private Sub mnuUserLogs_Click()
    Me.Hide
    rptUserLogs.Show
End Sub
