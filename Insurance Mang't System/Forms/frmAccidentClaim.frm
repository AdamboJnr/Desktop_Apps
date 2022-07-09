VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAccidentClaim 
   Caption         =   "Form1"
   ClientHeight    =   6690
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   7005
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc dtaAccidentClaim 
      Height          =   375
      Left            =   240
      Top             =   5760
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Insurance management database.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Insurance management database.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from tblAccidentInsuranceClaim"
      Caption         =   "Accident Claim"
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
      Left            =   5520
      Picture         =   "frmAccidentClaim.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5520
      Width           =   975
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
      Height          =   855
      Left            =   4200
      Picture         =   "frmAccidentClaim.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5520
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   4095
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   6615
      Begin VB.TextBox txtDetails 
         Height          =   375
         Left            =   2400
         TabIndex        =   11
         Top             =   3360
         Width           =   2895
      End
      Begin VB.TextBox txtPlaceOccured 
         DataSource      =   "dtaAccidentClaim"
         Height          =   375
         Left            =   2400
         TabIndex        =   10
         Top             =   2640
         Width           =   2895
      End
      Begin VB.TextBox txtDateOccured 
         DataSource      =   "dtaAccidentClaim"
         Height          =   375
         Left            =   2400
         TabIndex        =   9
         Top             =   1920
         Width           =   2895
      End
      Begin VB.TextBox txtClaimantName 
         Height          =   375
         Left            =   2400
         TabIndex        =   8
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox txtPolicyNumber 
         Height          =   375
         Left            =   2400
         TabIndex        =   7
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Details Of Injury"
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
         TabIndex        =   6
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Place Occurred"
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
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Date Occurred"
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
         TabIndex        =   4
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Claimant Name"
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
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lblPolicy 
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
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   5520
      Picture         =   "frmAccidentClaim.frx":0884
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Accident Insurance Claim"
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
      TabIndex        =   0
      Top             =   360
      Width           =   4695
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   5655
   End
End
Attribute VB_Name = "frmAccidentClaim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    If txtClaimantName.Text = "" And txtDateOccured.Text = "" And txtDetails.Text = "" And txtPlaceOccured.Text = "" And txtPolicyNumber.Text = "" Then
        Unload Me
        frmMain.Show
    Else
        txtHolderName.Text = ""
        txtDateOccured.Text = ""
        txtDescription.Text = ""
        txtTypeOfPremise.Text = ""
        txtPlaceOccurred.Text = ""
        txtPolicyNumber.Text = ""
    End If
End Sub

Private Sub cmdSave_Click()
    Dim curBuildingCost As Currency, curContentsCost As Currency
    Dim PolicyNumber As Long
    
    If txtClaimantName.Text = "" Or txtDateOccured.Text = "" Or txtDetails.Text = "" Or txtPolicyNumber.Text = "" Or txtPlaceOccured.Text = "" Then
        MsgBox "Please Fill In All Inputs"
    Else
        PolicyNumber = txtPolicyNumber.Text
        dtaAccidentClaim.Recordset.AddNew
        dtaAccidentClaim.Recordset.Fields(0).Value = PolicyNumber
        dtaAccidentClaim.Recordset.Fields(1).Value = txtClaimantName.Text
        dtaAccidentClaim.Recordset.Fields(2).Value = txtDateOccured.Text
        dtaAccidentClaim.Recordset.Fields(3).Value = txtPlaceOccured.Text
        dtaAccidentClaim.Recordset.Fields(4).Value = txtDetails.Text
        dtaAccidentClaim.Recordset.Update
        
        S = MsgBox("Claim Succesfully Submitted. Claim Number -" & ClaimNumber, vbInformation)
        'Clearing all inputs
        txtClaimantName.Text = ""
        txtDateOccured.Text = ""
        txtDetails.Text = ""
        txtPlaceOccured.Text = ""
        txtPolicyNumber.Text = ""
    End If
End Sub
Private Sub Form_Load()
    'Positioning the form
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
    
    txtPolicyNumber.Text = ClaimantPolicyNumber
    txtClaimantName.Text = ClaimantName
End Sub
Private Sub txtDateOccured_Validate(Cancel As Boolean)
    If IsDate(txtDateOccured.Text) = False Then
        MsgBox "Key in Valid Date In The Form of 8/7/2021", vbCritical
        txtDateOccured.Text = ""
        txtDateOccured.SetFocus
    End If
End Sub
