VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmFireAndPerilsClaim 
   Caption         =   "Fire And Perils"
   ClientHeight    =   6495
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   6240
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc dtaFireAndPerilsClaim 
      Height          =   375
      Left            =   480
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
      RecordSource    =   "select * from tblFireAndPerilsClaim"
      Caption         =   "Fire And Perils Claim"
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
      Left            =   5040
      Picture         =   "frmFireAndPerilsClaim.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
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
      Left            =   3720
      Picture         =   "frmFireAndPerilsClaim.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5520
      Width           =   975
   End
   Begin VB.Frame fraFireAndPerilsDetails 
      Height          =   4095
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   5895
      Begin VB.TextBox txtDescription 
         Height          =   375
         Left            =   2280
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   3360
         Width           =   3015
      End
      Begin VB.TextBox txtTypeOfPremise 
         DataSource      =   "dtaFireAndPerilsClaim"
         Height          =   375
         Left            =   2280
         TabIndex        =   11
         Top             =   2760
         Width           =   3015
      End
      Begin VB.TextBox txtPlaceOccurred 
         Height          =   375
         Left            =   2280
         TabIndex        =   10
         Top             =   2160
         Width           =   3015
      End
      Begin VB.TextBox txtDateOccurred 
         Height          =   375
         Left            =   2280
         TabIndex        =   9
         Top             =   1560
         Width           =   3015
      End
      Begin VB.TextBox txtHolderName 
         Height          =   375
         Left            =   2280
         TabIndex        =   8
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox txtPolicyNumber 
         Height          =   375
         Left            =   2280
         TabIndex        =   7
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label lblDescription 
         Alignment       =   1  'Right Justify
         Caption         =   "Description"
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
         Left            =   240
         TabIndex        =   12
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label lblTypeOfPremiseDamaged 
         Alignment       =   1  'Right Justify
         Caption         =   "Type Of Premise"
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
         Left            =   240
         TabIndex        =   6
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label lblPlaceOccurred 
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
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label lblDateOccured 
         Alignment       =   1  'Right Justify
         Caption         =   "Date Occured"
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
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label lblHolderName 
         Alignment       =   1  'Right Justify
         Caption         =   "Holder Name"
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
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblPolicyNumber 
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
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   4560
      Picture         =   "frmFireAndPerilsClaim.frx":0884
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Fire And Perils Claim"
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
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   840
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "frmFireAndPerilsClaim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    If txtHolderName.Text = "" And txtDateOccured.Text = "" And txtDescription.Text = "" And txtTypeOfPremise.Text = "" And txtPlaceOccured.Text = "" And txtPolicyNumber.Text = "" Then
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
    Dim PolicyNumber As Long
    
    PolicyNumber = txtPolicyNumber.Text
    
    'Saving to Database
    If txtDateOccurred.Text = "" Or txtDescription.Text = "" Or txtHolderName.Text = "" Or txtPlaceOccurred.Text = "" Or txtPolicyNumber.Text = "" Or txtTypeOfPremise.Text = "" Then
        MsgBox "Please Fill In All Inputs", vbCritical
    Else
        dtaFireAndPerilsClaim.Recordset.AddNew
        dtaFireAndPerilsClaim.Recordset.Fields(0).Value = PolicyNumber
        dtaFireAndPerilsClaim.Recordset.Fields(1).Value = txtHolderName.Text
        dtaFireAndPerilsClaim.Recordset.Fields(2).Value = txtDateOccurred
        dtaFireAndPerilsClaim.Recordset.Fields(3).Value = txtPlaceOccurred.Text
        dtaFireAndPerilsClaim.Recordset.Fields(4).Value = txtTypeOfPremise.Text
        dtaFireAndPerilsClaim.Recordset.Fields(5).Value = txtDescription.Text
        dtaFireAndPerilsClaim.Recordset.Update
        S = MsgBox("Claim Succesfully Submitted. Claim Number -" & ClaimNumber, vbInformation)
     End If
    'Clearing all inputs
    txtHolderName.Text = ""
    txtDateOccurred.Text = ""
    txtDescription.Text = ""
    txtTypeOfPremise.Text = ""
    txtPlaceOccurred.Text = ""
    txtPolicyNumber.Text = ""
End Sub
Private Sub Form_Load()
    'Positioning the form
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
    
    txtPolicyNumber.Text = ClaimantPolicyNumber
    txtHolderName.Text = ClaimantName
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
    frmMain.Show
End Sub

Private Sub txtDateOccurred_Validate(Cancel As Boolean)
    If IsDate(txtDateOccurred.Text) = False Then
        MsgBox "Key in Valid Date In The Form of 8/7/2021", vbCritical
        txtDateOccured.Text = ""
        txtDateOccured.SetFocus
    End If
End Sub
