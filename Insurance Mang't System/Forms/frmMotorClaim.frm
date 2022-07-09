VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmMotorClaim 
   Caption         =   "Motor Claim"
   ClientHeight    =   7425
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   10860
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc dtaVehicleDetails 
      Height          =   375
      Left            =   3840
      Top             =   6600
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
      RecordSource    =   "select * from tblMotorInsurance"
      Caption         =   "Vehicle Details"
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
   Begin MSAdodcLib.Adodc dtaMotorClaim 
      Height          =   375
      Left            =   480
      Top             =   6600
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
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
      RecordSource    =   "Select * from tblMotorClaims"
      Caption         =   "Motor Claim"
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
      Left            =   9360
      Picture         =   "frmMotorClaim.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6360
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
      Left            =   7920
      Picture         =   "frmMotorClaim.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Frame fraMotorInsuranceDetails 
      Height          =   4815
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   10215
      Begin VB.Frame fraAccidentDetails 
         Caption         =   "Accident Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   5040
         TabIndex        =   14
         Top             =   240
         Width           =   4815
         Begin VB.TextBox txtDescription 
            Height          =   375
            Left            =   2040
            MultiLine       =   -1  'True
            TabIndex        =   19
            Top             =   2160
            Width           =   2415
         End
         Begin VB.TextBox txtDateOccured 
            DataSource      =   "dtaMotorClaim"
            Height          =   375
            Left            =   2040
            TabIndex        =   18
            Top             =   1320
            Width           =   2415
         End
         Begin VB.TextBox txtPlaceOccured 
            Height          =   375
            Left            =   2040
            TabIndex        =   16
            Top             =   480
            Width           =   2415
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
            Height          =   495
            Left            =   240
            TabIndex        =   20
            Top             =   2160
            Width           =   1455
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
            Height          =   495
            Left            =   240
            TabIndex        =   17
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label lblPlaceOccured 
            Alignment       =   1  'Right Justify
            Caption         =   "Place Occured"
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
            TabIndex        =   15
            Top             =   480
            Width           =   1455
         End
      End
      Begin VB.Frame fraVehicleDetails 
         Caption         =   "Vehicle Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   240
         TabIndex        =   9
         Top             =   2880
         Width           =   4455
         Begin VB.TextBox txtVehicleMake 
            Height          =   375
            Left            =   1920
            TabIndex        =   13
            Top             =   1080
            Width           =   2175
         End
         Begin VB.TextBox txtRegNumber 
            DataSource      =   "dtaVehicleDetails"
            Height          =   375
            Left            =   1920
            TabIndex        =   10
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label lblMake 
            Alignment       =   1  'Right Justify
            Caption         =   "Vehicle Make"
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
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label lblRegNumber 
            Alignment       =   1  'Right Justify
            Caption         =   "Reg Number"
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
            TabIndex        =   11
            Top             =   480
            Width           =   1455
         End
      End
      Begin VB.Frame fraPersonalDetails 
         Caption         =   "Personal Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   4455
         Begin VB.TextBox txtOccupation 
            Height          =   375
            Left            =   1920
            TabIndex        =   8
            Top             =   1680
            Width           =   2175
         End
         Begin VB.TextBox txtClaimantName 
            Height          =   375
            Left            =   1920
            TabIndex        =   6
            Top             =   1080
            Width           =   2175
         End
         Begin VB.TextBox txtPolicyNumber 
            Height          =   375
            Left            =   1920
            TabIndex        =   4
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label lblOccupation 
            Alignment       =   1  'Right Justify
            Caption         =   "Occupation"
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
            TabIndex        =   7
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label lblClaimantName 
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
            Height          =   375
            Left            =   240
            TabIndex        =   5
            Top             =   1080
            Width           =   1455
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
            TabIndex        =   3
            Top             =   480
            Width           =   1455
         End
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   7440
      Picture         =   "frmMotorClaim.frx":0884
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Motor Insurance Claim"
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
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   5055
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   1920
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   6375
   End
End
Attribute VB_Name = "frmMotorClaim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    If txtClaimantName.Text = "" And txtDateOccured.Text = "" And txtDescription.Text = "" And txtOccupation.Text = "" And txtPlaceOccured.Text = "" And txtPolicyNumber.Text = "" And txtRegNumber.Text = "" And txtVehicleMake.Text Then
        Unload Me
        frmMain.Show
    Else
        txtClaimantName.Text = ""
        txtDateOccured.Text = ""
        txtDescription.Text = ""
        txtOccupation.Text = ""
        txtPlaceOccured.Text = ""
        txtPolicyNumber.Text = ""
        txtRegNumber.Text = ""
        txtVehicleMake.Text = ""
    End If
End Sub
Private Sub cmdSave_Click()
    Dim PolicyNumber As Long
    
    PolicyNumber = txtPolicyNumber.Text
    
    'Saving to Database
    If txtClaimantName.Text = "" Or txtDateOccured.Text = "" Or txtDescription.Text = "" Or txtOccupation.Text = "" Or txtPlaceOccured.Text = "" Or txtPolicyNumber.Text = "" Or txtRegNumber.Text = "" Or txtVehicleMake.Text = "" Then
        MsgBox "Please Fill In All Inputs", vbCritical
    Else
        dtaMotorClaim.Recordset.AddNew
        dtaMotorClaim.Recordset.Fields(0).Value = PolicyNumber
        dtaMotorClaim.Recordset.Fields(1).Value = txtClaimantName.Text
        dtaMotorClaim.Recordset.Fields(2).Value = txtOccupation.Text
        dtaMotorClaim.Recordset.Fields(3).Value = txtRegNumber.Text
        dtaMotorClaim.Recordset.Fields(4).Value = txtVehicleMake.Text
        dtaMotorClaim.Recordset.Fields(5).Value = txtPlaceOccured.Text
        dtaMotorClaim.Recordset.Fields(6).Value = txtDateOccured.Text
        dtaMotorClaim.Recordset.Fields(7).Value = txtDescription.Text
        dtaMotorClaim.Recordset.Update
        
        S = MsgBox("Claim Succesfully Submitted. Claim Number -" & ClaimNumber, vbInformation)
        
        'Clearing all inputs
        txtClaimantName.Text = ""
        txtDateOccured.Text = ""
        txtDescription.Text = ""
        txtOccupation.Text = ""
        txtPlaceOccured.Text = ""
        txtPolicyNumber.Text = ""
        txtRegNumber.Text = ""
        txtVehicleMake.Text = ""
     End If
End Sub
Private Sub Form_Load()
    'Positioning the form
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
    
    txtPolicyNumber.Text = ClaimantPolicyNumber
    txtClaimantName.Text = ClaimantName
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
    frmMain.Show
End Sub
Private Sub txtDateOccured_Validate(Cancel As Boolean)
    If IsDate(txtDateOccured.Text) = False Then
        MsgBox "Key in Valid Date In The Form of 8/7/2021", vbCritical
        txtDateOccured.Text = ""
        txtDateOccured.SetFocus
    End If
End Sub
Private Sub txtPolicyNumber_LostFocus()
    Dim PolicyNumber As Long
    
    PolicyNumber = txtPolicyNumber.Text
    'Automatically filling in the Insurance Name
    dtaVehicleDetails.Recordset.MoveFirst
    dtaVehicleDetails.Recordset.Find "[Policy Number]= " & PolicyNumber, 0, adSearchForward
    If dtaVehicleDetails.Recordset.EOF = True Then
        MsgBox "Record Not Found"
        dtaVehicleDetails.Recordset.MoveFirst
    ElseIf dtaVehicleDetails.Recordset.Fields(0).Value = PolicyNumber Then
        txtRegNumber.Text = dtaVehicleDetails.Recordset.Fields(1).Value
        txtVehicleMake.Text = dtaVehicleDetails.Recordset.Fields(2).Value
    End If
End Sub
