VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmMotorInsurance 
   Caption         =   "Motor Insurance"
   ClientHeight    =   7755
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   7140
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc dtaMotorInsurance 
      Height          =   375
      Left            =   480
      Top             =   6840
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "Motor Insurance"
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
      Left            =   5760
      Picture         =   "frmMotorInsurance.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6720
      Width           =   1095
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
      Left            =   4320
      Picture         =   "frmMotorInsurance.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Frame fraMotorInsuranceDetails 
      Height          =   5295
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   6615
      Begin VB.ComboBox cboYearOfManufacture 
         Height          =   315
         ItemData        =   "frmMotorInsurance.frx":0884
         Left            =   2880
         List            =   "frmMotorInsurance.frx":08AC
         TabIndex        =   19
         Top             =   2760
         Width           =   2775
      End
      Begin VB.ComboBox cboTypeOfVehicle 
         Height          =   315
         ItemData        =   "frmMotorInsurance.frx":08F8
         Left            =   2880
         List            =   "frmMotorInsurance.frx":090E
         TabIndex        =   16
         Top             =   3960
         Width           =   2775
      End
      Begin VB.TextBox txtVehicleCost 
         Height          =   375
         Left            =   2880
         TabIndex        =   14
         Top             =   4560
         Width           =   2775
      End
      Begin VB.TextBox txtSeatingCapacity 
         Height          =   375
         Left            =   2880
         TabIndex        =   12
         Top             =   3360
         Width           =   2775
      End
      Begin VB.TextBox txtMotorModel 
         Height          =   375
         Left            =   2880
         TabIndex        =   9
         Top             =   2160
         Width           =   2775
      End
      Begin VB.TextBox txtVehicleMake 
         DataSource      =   "dtaMotorInsurance"
         Height          =   375
         Left            =   2880
         TabIndex        =   7
         Top             =   1560
         Width           =   2775
      End
      Begin VB.TextBox txtRegistrationNumber 
         Height          =   375
         Left            =   2880
         TabIndex        =   4
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtPolicyNumber 
         Height          =   375
         Left            =   2880
         TabIndex        =   3
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label lblCost 
         Alignment       =   1  'Right Justify
         Caption         =   "Cost"
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
         TabIndex        =   15
         Top             =   4560
         Width           =   1935
      End
      Begin VB.Label lblTypeOfVehicle 
         Alignment       =   1  'Right Justify
         Caption         =   "Type Of Vehicle"
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
         TabIndex        =   13
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label lblSeatingCapacity 
         Alignment       =   1  'Right Justify
         Caption         =   "Seating Capacity"
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
         TabIndex        =   11
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label lblYearOfManufacture 
         Alignment       =   1  'Right Justify
         Caption         =   "Year Of Manufacture"
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
         TabIndex        =   10
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label lblMotorModel 
         Alignment       =   1  'Right Justify
         Caption         =   "Motor Model"
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
         TabIndex        =   8
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label lblVehicleMake 
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
         Left            =   480
         TabIndex        =   6
         Top             =   1560
         Width           =   1935
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
         Left            =   480
         TabIndex        =   5
         Top             =   960
         Width           =   1935
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
         Left            =   480
         TabIndex        =   2
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   5520
      Picture         =   "frmMotorInsurance.frx":094C
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Motor Insurance"
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
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   5535
   End
End
Attribute VB_Name = "frmMotorInsurance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
    frmMain.Show
End Sub
Private Sub cmdSave_Click()
    Dim ManufacturedYear As Long
    Dim SeatingCapacity As Long
    Dim cost As Currency
    If txtMotorModel.Text = "" Or txtPolicyNumber.Text = "" Or txtRegistrationNumber.Text = "" Or txtSeatingCapacity.Text = "" Or txtVehicleCost.Text = "" Or txtVehicleMake.Text = "" Or cboTypeOfVehicle.Text = "" Or cboYearOfManufacture.Text = "" Then
        MsgBox "Please Fill In All Inputs", vbInformation
    Else
        ManufacturedYear = cboYearOfManufacture
        SeatingCapacity = txtSeatingCapacity.Text
        cost = txtVehicleCost.Text
        
        dtaMotorInsurance.Recordset.AddNew
        dtaMotorInsurance.Recordset.Fields(0).Value = txtPolicyNumber.Text 'lngPolicyNumber
        dtaMotorInsurance.Recordset.Fields(1).Value = txtRegistrationNumber.Text
        dtaMotorInsurance.Recordset.Fields(2).Value = txtVehicleMake.Text
        dtaMotorInsurance.Recordset.Fields(3).Value = txtMotorModel.Text
        dtaMotorInsurance.Recordset.Fields(4).Value = ManufacturedYear
        dtaMotorInsurance.Recordset.Fields(5).Value = SeatingCapacity
        dtaMotorInsurance.Recordset.Fields(6).Value = cboTypeOfVehicle.Text
        dtaMotorInsurance.Recordset.Fields(7).Value = cost
        dtaMotorInsurance.Recordset.Update
        MsgBox "Record Updated Succesfully", vbInformation
        txtMotorModel.Text = ""
        txtPolicyNumber.Text = ""
        txtRegistrationNumber.Text = ""
        txtSeatingCapacity.Text = ""
        txtVehicleCost.Text = ""
        txtVehicleMake.Text = ""
        cboTypeOfVehicle.Text = ""
        cboYearOfManufacture.Text = ""
    End If
End Sub
Private Sub Form_Load()
    'Positioning the form
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
    'Automatically Filling the Policy Number
    txtPolicyNumber.Text = lngPolicyNumber
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
    frmMain.Show
End Sub
