VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmFirePerils 
   Caption         =   "Fire and Perils Insurance"
   ClientHeight    =   5940
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   7125
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc dtaFireAndPerils 
      Height          =   375
      Left            =   240
      Top             =   5040
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      RecordSource    =   "select * from  tblFireAndPerils"
      Caption         =   "Fire And Perils"
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
      Cancel          =   -1  'True
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
      Height          =   975
      Left            =   5640
      Picture         =   "frmFirePerils.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4800
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
      Height          =   975
      Left            =   4200
      Picture         =   "frmFirePerils.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Frame fraBuildingDetails 
      Height          =   3375
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   6255
      Begin VB.TextBox txtPolicyNumber 
         Height          =   375
         Left            =   2640
         TabIndex        =   9
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox txtDescription 
         DataSource      =   "dtaFireAndPerils"
         Height          =   375
         Left            =   2640
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   2640
         Width           =   2895
      End
      Begin VB.TextBox txtContentsCost 
         Height          =   375
         Left            =   2640
         TabIndex        =   5
         Top             =   1920
         Width           =   2895
      End
      Begin VB.TextBox txtBuildingCost 
         DataSource      =   "dtaFireAndPerils"
         Height          =   375
         Left            =   2640
         TabIndex        =   3
         Top             =   1200
         Width           =   2895
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
         Left            =   360
         TabIndex        =   8
         Top             =   480
         Width           =   1935
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
         Left            =   360
         TabIndex        =   6
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label lblContentsCost 
         Alignment       =   1  'Right Justify
         Caption         =   "Contents Cost"
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
         Left            =   360
         TabIndex        =   4
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label lblBuildingSumAssured 
         Alignment       =   1  'Right Justify
         Caption         =   "Building Cost"
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
         Left            =   360
         TabIndex        =   2
         Top             =   1200
         Width           =   1935
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   5520
      Picture         =   "frmFirePerils.frx":0884
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Fire And Perils Insurance"
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
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   4215
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   960
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   5295
   End
End
Attribute VB_Name = "frmFirePerils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    If txtBuildingCost.Text = "" And txtContentsCost.Text = "" And txtDescription.Text = "" And txtPolicyNumber.Text = "" Then
        Unload Me
        frmPolicyCreationForm.Show
    Else
        txtBuildingCost.Text = ""
        txtContentsCost.Text = ""
        txtDescription.Text = ""
        txtPolicyNumber.Text = ""
    End If
End Sub
Private Sub cmdSave_Click()
    Dim curBuildingCost As Currency, curContentsCost As Currency
    Dim PolicyNumber As Long
    
    If txtBuildingCost.Text = "" Or txtContentsCost.Text = "" Or txtDescription.Text = "" Or txtPolicyNumber.Text = "" Then
        MsgBox "Please Fill In All Inputs"
    Else
        curBuildingCost = txtBuildingCost.Text
        curContentsCost = txtContentsCost.Text
        PolicyNumber = txtPolicyNumber.Text
        
        dtaFireAndPerils.Recordset.AddNew
        dtaFireAndPerils.Recordset.Fields(0).Value = PolicyNumber
        dtaFireAndPerils.Recordset.Fields(1).Value = curBuildingCost
        dtaFireAndPerils.Recordset.Fields(2).Value = curContentsCost
        dtaFireAndPerils.Recordset.Fields(3).Value = txtDescription.Text
        dtaFireAndPerils.Recordset.Update
        
        MsgBox "Updated Succesfully", vbInformation
        txtBuildingCost.Text = ""
        txtContentsCost.Text = ""
        txtDescription.Text = ""
        txtPolicyNumber.Text = ""
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
    frmPolicyCreationForm.Show
End Sub
