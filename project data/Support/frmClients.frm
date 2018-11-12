VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmClients 
   BackColor       =   &H000080FF&
   Caption         =   "Add Clients"
   ClientHeight    =   4875
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6465
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   14
      Top             =   3240
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc addAdo 
      Height          =   375
      Left            =   1440
      Top             =   3960
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   $"frmClients.frx":0000
      OLEDBString     =   $"frmClients.frx":0093
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "clients"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
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
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   13
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   12
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox txtArea 
      Height          =   405
      Left            =   2760
      TabIndex        =   11
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox txtPhone 
      Height          =   405
      Left            =   2760
      TabIndex        =   10
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox txtAdress 
      Height          =   405
      Left            =   2760
      TabIndex        =   9
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox txtDL 
      Height          =   405
      Left            =   2760
      TabIndex        =   8
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox txtID 
      Height          =   405
      Left            =   2760
      TabIndex        =   7
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox txtName 
      Height          =   405
      Left            =   2760
      TabIndex        =   6
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblArea 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Area of Residence"
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label lblPhone 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Phone Number"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lblAdress 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Postal Adress"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblDL 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "DL Number"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label lblID 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Id/Passport No"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Name"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmClients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
    addAdo.Refresh
    addAdo.Recordset.AddNew
    addAdo.Recordset.Fields("client_name") = txtName
    addAdo.Recordset.Fields("passport_no") = txtID
    addAdo.Recordset.Fields("address") = txtAdress
    addAdo.Recordset.Fields("phone_no") = txtPhone
    addAdo.Recordset.Fields("area") = txtArea
    addAdo.Recordset.Fields("dl_no") = txtDL
    addAdo.Recordset.Update
    MsgBox "New Client Added"
End Sub

Private Sub cmdCancel_Click()
Unload Me
frmUserHome.Show
End Sub

Private Sub cmdClear_Click()
txtName = ""
txtID = ""
txtAdress = ""
txtPhone = ""
txtArea = ""
txtDL = ""
End Sub
