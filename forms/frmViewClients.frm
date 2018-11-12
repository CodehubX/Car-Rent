VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmViewClients 
   BackColor       =   &H000080FF&
   Caption         =   "View Clients"
   ClientHeight    =   4335
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13785
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
   ScaleHeight     =   4335
   ScaleWidth      =   13785
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid dgClients 
      Bindings        =   "frmViewClients.frx":0000
      Height          =   2775
      Left            =   6600
      TabIndex        =   17
      Top             =   360
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   4895
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtDL 
      DataField       =   "dl_no"
      DataSource      =   "UsersAdo"
      Height          =   405
      Left            =   4320
      TabIndex        =   16
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox txtArea 
      DataField       =   "area"
      DataSource      =   "UsersAdo"
      Height          =   405
      Left            =   4320
      TabIndex        =   15
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox txtPhone 
      DataField       =   "phone_no"
      DataSource      =   "UsersAdo"
      Height          =   405
      Left            =   4320
      TabIndex        =   14
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox txtAdress 
      DataField       =   "address"
      DataSource      =   "UsersAdo"
      Height          =   405
      Left            =   4320
      TabIndex        =   13
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txtID 
      DataField       =   "passport_no"
      DataSource      =   "UsersAdo"
      Height          =   405
      Left            =   4320
      TabIndex        =   12
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtName 
      DataField       =   "client_name"
      DataSource      =   "UsersAdo"
      Height          =   405
      Left            =   4320
      TabIndex        =   11
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "&EXIT"
      Height          =   375
      Left            =   8160
      TabIndex        =   4
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&CLEAR"
      Height          =   375
      Left            =   6840
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&DELETE"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&ADD"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&SAVE"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   3480
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc UsersAdo 
      Height          =   375
      Left            =   4080
      Top             =   3480
      Width           =   2655
      _ExtentX        =   4683
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
      Connect         =   $"frmViewClients.frx":0017
      OLEDBString     =   $"frmViewClients.frx":00AA
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "clients"
      Caption         =   "CLIENTS"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lblDL 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "DL Number"
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label lblArea 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Area of Residence"
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label lblPhone 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Phone Number"
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label lblAddress 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Address"
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label lblID 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "ID/Passport Number"
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Name"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "frmViewClients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
UsersAdo.Recordset.AddNew
End Sub

Private Sub cmdCancel_Click()
txtName = ""
txtID = ""
txtAdress = ""
txtPhone = ""
txtArea = ""
txtDL = ""
End Sub

Private Sub cmdDel_Click()
Confirm = MsgBox("Are you sure you want to delete this record?", vbYesNo, "Deletion Confirmation")
If Confirm = vbYes Then
UsersAdo.Recordset.Delete
MsgBox "Record Deleted!", , "Message"
Else
MsgBox "Record Not Deleted!", , "Message"
End If
End Sub

Private Sub cmdExit_Click()
frmUserHome.Show
Unload Me
End Sub

Private Sub cmdSave_Click()
UsersAdo.Recordset.Fields("client_name") = txtName
UsersAdo.Recordset.Fields("passport_no") = txtID
UsersAdo.Recordset.Fields("address") = txtAdress
UsersAdo.Recordset.Fields("phone_no") = txtPhone
UsersAdo.Recordset.Fields("area") = txtArea
UsersAdo.Recordset.Fields("dl_no") = txtDL
UsersAdo.Recordset.Update
End Sub

