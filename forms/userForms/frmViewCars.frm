VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmViewCars 
   BackColor       =   &H000080FF&
   Caption         =   "View Vehicles"
   ClientHeight    =   3810
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14220
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
   ScaleHeight     =   3810
   ScaleWidth      =   14220
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&ADD"
      Height          =   375
      Left            =   1440
      TabIndex        =   15
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&DELETE"
      Height          =   375
      Left            =   2640
      TabIndex        =   14
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&CLEAR"
      Height          =   375
      Left            =   6600
      TabIndex        =   13
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&EXIT"
      Height          =   375
      Left            =   7800
      TabIndex        =   12
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&SAVE"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   3240
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid dgVehicles 
      Bindings        =   "frmViewCars.frx":0000
      Height          =   2295
      Left            =   5520
      TabIndex        =   10
      Top             =   360
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
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
   Begin MSAdodcLib.Adodc ViewAdo 
      Height          =   330
      Left            =   3840
      Top             =   3240
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
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
      Connect         =   $"frmViewCars.frx":0016
      OLEDBString     =   $"frmViewCars.frx":00A9
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "vehicles"
      Caption         =   "VEHICLES"
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
   Begin VB.TextBox txtPrice 
      DataField       =   "price"
      DataSource      =   "ViewAdo"
      Height          =   405
      Left            =   3120
      TabIndex        =   9
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox txtSize 
      DataField       =   "size"
      DataSource      =   "ViewAdo"
      Height          =   405
      Left            =   3120
      TabIndex        =   8
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox txtSeat 
      DataField       =   "capacity"
      DataSource      =   "ViewAdo"
      Height          =   405
      Left            =   3120
      TabIndex        =   7
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox txtPlate 
      DataField       =   "plate"
      DataSource      =   "ViewAdo"
      Height          =   405
      Left            =   3120
      TabIndex        =   6
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox txtMake 
      DataField       =   "make"
      DataSource      =   "ViewAdo"
      Height          =   405
      Left            =   3120
      TabIndex        =   5
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label lblPrice 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Price/Day"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label lblSize 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Size of car"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label lblSeat 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Seating Capacity"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label lblplate 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Plate Number"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lblmake 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Make/Type"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "frmViewCars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
ViewAdo.Recordset.AddNew
End Sub

Private Sub cmdCancel_Click()
txtMake = ""
txtPlate = ""
txtSeat = ""
txtSize = ""
txtPrice = ""
End Sub

Private Sub cmdDel_Click()
Confirm = MsgBox("Are you sure you want to delete this record?", vbYesNo, "Deletion Confirmation")
If Confirm = vbYes Then
ViewAdo.Recordset.Delete
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
ViewAdo.Recordset.Fields("make") = txtMake
ViewAdo.Recordset.Fields("plate") = txtPlate
ViewAdo.Recordset.Fields("capacity") = txtSeat
ViewAdo.Recordset.Fields("size") = txtSize
ViewAdo.Recordset.Fields("price") = txtPrice
ViewAdo.Recordset.Update
End Sub
