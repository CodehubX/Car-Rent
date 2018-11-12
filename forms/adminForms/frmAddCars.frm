VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAddCars 
   BackColor       =   &H000080FF&
   Caption         =   "Add Cars"
   ClientHeight    =   4095
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5235
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   5235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Cancel          =   -1  'True
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
      Left            =   1920
      TabIndex        =   12
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
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
      Left            =   3840
      TabIndex        =   11
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "SAVE"
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
      Left            =   240
      TabIndex        =   10
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox txtPrice 
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox txtSize 
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox txtCapacity 
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtPlate 
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox txtMake 
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   240
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc carAdo 
      Height          =   375
      Left            =   1440
      Top             =   3720
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Connect         =   $"frmAddCars.frx":0000
      OLEDBString     =   $"frmAddCars.frx":0093
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "vehicles"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lblprice 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Price/Day"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label lblsize 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Size of car eg  4WD"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label lblseats 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Seating Capacity"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label lblplate 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Plate Number"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lblmake 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Make/Type"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "frmAddCars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    txtMake = ""
    txtPlate = ""
    txtCapacity = ""
    txtSize = ""
    txtPrice = ""
End Sub

Private Sub cmdSave_Click()
    carAdo.Refresh
    carAdo.Recordset.AddNew
    carAdo.Recordset.Fields("make") = txtMake
    carAdo.Recordset.Fields("plate") = txtPlate
    carAdo.Recordset.Fields("capacity") = txtCapacity
    carAdo.Recordset.Fields("size") = txtSize
    carAdo.Recordset.Fields("price") = txtPrice
    carAdo.Recordset.Update
    MsgBox "New Vehicle Added"
End Sub
