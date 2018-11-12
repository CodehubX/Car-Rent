VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmBook 
   BackColor       =   &H000080FF&
   Caption         =   " Book Vehicle"
   ClientHeight    =   9615
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14970
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9615
   ScaleWidth      =   14970
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc carAdo 
      Height          =   330
      Left            =   5880
      Top             =   4680
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   $"frmBook.frx":0000
      OLEDBString     =   $"frmBook.frx":0093
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from vehicles"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame frmbook 
      BackColor       =   &H000080FF&
      Caption         =   " Booking Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   2160
      TabIndex        =   4
      Top             =   1800
      Width           =   7335
      Begin MSComCtl2.DTPicker dtpick 
         Height          =   375
         Left            =   2640
         TabIndex        =   17
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Format          =   50135041
         CurrentDate     =   43412
      End
      Begin VB.CommandButton cmdCalc 
         Caption         =   "CALCULATE"
         Height          =   585
         Left            =   5280
         TabIndex        =   16
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtNumber 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2640
         TabIndex        =   15
         Top             =   1440
         Width           =   2055
      End
      Begin MSAdodcLib.Adodc bookAdo 
         Height          =   330
         Left            =   1080
         Top             =   2880
         Visible         =   0   'False
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
         Connect         =   $"frmBook.frx":0126
         OLEDBString     =   $"frmBook.frx":01B9
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "book"
         Caption         =   "Adodc3"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "CANCEL"
         CausesValidation=   0   'False
         Height          =   495
         Left            =   5280
         TabIndex        =   12
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "SAVE"
         Height          =   495
         Left            =   5400
         TabIndex        =   11
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtTotal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2640
         TabIndex        =   10
         Top             =   2400
         Width           =   2055
      End
      Begin VB.TextBox txtPerDay 
         DataField       =   "price"
         DataSource      =   "carAdo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2640
         TabIndex        =   9
         Top             =   1920
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker dtreturn 
         Height          =   375
         Left            =   2640
         TabIndex        =   18
         Top             =   960
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Format          =   50135041
         CurrentDate     =   43412
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         Caption         =   "Number of Days"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label lblTotal 
         Alignment       =   2  'Center
         Caption         =   "Total Price"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label lblpriceday 
         Alignment       =   2  'Center
         Caption         =   "Price/day"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label lblReturn 
         Alignment       =   2  'Center
         Caption         =   "Return Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label lblstart 
         Alignment       =   2  'Center
         Caption         =   "Pick Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Frame frmCar 
      BackColor       =   &H000080FF&
      Caption         =   "Vehicle Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5520
      TabIndex        =   2
      Top             =   840
      Width           =   6135
      Begin VB.ComboBox Combo1 
         DataSource      =   "carAdo"
         Height          =   465
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label lblplate 
         Alignment       =   2  'Center
         Caption         =   "Plate:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame frmClient 
      BackColor       =   &H000080FF&
      Caption         =   "Client Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   5295
      Begin VB.TextBox txtPass 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   13
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label lblID 
         Alignment       =   2  'Center
         Caption         =   "ID/Passport No:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalc_Click()
    txtNumber = (DateValue(dtreturn) - DateValue(dtpick))
    If txtNumber = 0 Then txtNumber = 1
    txtTotal(0) = txtNumber * txtPerDay(0)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    bookAdo.Refresh
    bookAdo.Recordset.AddNew
    bookAdo.Recordset.Fields("pass_no") = txtPass
    bookAdo.Recordset.Fields("plate") = Combo1
    bookAdo.Recordset.Fields("pick") = dtpick
    bookAdo.Recordset.Fields("return") = dtreturn
    bookAdo.Recordset.Fields("number") = txtNumber
    bookAdo.Recordset.Fields("price") = txtPerDay(0)
    bookAdo.Recordset.Fields("total") = txtTotal(0)
    bookAdo.Recordset.Update
    MsgBox "Car Booked"
    If txtPass = "" And Combo1 = "" Then
     MsgBox ("Cannot Save empty record")
    End If
End Sub

Private Sub Combo1_Click()
    If Not Combo1.Text = "" Then
    carAdo.RecordSource = "select distinct price from vehicles where vehicles.plate = '" & Combo1.Text & "';"
    carAdo.Refresh
    Set txtPerDay(0).DataSource = carAdo
    txtPerDay(0).Refresh
    End If
End Sub

Private Sub Form_Load()

    carAdo.Refresh
    With carAdo.Recordset
    Do Until .EOF
    Combo1.AddItem ![plate]
    .MoveNext
    Loop
    End With
End Sub


