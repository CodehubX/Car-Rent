VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmUsers 
   BackColor       =   &H000080FF&
   Caption         =   "Add Users"
   ClientHeight    =   2445
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4635
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
   ScaleHeight     =   2445
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H000080FF&
      Caption         =   "BACK"
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H000080FF&
      Caption         =   "CLEAR"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc AdoSave 
      Height          =   375
      Left            =   840
      Top             =   2040
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Connect         =   $"frmUsers.frx":0000
      OLEDBString     =   $"frmUsers.frx":0093
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tbluser"
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
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H000080FF&
      Caption         =   "SAVE"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtPass 
      Height          =   405
      Left            =   1920
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox txtUser 
      Height          =   405
      Left            =   1920
      TabIndex        =   2
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label lblpassword 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "Password"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lbluser 
      BackColor       =   &H80000012&
      Caption         =   "Username"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
Unload Me
frmAdmHome.Show
End Sub

Private Sub cmdClear_Click()
    txtUser = ""
    txtPass = ""
End Sub

Private Sub cmdSave_Click()
    AdoSave.Refresh
    AdoSave.Recordset.AddNew
    AdoSave.Recordset.Fields("username") = txtUser
    AdoSave.Recordset.Fields("password") = txtPass
    AdoSave.Recordset.Update
    MsgBox "New User Added"
End Sub
