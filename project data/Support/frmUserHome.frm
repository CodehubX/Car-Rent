VERSION 5.00
Begin VB.Form frmUserHome 
   BackColor       =   &H000080FF&
   Caption         =   "Home"
   ClientHeight    =   4065
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5805
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
   ScaleHeight     =   4065
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBooked 
      Caption         =   "Booking Records"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1680
      TabIndex        =   6
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CommandButton cmdBookVehicles 
      Caption         =   "Book Vehicle"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   5
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdCars 
      Caption         =   "View Cars"
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton cmdClients 
      Caption         =   "View Clients"
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdClient 
      Caption         =   "Add Clients"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdLogout 
      Caption         =   "Logout"
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lbltitle 
      BackColor       =   &H000080FF&
      Caption         =   "You are Logged In As User!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "frmUserHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBooked_Click()
    frmRecords.Show
End Sub

Private Sub cmdBookVehicles_Click()
    frmbook.Show
End Sub

Private Sub cmdCars_Click()
    frmViewCars.Show
End Sub

Private Sub cmdClient_Click()
    frmClients.Show
End Sub

Private Sub cmdClients_Click()
    frmViewClients.Show
End Sub

Private Sub cmdLogout_Click()
    frmhome.Show
    Unload Me
End Sub
