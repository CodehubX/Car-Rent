VERSION 5.00
Begin VB.Form frmAdmHome 
   BackColor       =   &H000080FF&
   Caption         =   "Home"
   ClientHeight    =   2835
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6225
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
   ScaleHeight     =   2835
   ScaleWidth      =   6225
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLogout 
      Caption         =   "Logout"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddCars 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add Cars"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdAddUsers 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add Users"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblAdminTitle 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "You are Logged In As Admin!"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
End
Attribute VB_Name = "frmAdmHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddCars_Click()
    frmAddCars.Show
End Sub

Private Sub cmdAddUsers_Click()
    frmUsers.Show
End Sub

Private Sub cmdLogout_Click()
    frmhome.Show
    Unload Me
End Sub
