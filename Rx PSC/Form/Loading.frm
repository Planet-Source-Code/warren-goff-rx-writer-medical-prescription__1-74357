VERSION 5.00
Begin VB.Form Loading 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Loading.frx":0000
   ScaleHeight     =   930
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Rx.ocxFormShape ocxFormShape1 
      Left            =   4200
      Top             =   720
      _ExtentX        =   794
      _ExtentY        =   873
      Shape           =   7
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   3600
      Picture         =   "Loading.frx":5681
      Top             =   270
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   330
      Picture         =   "Loading.frx":5F4B
      Top             =   285
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Data...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   585
      TabIndex        =   0
      Top             =   240
      Width           =   3360
   End
End
Attribute VB_Name = "Loading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
DoEvents
    Delay 0.25
    Load Form1
    Form1.Show
End Sub

Private Sub Form_Load()
    SetTopMostWindow Me.hwnd, True
End Sub

