VERSION 5.00
Begin VB.Form Disclaimer 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   Caption         =   "Disclaimer of Warranty and Liability"
   ClientHeight    =   4425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Rx.ocxFormShape ocxFormShape1 
      Left            =   120
      Top             =   3960
      _ExtentX        =   794
      _ExtentY        =   873
      Shape           =   4
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000E&
      Caption         =   "Decline"
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000E&
      Caption         =   "Accept"
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
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000E&
      Height          =   3675
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Disclaimer.frx":0000
      Top             =   240
      Width           =   8895
   End
End
Attribute VB_Name = "Disclaimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Open App.Path & "\Accept" For Output As #1: Close #1
    Unload Me
    Load Loading
    Loading.Show
End Sub

Private Sub Command2_Click()
On Error Resume Next
Kill App.Path & "\Accept"
Unload Me
End Sub

Private Sub Form_Deactivate()
SetTopMostWindow Me.hwnd, True
End Sub

Private Sub Form_Load()
    SetTopMostWindow Me.hwnd, True
    If Dir(App.Path & "\Accept") <> "" Then
        Unload Me
        Load Loading
        Loading.Show
        Set Disclaimer = Nothing
    End If
End Sub

Private Sub Text1_GotFocus()
Command1.SetFocus
End Sub
