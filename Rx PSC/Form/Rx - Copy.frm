VERSION 5.00
Begin VB.Form RxForm2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7335
   Icon            =   "Rx - Copy.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleMode       =   0  'User
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pt Educ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   9
      Top             =   6000
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.TextBox Sig1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   720
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   4440
      Width           =   6615
   End
   Begin VB.TextBox Drug1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   1680
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   3000
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7920
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1710
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7905
      Width           =   1095
   End
   Begin VB.TextBox Comment 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   6360
      Width           =   4815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Medication:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   18
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15
      TabIndex        =   17
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   15
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label lblText 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Dosage and Sig:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Substitute"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Refills"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label Num1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   3120
      TabIndex        =   5
      Top             =   3480
      Width           =   795
   End
   Begin VB.Label Ref1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   1200
      TabIndex        =   4
      Top             =   5160
      Width           =   660
   End
   Begin VB.Label Sub1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   1200
      TabIndex        =   3
      Top             =   5520
      Width           =   660
   End
   Begin VB.Label Named 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   960
      TabIndex        =   2
      Top             =   1500
      Width           =   3495
   End
   Begin VB.Label Aged 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4920
      TabIndex        =   1
      Top             =   1560
      Width           =   300
   End
   Begin VB.Label Dated 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5640
      TabIndex        =   0
      Top             =   1560
      Width           =   780
   End
   Begin VB.Menu mnuRxx 
      Caption         =   "Rxx"
      Visible         =   0   'False
      Begin VB.Menu mnuWriteRx 
         Caption         =   "Write on Rx"
      End
      Begin VB.Menu mnuErase 
         Caption         =   "Erase Written"
         Visible         =   0   'False
      End
      Begin VB.Menu dsafadsf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSigned 
         Caption         =   "Signed Rx"
      End
      Begin VB.Menu mnuUnsigned 
         Caption         =   "Unsigned Rx"
      End
   End
End
Attribute VB_Name = "RxForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
'On Error GoTo errorOR
Merge
SavePicture Me.Picture, App.Path & "\Saved\" & Named & " " & "Multiple " & Replace(Dated, "/", "-") & " " & Format(Now, "ddmmyyhhmmss") & ".bmp"
Clipboard.SetText App.Path & "\" & Named & " " & "Multiple " & Replace(Dated, "/", "-") & ".pdf"
Dim X As Printer
List1.Visible = True
For Each X In Printers
      List1.AddItem X.DeviceName
Next
List1.SetFocus
'Printer.PaintPicture Me.Picture, 0, 0
'Printer.EndDoc
errorOR:
End Sub
Private Sub Command1z_Click()
'On Error GoTo errorOR
Dim i As Integer, Signedd As Integer, QbNumber As Integer

    Me.CurrentY = Named.Top
    Me.CurrentX = Named.Left
    Me.FontBold = Named.FontBold
    Me.FontItalic = Named.FontItalic
    Me.FontSize = Named.FontSize
    Me.ForeColor = Named.ForeColor
    Me.Font = Named.Font
    Me.Print Named
    Named.Visible = False
    Me.CurrentY = Aged.Top
    Me.CurrentX = Aged.Left
    Me.FontBold = Aged.FontBold
    Me.FontItalic = Aged.FontItalic
    Me.FontSize = Aged.FontSize
    Me.ForeColor = Aged.ForeColor
    Me.Font = Aged.Font
    Me.Print Aged
    Aged.Visible = False
    Me.CurrentY = Dated.Top
    Me.CurrentX = Dated.Left
    Me.FontBold = Dated.FontBold
    Me.FontItalic = Dated.FontItalic
    Me.FontSize = Dated.FontSize
    Me.ForeColor = Dated.ForeColor
    Me.Font = Dated.Font
    Me.Print Dated
    Dated.Visible = False
For i = 0 To 7
    Me.CurrentY = Drug1(i).Top
    Me.CurrentX = Drug1(i).Left
    Me.FontBold = Drug1(i).FontBold
    Me.FontItalic = Drug1(i).FontItalic
    Me.FontSize = Drug1(i).FontSize
    Me.ForeColor = Drug1(i).ForeColor
    Me.Font = Drug1(i).Font
    Me.Print Drug1(i)
    Drug1(i).Visible = False
    Me.CurrentY = Num1(i).Top
    Me.CurrentX = Num1(i).Left
    Me.FontBold = Num1(i).FontBold
    Me.FontItalic = Num1(i).FontItalic
    Me.FontSize = Num1(i).FontSize
    Me.ForeColor = Num1(i).ForeColor
    Me.Font = Num1(i).Font
    Me.Print Num1(i)
    Num1(i).Visible = False
    Me.CurrentY = Sig1(i).Top
    Me.CurrentX = Sig1(i).Left
    Me.FontBold = Sig1(i).FontBold
    Me.FontItalic = Sig1(i).FontItalic
    Me.FontSize = Sig1(i).FontSize
    Me.ForeColor = Sig1(i).ForeColor
    Me.Font = Sig1(i).Font
    Me.Print Sig1(i)
    Sig1(i).Visible = False
    Me.CurrentY = Ref1(i).Top
    Me.CurrentX = Ref1(i).Left
    Me.FontBold = Ref1(i).FontBold
    Me.FontItalic = Ref1(i).FontItalic
    Me.FontSize = Ref1(i).FontSize
    Me.ForeColor = Ref1(i).ForeColor
    Me.Font = Ref1(i).Font
    Me.Print Ref1(i)
    Ref1(i).Visible = False
    Me.CurrentY = Sub1(i).Top
    Me.CurrentX = Sub1(i).Left
    Me.FontBold = Sub1(i).FontBold
    Me.FontItalic = Sub1(i).FontItalic
    Me.FontSize = Sub1(i).FontSize
    Me.ForeColor = Sub1(i).ForeColor
    Me.Font = Sub1(i).Font
    Me.Print Sub1(i)
    Sub1(i).Visible = False
Next
Me.Refresh
Me.Picture = Me.Image
SavePicture Me.Picture, App.Path & "\" & Named & " " & Replace(Dated, "/", "-") & ".bmp"
Dim X As Printer
List1.Visible = True
For Each X In Printers
      List1.AddItem X.DeviceName
Next
List1.SetFocus
'Printer.PaintPicture Me.Picture, 0, 0
'Printer.EndDoc
errorOR:
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = 0
    'Me.MouseIcon = LoadPicture(App.Path & "\Pencil.ico")

End Sub

Private Sub Command2_Click()
List1.Visible = False
Me.Visible = False
Form1.mnuCollapse_Click
Form1.Visible = True
If MergedRx = False Then
    Form1.mnuSaveRx.Enabled = True
    Form1.mnuNewRx.Enabled = True
    Form1.mnuOpenRx.Enabled = True
    Form1.mnuPrint.Enabled = True
Else
    WriteBlankSingle = True
    WriteBlankMultiple = False
    Form1.Option1(0).Enabled = False
    Form1.TreeView1.Enabled = False
    Form1.Option1(1).Enabled = False
    Form1.mnuNewRx.Enabled = True
End If
Me.Visible = False
Form1.mnuPrint.Enabled = False
Form1.mnuRx.Enabled = False
Form1.mnuView.Enabled = False
Form1.mnuEdit.Enabled = False
Form1.mnuCollapse_Click
Form1.TreeView1.Enabled = False

Form1.Visible = True
PopupMenu Form1.mnuFile
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = 0

End Sub

Private Sub Command3_Click()
On Error Resume Next
Dim PtEdd As String
PtEdd = App.Path & "\Patient Education\" & PatientEducation & ".pdf"
StartDoc PtEdd
End Sub

Private Sub Drug1_Change(Index As Integer)
Exit Sub
    Drug1(Index).Text = Left(Drug1(Index).Text, 20)
    Drug1(Index).Refresh
End Sub

Private Sub Form_Activate()
If OpenFlag = True Then Merge

End Sub

Private Sub Form_Load()
Signedd = 0
Me.DrawWidth = 2
lblText(0).Left = -100000
    mnuSigned.Enabled = True
    mnuUnsigned.Enabled = True

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
List1.Visible = False
If Button = 1 And mnuWriteRx.Checked = True Then
    Me.CurrentX = X
    Me.CurrentY = Y
End If
Dim im As Integer
If InsertFlag = True Then
    Load lblText(k)
    InsertFlag = False
    lblExist = True
    Text1.Left = X
    Text1.Top = Y
    lblText(k).Left = X
    lblText(k).Top = Y
    Text1.Visible = True
    Text1.SetFocus
    k = k + 1
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And mnuWriteRx.Checked = True Then
        If mnuErase.Checked = True Then
            Line (Me.CurrentX, Me.CurrentY)-(X, Y), QBColor(15)
        End If
        If mnuErase.Checked = False Then
            Line (Me.CurrentX, Me.CurrentY)-(X, Y), QBColor(0)
        End If
    End If
    If mnuWriteRx.Checked = True And mnuErase.Checked = False Then
        Me.MousePointer = 99
        Me.MouseIcon = LoadPicture(App.Path & "\Pencil.ico")
    End If
    If mnuWriteRx.Checked = True And mnuErase.Checked = True Then
        Me.MousePointer = 99
        Me.MouseIcon = LoadPicture(App.Path & "\Eraser.ico")
    End If

    Me.Picture = Me.Image
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuRxx
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.mnuCollapse_Click
Form2.Visible = False
Form1.Visible = True
If ExitFlag = False Then PopupMenu Form1.mnuFile
End Sub

Private Sub Text1_Change()

End Sub

Private Sub List1_DblClick()
On Error GoTo errorOR
List1.Visible = False
Dim X As Printer
For Each X In Printers
    If X.DeviceName = List1.List(List1.ListIndex) Then Set Printer = X
Next
    Printer.PaintPicture Me.Picture, 0, 0
    Printer.EndDoc
    PrintedFlag = True
    Command2_Click
errorOR:
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = 0

End Sub

Private Sub mnuErase_Click()
If mnuErase.Checked = True Then
    mnuErase.Checked = False
    Me.MousePointer = 0
    Me.DrawWidth = 2: QbNumber = 0
Else
    mnuErase.Checked = True
    Me.MousePointer = 99
    Me.MouseIcon = LoadPicture(App.Path & "\Eraser.ico")
    QbNumber = 15
    Me.DrawWidth = 10
End If

End Sub

Private Sub mnuSigned_Click()
Me.Picture = LoadPicture(App.Path & "\MySignedSingle.jpg")
End Sub

Private Sub mnuUnsigned_Click()
Me.Picture = LoadPicture(App.Path & "\UnsignedSingle.jpg")

End Sub
Private Sub Merge()
On Error Resume Next
Dim i As Integer, k As Integer, l As Integer, Signedd As Integer, QbNumber As Integer
Dim SigTop As Integer, SigLeft As Integer, flag61 As Boolean
flag61 = False
MergedRx = True
    mnuSigned.Enabled = False
    mnuUnsigned.Enabled = False

    WriteBlank = True
    WriteBlankMultiple = False
    WriteBlankSingle = True
    Form1.Option1(0).Enabled = False
    Form1.Option1(1).Enabled = False
    Form1.TreeView1.Enabled = False
    Me.CurrentY = Label6.Top
    Me.CurrentX = Label6.Left
    Me.FontBold = Label6.FontBold
    Me.FontItalic = Label6.FontItalic
    Me.FontSize = Label6.FontSize
    Me.ForeColor = Label6.ForeColor
    Me.Font = Label6.Font
    Me.Print Label6
    Label6.Visible = False

    
    Me.CurrentY = Label5.Top
    Me.CurrentX = Label5.Left
    Me.FontBold = Label5.FontBold
    Me.FontItalic = Label5.FontItalic
    Me.FontSize = Label5.FontSize
    Me.ForeColor = Label5.ForeColor
    Me.Font = Label5.Font
    Me.Print Label5
    Label5.Visible = False
    
    Me.CurrentY = Named.Top
    Me.CurrentX = Named.Left
    Me.FontBold = Named.FontBold
    Me.FontItalic = Named.FontItalic
    Me.FontSize = Named.FontSize
    Me.ForeColor = Named.ForeColor
    Me.Font = Named.Font
    Me.Print Named
    Named.Visible = False
    Me.CurrentY = Aged.Top
    Me.CurrentX = Aged.Left
    Me.FontBold = Aged.FontBold
    Me.FontItalic = Aged.FontItalic
    Me.FontSize = Aged.FontSize
    Me.ForeColor = Aged.ForeColor
    Me.Font = Aged.Font
    Me.Print Aged
    Aged.Visible = False
    Me.CurrentY = Dated.Top
    Me.CurrentX = Dated.Left
    Me.FontBold = Dated.FontBold
    Me.FontItalic = Dated.FontItalic
    Me.FontSize = Dated.FontSize
    Me.ForeColor = Dated.ForeColor
    Me.Font = Dated.Font
    Me.Print Dated
    Dated.Visible = False
    
    Me.CurrentY = Label1.Top
    Me.CurrentX = Label1.Left
    Me.FontBold = Label1.FontBold
    Me.FontItalic = Label1.FontItalic
    Me.FontSize = Label1.FontSize
    Me.ForeColor = Label1.ForeColor
    Me.Font = Label1.Font
    Me.Print Label1
    Label1.Visible = False
    
    Me.CurrentY = Label3.Top
    Me.CurrentX = Label3.Left
    Me.FontBold = Label3.FontBold
    Me.FontItalic = Label3.FontItalic
    Me.FontSize = Label3.FontSize
    Me.ForeColor = Label3.ForeColor
    Me.Font = Label3.Font
    Me.Print Label3
    Label3.Visible = False
    
    Me.CurrentY = Label4.Top
    Me.CurrentX = Label4.Left
    Me.FontBold = Label4.FontBold
    Me.FontItalic = Label4.FontItalic
    Me.FontSize = Label4.FontSize
    Me.ForeColor = Label4.ForeColor
    Me.Font = Label4.Font
    Me.Print Label4
    Label4.Visible = False
    
    
    Me.CurrentY = Label2.Top
    Me.CurrentX = Label2.Left
    Me.FontBold = Label2.FontBold
    Me.FontItalic = Label2.FontItalic
    Me.FontSize = Label2.FontSize
    Me.ForeColor = Label2.ForeColor
    Me.Font = Label2.Font
    Me.Print Label2
    Label2.Visible = False
    
    Me.CurrentY = Drug1(0).Top
    Me.CurrentX = Drug1(0).Left
    Me.FontBold = Drug1(0).FontBold
    Me.FontItalic = Drug1(0).FontItalic
    Me.FontSize = Drug1(0).FontSize
    Me.ForeColor = Drug1(0).ForeColor
    Me.Font = Drug1(0).Font
    Me.Print Drug1(0)
    Drug1(0).Visible = False
    Me.CurrentY = Num1(0).Top
    Me.CurrentX = Num1(0).Left
    Me.FontBold = Num1(0).FontBold
    Me.FontItalic = Num1(0).FontItalic
    Me.FontSize = Num1(0).FontSize
    Me.ForeColor = Num1(0).ForeColor
    Me.Font = Num1(0).Font
    Me.Print Num1(0)
    Num1(0).Visible = False
    
    Me.CurrentY = Sig1(0).Top
    Me.CurrentX = Sig1(0).Left
    Me.FontBold = Sig1(0).FontBold
    Me.FontItalic = Sig1(0).FontItalic
    Me.FontSize = Sig1(0).FontSize
    Me.ForeColor = Sig1(0).ForeColor
    Me.Font = Sig1(0).Font
    Me.Print Sig1(0)
    Sig1(0).Visible = False
    
    
    k = 1
    Me.CurrentY = Comment.Top
    Me.CurrentX = Comment.Left
    Me.FontBold = Comment.FontBold
    Me.FontItalic = Comment.FontItalic
    Me.FontSize = Comment.FontSize
    Me.ForeColor = Comment.ForeColor
    Me.Font = Comment.Font
    'Me.Print Comment
    Load lblText(k): lblText(k).Visible = True
    lblText(k).Left = SigLeft
    lblText(k).Top = Comment.Top
    For i = 1 To Len(Comment.Text)
         If l = 61 Then
             Me.Print lblText(k).Caption
             lblText(k).Visible = False
             k = k + 1
             l = 0
             Load lblText(k): lblText(k).Visible = True
             'lblText(k).Left = SigLeft
             'lblText(k).Top = Comment.Top - 10
             Me.CurrentX = SigLeft
             Me.CurrentY = Me.CurrentY - 10
         End If
         lblText(k).Caption = lblText(k).Caption & Mid(Comment, i, 1)
         l = l + 1
     Next
     If flag61 = False Then
        Me.Print lblText(k).Caption
        lblText(k).Visible = False
     End If
    lblText(0).Visible = False
    Comment.Visible = False
    
    Me.CurrentY = Ref1(0).Top
    Me.CurrentX = Ref1(0).Left
    Me.FontBold = Ref1(0).FontBold
    Me.FontItalic = Ref1(0).FontItalic
    Me.FontSize = Ref1(0).FontSize
    Me.ForeColor = Ref1(0).ForeColor
    Me.Font = Ref1(0).Font
    Me.Print Ref1(0)
    Ref1(0).Visible = False
        
    Me.CurrentY = Sub1(0).Top
    Me.CurrentX = Sub1(0).Left
    Me.FontBold = Sub1(0).FontBold
    Me.FontItalic = Sub1(0).FontItalic
    Me.FontSize = Sub1(0).FontSize
    Me.ForeColor = Sub1(0).ForeColor
    Me.Font = Sub1(0).Font
    Me.Print Sub1(0)
    Sub1(0).Visible = False
Me.Refresh
Me.Picture = Me.Image

Exit Sub


End Sub
Public Sub mnuWriteRx_Click()
Dim intsave As Integer
If WriteBlank = False Then
    intsave = MsgBox("You will no longer be able to add Rx to this Prescription! Do you want to Continue?", _
                     vbYesNoCancel + vbExclamation)
    Select Case intsave
        Case vbYes
            Merge
            WriteOnFlag = True
            mnuWriteRx.Checked = True
            Form1.Command3.Enabled = False
            mnuWriteRx.Enabled = False
            Me.MousePointer = 99
            Me.MouseIcon = LoadPicture(App.Path & "\Pencil.ico")
            mnuErase.Visible = True
    End Select
Else
            Merge
            WriteOnFlag = True
            mnuWriteRx.Checked = True
            Form1.Command3.Enabled = False
            mnuWriteRx.Enabled = False
            Me.MousePointer = 99
            Me.MouseIcon = LoadPicture(App.Path & "\Pencil.ico")
            mnuErase.Visible = True
End If
End Sub
