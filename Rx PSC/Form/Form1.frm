VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00DCE9EB&
   Caption         =   "Writer"
   ClientHeight    =   7860
   ClientLeft      =   1785
   ClientTop       =   1755
   ClientWidth     =   9015
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form10"
   Picture         =   "Form1.frx":08CA
   ScaleHeight     =   7860
   ScaleWidth      =   9015
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00DCE9EB&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   1440
      ScaleHeight     =   6465
      ScaleWidth      =   6465
      TabIndex        =   14
      Top             =   360
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CommandButton Command6 
         Caption         =   "Add"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5955
         TabIndex        =   35
         Top             =   1290
         Width           =   450
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00DFEDEC&
         Caption         =   "==="
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   5.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   5955
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1110
         Width           =   450
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EDFBF9&
         Caption         =   "Multiple Rx"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   5520
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EDFBF9&
         Caption         =   "Single Rx"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   5520
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00DFEDEC&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1215
         TabIndex        =   4
         Top             =   2160
         Width           =   4935
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00EDFBF9&
         Caption         =   "Add Rx"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   5880
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BackColor       =   &H00DFEDEC&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BackColor       =   &H00DFEDEC&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Text            =   "58"
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00DFEDEC&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Text            =   "Warren S. Goff DO"
         Top             =   1080
         Width           =   4695
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00DFEDEC&
         Caption         =   "Do Not Substitute Generic"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1920
         TabIndex        =   7
         Top             =   3720
         Width           =   3015
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00EDFBF9&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   5880
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00EDFBF9&
         Caption         =   "Print"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2790
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   5880
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00DFEDEC&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   4320
         Width           =   5895
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00DFEDEC&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2985
         TabIndex        =   6
         Text            =   "4"
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00DFEDEC&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   5
         Text            =   "30"
         Top             =   2640
         Width           =   855
      End
      Begin VB.ListBox List3 
         BackColor       =   &H00DCE9EB&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4890
         Left            =   15
         Sorted          =   -1  'True
         TabIndex        =   32
         Top             =   1680
         Visible         =   0   'False
         Width           =   6495
      End
      Begin VB.Shape Shape1 
         Height          =   255
         Left            =   5925
         Top             =   1140
         Width           =   495
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5910
         TabIndex        =   2
         Top             =   900
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Edit Sig:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   240
         TabIndex        =   31
         Top             =   2160
         Width           =   825
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Print Rx"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   2160
         TabIndex        =   30
         Top             =   0
         Width           =   2055
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date: "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2280
         TabIndex        =   22
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age: "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   240
         TabIndex        =   20
         Top             =   1680
         Width           =   510
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name: "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   240
         TabIndex        =   19
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comment:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   240
         TabIndex        =   17
         Top             =   4080
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Refills: "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1920
         TabIndex        =   16
         Top             =   3000
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Number: "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1920
         TabIndex        =   15
         Top             =   2640
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rx Name"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   300
         TabIndex        =   18
         Top             =   585
         Width           =   1065
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rx Name"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   285
         TabIndex        =   33
         Top             =   585
         Width           =   1065
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   7200
         Left            =   -120
         Picture         =   "Form1.frx":ADB3
         Stretch         =   -1  'True
         Top             =   0
         Width           =   6600
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00DCE9EB&
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   1785
      Picture         =   "Form1.frx":1529C
      ScaleHeight     =   3705
      ScaleWidth      =   6345
      TabIndex        =   23
      Top             =   2400
      Visible         =   0   'False
      Width           =   6375
      Begin VB.ListBox List2 
         Height          =   255
         Left            =   720
         TabIndex        =   29
         Top             =   3360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00EDFBF9&
         Caption         =   "Done"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   360
         TabIndex        =   25
         Text            =   "Search Term"
         Top             =   600
         Width           =   5655
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00EDFBF9&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   3240
         Width           =   1215
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2085
         Left            =   360
         TabIndex        =   26
         Top             =   1080
         Width           =   5655
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Search Rx"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   2160
         TabIndex        =   28
         Top             =   120
         Width           =   2055
      End
   End
   Begin MSComDlg.CommonDialog cdlOpen 
      Left            =   960
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   3413
      _Version        =   393217
      Sorted          =   -1  'True
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00DCE9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   2970
      Picture         =   "Form1.frx":1F785
      ScaleHeight     =   870
      ScaleWidth      =   4320
      TabIndex        =   36
      Top             =   3360
      Width           =   4320
   End
   Begin VB.Image Img 
      Height          =   2880
      Left            =   5280
      Picture         =   "Form1.frx":20A5B
      Top             =   0
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNewRx 
         Caption         =   "New Rx"
      End
      Begin VB.Menu mnuNewPatient 
         Caption         =   "New Patient"
      End
      Begin VB.Menu mnuSaveRx 
         Caption         =   "Save Rx"
      End
      Begin VB.Menu mnuOpenRx 
         Caption         =   "Open Rx"
      End
      Begin VB.Menu asfasdf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSendEmail 
         Caption         =   "Send Email"
      End
      Begin VB.Menu sdfsgs 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print Rx"
      End
      Begin VB.Menu mnuPtEd1 
         Caption         =   "Patient Education"
      End
      Begin VB.Menu afafsdf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuAdd2 
         Caption         =   "&Add Item"
      End
      Begin VB.Menu mnuAddDose1 
         Caption         =   "Add Dosage"
      End
      Begin VB.Menu mnuAddComment1 
         Caption         =   "Add Comment"
      End
      Begin VB.Menu fklfsl 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExternal 
         Caption         =   "External Editor Drugs"
      End
      Begin VB.Menu mnuUndoRx 
         Caption         =   "Undo Edit Drugs"
      End
      Begin VB.Menu mnuExtPatients 
         Caption         =   "External Editor Patients"
      End
      Begin VB.Menu mnuUndoPt 
         Caption         =   "Undo Edit Patient List"
      End
      Begin VB.Menu jdlslf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearchRx 
         Caption         =   "Search Rx"
      End
   End
   Begin VB.Menu mnuRx 
      Caption         =   "Rx"
      Begin VB.Menu mnuAddFav 
         Caption         =   "Add to Favorites"
      End
      Begin VB.Menu kfdkfdsdsf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFavList 
         Caption         =   "Open Favorites List"
      End
      Begin VB.Menu mnuSaveFave 
         Caption         =   "Save Favorties"
      End
      Begin VB.Menu mnuAllDrugs 
         Caption         =   "Open All Drugs List"
      End
      Begin VB.Menu lkkfkd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBlankSingle 
         Caption         =   "Blank Handwritten Single Rx"
      End
      Begin VB.Menu mnuBlankMultiRx 
         Caption         =   "Bland Handwritten Multi Rx"
      End
      Begin VB.Menu fsadasgg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSignedRx 
         Caption         =   "Signed Rx"
      End
      Begin VB.Menu mnuUnsignedRx 
         Caption         =   "Unsigned  Rx"
      End
      Begin VB.Menu sfsggasg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSRx1 
         Caption         =   "Search Rx"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuMedscape 
         Caption         =   "Search Medscape"
      End
      Begin VB.Menu kfjdfk 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTopp 
         Caption         =   "Top"
      End
      Begin VB.Menu mnuCollapse 
         Caption         =   "Collapse"
      End
      Begin VB.Menu mnuExpand 
         Caption         =   "Expand"
      End
      Begin VB.Menu mnuBottomm 
         Caption         =   "Bottom"
      End
      Begin VB.Menu dfsadfsafd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSrch 
         Caption         =   "Search Rx"
      End
      Begin VB.Menu dwdd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTop 
         Caption         =   "Topmost"
      End
      Begin VB.Menu dafsadf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMU 
         Caption         =   "Move Up"
      End
      Begin VB.Menu mnuMD 
         Caption         =   "Move Down"
      End
      Begin VB.Menu dfgdfdfg 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh List"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupAddNode 
         Caption         =   "&Add Item"
      End
      Begin VB.Menu mnuAddDose 
         Caption         =   "Add Dosage"
      End
      Begin VB.Menu mnuComment 
         Caption         =   "Add Comment"
      End
      Begin VB.Menu mnuPopupDeleteNode 
         Caption         =   "&Delete Item"
      End
      Begin VB.Menu sfsadf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuATF 
         Caption         =   "Add to Favorites"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAddPE 
         Caption         =   "Add Patient Education"
         Visible         =   0   'False
      End
      Begin VB.Menu adsfasdfdsa 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMD1 
         Caption         =   "Search Medscape"
      End
      Begin VB.Menu kkfl 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPRx 
         Caption         =   "Print Rx"
      End
      Begin VB.Menu safsaf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "Search"
         Shortcut        =   ^S
      End
      Begin VB.Menu dwedwed 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMoveUp 
         Caption         =   "Move Up"
      End
      Begin VB.Menu mnuMoveDown 
         Caption         =   "Move Down"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuInstruct 
         Caption         =   "Help File"
      End
      Begin VB.Menu dwfwef 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMedscape1 
         Caption         =   "Medscape Drugs"
      End
      Begin VB.Menu mnuEpocrates 
         Caption         =   "Epocrates"
      End
      Begin VB.Menu lfgdfg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInteract 
         Caption         =   "Drug Interactions"
      End
      Begin VB.Menu mnuPtEdFold 
         Caption         =   "Pt Education Folder"
      End
      Begin VB.Menu kkfkfsf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWalmartRx 
         Caption         =   "Walmart $4.00 Rx "
      End
      Begin VB.Menu mnuWalmartOTC 
         Caption         =   "Walmart $4.00 OTC"
      End
      Begin VB.Menu oorfw 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNEJM 
         Caption         =   "NEJM"
      End
      Begin VB.Menu mnuJAMA 
         Caption         =   "JAMA"
      End
      Begin VB.Menu mnuArchives 
         Caption         =   "Archives IM"
      End
      Begin VB.Menu mnuAAFP 
         Caption         =   "AAFP"
      End
      Begin VB.Menu mnuCochrane 
         Caption         =   "Cochrane"
      End
      Begin VB.Menu mnuAnnalsFP 
         Caption         =   "Annals Family Practice"
      End
      Begin VB.Menu mnuMayo 
         Caption         =   "Mayo Clinics"
      End
      Begin VB.Menu mnuSage 
         Caption         =   "Sage Journals"
      End
      Begin VB.Menu mnuPubmed 
         Caption         =   "PubMed"
      End
      Begin VB.Menu dsded 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWebsite 
         Caption         =   "Cyberjournal of Medicine Website"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type PAINTSTRUCT
    hDC As Long
    fErase As Long
    rcPaint As RECT
    fRestore As Long
    fIncUpdate As Long
    rgbReserved As Byte
End Type

Enum BrowseForFolderFlags
    BIF_RETURNONLYFSDIRS = &H1
    BIF_DONTGOBELOWDOMAIN = &H2
    BIF_STATUSTEXT = &H4
    BIF_BROWSEFORCOMPUTER = &H1000
    BIF_BROWSEFORPRINTER = &H2000
    BIF_BROWSEINCLUDEFILES = &H4000
    BIF_EDITBOX = &H10
    BIF_RETURNFSANCESTORS = &H8
End Enum

'BrowseInfo is a type used with the SHBrowseForFolder API call
Private Type BrowseInfo
     hwndOwner As Long
     pIDLRoot As Long
     pszDisplayName As Long
     lpszTitle As Long
     ulFlags As Long
     lpfnCallback As Long
     lParam As Long
     iImage As Long
End Type 'Shell APIs from Shell32.dll file:
'SHBrowseForFolder - Gets the Browse For Folder Dialog
Private Declare Function SHBrowseForFolder Lib "shell32" (lpBI As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public shlShell As shell32.Shell
Public shlFolder As shell32.Folder

Private Declare Function BeginPaint Lib "user32" _
    (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32" _
    (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" _
    (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" _
    (ByVal hDC As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" _
    (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" (ByVal hwnd As Long, _
    ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long
Private Declare Function BitBlt Lib "gdi32" _
    (ByVal hDestDC As Long, ByVal X As Long, _
    ByVal Y As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long, ByVal hSrcDC As Long, _
    ByVal xSrc As Long, ByVal ySrc As Long, _
    ByVal dwRop As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" _
    (ByVal hObject As Long) As Long
Private Declare Function InvalidateRect Lib "user32" _
    (ByVal hwnd As Long, ByVal lpRect As Long, _
    ByVal bErase As Long) As Long

Private Const WM_PAINT = &HF
Private Const WM_ERASEBKGND = &H14
Private Const WM_HSCROLL = &H114
Private Const WM_VSCROLL = &H115
Private Const WM_MOUSEWHEEL = &H20A
Dim RxNumber As Integer, Another As Boolean, Saved As Boolean, SigChanged As Boolean
Dim oRx As Boolean, oRx1 As Boolean
Dim file_name As String, Favorites As Boolean
' Load a TreeView control from a file that uses tabs
' to show indentation.
Private Sub LoadTreeViewFromFile(ByVal file_named As String, ByVal trv As TreeView)
Dim fnum As Integer
Dim text_line As String
Dim level As Integer
Dim tree_nodes() As Node
Dim num_nodes As Integer
    DoEvents: DoEvents

    fnum = FreeFile
    Open file_named For Input As fnum

    TreeView1.Nodes.Clear
    Do While Not EOF(fnum)
    DoEvents: DoEvents
        ' Get a line.
        Line Input #fnum, text_line

        ' Find the level of indentation.
        level = 1
        Do While Left$(text_line, 1) = vbTab
            level = level + 1
            text_line = Mid$(text_line, 2)
        Loop

        ' Make room for the new node.
        If level > num_nodes Then
            num_nodes = level
            ReDim Preserve tree_nodes(1 To num_nodes)
        End If

        ' Add the new node.
        If level = 1 Then
            Set tree_nodes(level) = TreeView1.Nodes.Add(, , , text_line)
        Else
            Set tree_nodes(level) = TreeView1.Nodes.Add(tree_nodes(level - 1), tvwChild, , text_line)
            tree_nodes(level).EnsureVisible
        End If
    DoEvents: DoEvents
    Loop
    DoEvents: DoEvents

    Close fnum
    If trv.Nodes.Count > 0 Then trv.Nodes(1).EnsureVisible
End Sub
' Write tabs indicating this node's depth in
' the tree followed by the node's text.
' Then save its children and its siblings.
Private Sub SaveNode(ByVal fnum As Integer, ByVal n As Node, ByVal level As Integer)
    If n Is Nothing Then Exit Sub

    ' Save the node.
    Print #fnum, String$(level, vbTab) & n.Text

    ' Save its children.
    SaveNode fnum, n.Child, level + 1

    ' Save its next sibling.
    SaveNode fnum, n.Next, level
End Sub
' Save a TreeView control into a file that uses tabs
' to show indentation.
Private Sub SaveTreeViewIntoFile(ByVal file_named As String, ByVal trv As TreeView)
Dim fnum As Integer
'MsgBox file_name
    fnum = FreeFile
    Open file_named For Output As fnum

    ' Find the root nodes.
    If TreeView1.Nodes.Count > 0 Then SaveNode fnum, TreeView1.Nodes(1), 0

    Close fnum
End Sub

Private Sub Check2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim XXX As String
If Check2.Value <> False Then
    List3.Clear
    'List3.AddItem "Cancel"
    Open App.Path & "\PatientList.txt" For Input As #1
        Do While Not EOF(1)
            Line Input #1, XXX
            List3.AddItem XXX
        Loop
    Close #1
    List3.Visible = True
    List3.SetFocus
Else
    List3.Visible = False
    Exit Sub
End If

End Sub



Private Sub Check3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub Command1_Click()
'If Saved <> True Then
'    Dim intsave As Integer
'    intsave = MsgBox("Do you want to Save the Rx before Printing?", _
'                     vbYesNoCancel + vbExclamation)
'    Select Case intsave
'        Case vbYes
'            mnuSaveRx_Click
'            Exit Sub
'        Case vbCancel
'            Exit Sub
'    End Select
'End If
'mnuSaveRx.Enabled = False
If Option1(0).Value = True Or oRx = False Then
    If Dir(App.Path & "\Patient Education\" & PatientEducation & ".pdf") <> "" Then
        RxForm2.Command3.Visible = True
    Else
        RxForm2.Command3.Visible = False
    End If
    Load RxForm2
    RxForm2.Show
Else
    If Dir(App.Path & "\Patient Education\" & PatientEducation & ".pdf") <> "" Then
        Form2.Command3.Visible = True
    Else
        Form2.Command3.Visible = False
    End If
    Load Form2
    Form2.Show
End If


Picture1.Visible = False
TreeView1.Enabled = True
Me.Visible = False
End Sub

Private Sub Command2_Click()
Newed
Form1.mnuView.Enabled = False
Form1.mnuEdit.Enabled = False
Form1.mnuRx.Enabled = False
Form1.mnuPrint.Enabled = False
Form1.mnuCollapse_Click
Form1.TreeView1.Enabled = False
Picture1.Visible = False
Form1.Visible = True
PopupMenu Form1.mnuFile
End Sub

Public Sub Command3_Click()
On Error Resume Next
'Form2.Caption = Trim(Text4.Text) & " " & Trim(Text6.Text)
mnuSaveRx.Enabled = True
Saved = False

If RxNumber = 8 Then
    MsgBox "You CANNOT Add Anymore Items to this Rx!"
    Command3.Enabled = False
    Another = False
    Exit Sub
End If

Option1(1).Enabled = False
Option1(0).Enabled = False
If RxNumber = 1 And Option1(0).Value = True Then
    MsgBox "You have Selected a Single Rx Format and only one Rx can be added at a time!"
    Exit Sub
End If
If Trim(Trim(Text1.Text)) = "" Then
    MsgBox "Please Enter the Quantity!"
    Text1.SetFocus
    Exit Sub
End If
If Trim(Trim(Text4.Text)) = "" Then
    MsgBox "Please Enter the Patient's Name!"
    Text4.SetFocus
    Exit Sub
End If
Dim XX As String
Command1.Enabled = True
XX = Trim(Text4.Text) & " " & Replace(Trim(Text6.Text), "/", "-")
'Clipboard.SetText XX
Select Case Option1(0).Value
    Case False
        Form2.Drug1(RxNumber).Text = Left(Trim(TreeView1.SelectedItem.Parent), 24)
        Form2.Num1(RxNumber).Caption = Trim(Text1.Text)
        If SigChanged = False Then
            Form2.Sig1(RxNumber).Caption = Replace(Trim(TreeView1.SelectedItem.Text), "(Rx)", "")
        Else
            SigChanged = False
            Form2.Sig1(RxNumber).Caption = Text8.Text
        End If
        Form2.Ref1(RxNumber).Caption = Trim(Text2.Text)
        If Trim(Text2.Text) = "" Then Form2.Ref1(RxNumber).Caption = "0"
        Form2.Named.Caption = Trim(Text4.Text)
        Form2.Aged.Caption = Trim(Text5.Text)
        Form2.Dated.Caption = Trim(Text6.Text)
        Form2.Comment.Text = Text3.Text
        If Check1.Value = 0 Then
            Form2.Sub1(RxNumber).Caption = "Yes"
        Else
            Form2.Sub1(RxNumber).Caption = "No"
        End If
    Case True  'single rx1
        RxForm2.Drug1(RxNumber).Text = Left(Trim(TreeView1.SelectedItem.Parent), 30)
        RxForm2.Num1(RxNumber).Caption = Trim(Text1.Text)
        If SigChanged = False Then
            RxForm2.Sig1(RxNumber).Text = Replace(Trim(TreeView1.SelectedItem.Text), "(Rx)", "")
        Else
            SigChanged = False
            RxForm2.Sig1(RxNumber).Text = Form2.Sig1(RxNumber).Caption = Text8.Text
        End If
        RxForm2.Ref1(RxNumber).Caption = Trim(Text2.Text)
        If Trim(Text2.Text) = "" Then Form2.Ref1(RxNumber).Caption = "0"
        RxForm2.Named.Caption = Trim(Text4.Text)
        RxForm2.Aged.Caption = Trim(Text5.Text)
        RxForm2.Dated.Caption = Trim(Text6.Text)
        RxForm2.Comment.Text = Text3.Text
        If Check1.Value = 0 Then
            RxForm2.Sub1(RxNumber).Caption = "Yes"
        Else
            RxForm2.Sub1(RxNumber).Caption = "No"
        End If
End Select
SaveSub
mnuNewRx.Enabled = True
RxNumber = RxNumber + 1

If Option1(0).Value = True Then
    Command3.Enabled = False
    Command1_Click
    Exit Sub  ': Command1_Click
Else
End If
Dim intsave As Integer
intsave = MsgBox("Do you want to Add Another Rx?", _
                 vbYesNoCancel + vbExclamation)
Select Case intsave
    Case vbYes
        Picture1.Visible = False
        TreeView1.Enabled = True
    Case vbNo
        SaveSub
        Saved = False
        mnuSaveRx.Enabled = True
        Command3.Enabled = False
        Command1_Click
End Select


End Sub

Private Sub Command4_Click()
On Error Resume Next
List1.Clear
Dim w As Integer, z As Integer, X As Integer
Dim flg As Byte: flg = 0
UnSubclass TreeView1
TreeView1.Visible = False
mnuCollapse_Click
For w = 1 To TreeView1.Nodes.Count ' step thru all of the nodes
    If TreeView1.Nodes(w).Parent Is Nothing Then 'this is the top node
        For z = 1 To TreeView1.Nodes(w).Children
            If z = 1 Then
                X = TreeView1.Nodes(w).Child.Index
            Else
                X = TreeView1.Nodes(X).Next.Index
            End If
                If InStrRev(TreeView1.Nodes(X).Text, Text7.Text, , 1) > 0 Then 'search the children node's Titles
                    TreeView1.Nodes(w).Expanded = True
                    TreeView1.Nodes(w).ForeColor = &HFF&
                    TreeView1.Nodes(X).Expanded = True
                    'Text8 = Text8 + TreeView1.Nodes(w).Text + "  -  " + TreeView1.Nodes(X).Text + Format$(X, " 00") + vbCrLf
                    List1.AddItem TreeView1.Nodes(X)
                    List2.AddItem TreeView1.Nodes(w).Text
                    flg = 1
                End If
        Next z
    End If
Next w
Subclass Me, TreeView1
TreeView1.Visible = True
If flg = 0 Then Text7 = "Search String Not Found"
End Sub

Private Sub Command5_Click()
Picture2.Visible = False
TreeView1.Enabled = True
mnuTopp_Click
End Sub

Private Sub Command6_Click()
Dim XXX As String

    List3.Clear
    Open App.Path & "\PatientList.txt" For Append As #1
        Print #1, Text4.Text
    Close #1

End Sub

Private Sub Form_Activate()
On Error Resume Next
Picture1.Visible = False
    TreeView1.SetFocus
    SendKeys "^{HOME}"
    SendKeys "{HOME}"
    If mnuTop.Checked = True Then SetTopMostWindow Me.hwnd, True
End Sub

Private Sub Form_Deactivate()
    If mnuTop.Checked = True Then SetTopMostWindow Me.hwnd, True

End Sub

Private Sub Form_Initialize()
    'Newed
    Text6.Text = Date
    Saved = True
    Unload Loading
    Set Loading = Nothing
End Sub

Private Sub Form_Load()
'Subclass Me, TreeView1
On Error Resume Next
If Dir("C:\1down\OfficeVisitorPSC\PatientEducation\", vbDirectory) <> "" Then
    mnuPtEdFold.Visible = True
Else
    mnuPtEdFold.Visible = False
End If
RxNumber = 0
Form2.Picture = LoadPicture(App.Path & "\UnsignedSingle.jpg")
RxForm2.Picture = LoadPicture(App.Path & "\UnsignedSingle.jpg")
DoEvents: DoEvents
List3.ZOrder 0
    
    'mnuMD.Enabled = False
    'mnuMU.Enabled = False
    DoEvents: DoEvents
    oRx = True
    FM = Format(Now, "ddmmyyhhmmss")
    If RxNumber = 0 Then mnuPrint.Enabled = False
    mnuPRx.Enabled = False
    mnuSaveRx.Enabled = False
    mnuPtEd1.Enabled = False
    mnuAddFav.Enabled = False
    mnuFavList.Enabled = True
    mnuAllDrugs = False
    mnuSaveFave.Visible = False
    mnuUndoRx.Visible = False
    mnuUndoPt.Visible = False
    'mnuPrintED.Enabled = False
    
    mnuNewRx.Enabled = False
    file_name = App.Path
    If Right$(file_name, 1) <> "\" Then file_name = file_name & "\"
    DoEvents: DoEvents
    file_name = file_name & "test.txt"
    LoadTreeViewFromFile file_name, TreeView1
    DoEvents: DoEvents
Dim n As Node
    DoEvents: DoEvents
    For Each n In TreeView1.Nodes
        If n.Expanded Then n.Expanded = False
    Next
    Text6.Text = Date
    DoEvents: DoEvents
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.width - Me.width) / 2
Picture1.Left = (Me.width - Picture1.width) / 2
Picture1.Top = (Me.Height - Picture1.Height) / 2
Picture2.Left = (Me.width - Picture2.width) / 2
Picture2.Top = (Me.Height - Picture2.Height) / 2
Picture3.Left = (Me.width - Picture3.width) / 2
Picture3.Top = (Me.Height - Picture3.Height) / 2
DoEvents: DoEvents
End Sub

Private Sub Form_LostFocus()
    If mnuTop.Checked = True Then SetTopMostWindow Me.hwnd, True

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Saved = False Then
    Dim intsave As Integer
    intsave = MsgBox("Do you want to Save the Rx before Exiting?", _
                     vbYesNoCancel + vbExclamation)
    Select Case intsave
        Case vbYes
            mnuSaveRx_Click
            Cancel = 1
            Exit Sub
        Case vbCancel
            Cancel = 1
    End Select
End If
UnSubclass TreeView1
ExitFlag = True
'Dim file_name As String

'If Favorites = True Then
'    FileCopy App.Path & "\test.txt", App.Path & "\Favorites.txt"
'    FileCopy App.Path & "\test.bak", App.Path & "\test.txt"
'Else
    'file_name = App.Path
    'If Right$(file_name, 1) <> "\" Then file_name = file_name & "\"
    'file_name = file_name & "test.txt"
    SaveTreeViewIntoFile file_name, TreeView1
    Unload Me
    Unload Form2
    Set Form2 = Nothing
    Unload RxForm2
    Set RxForm2 = Nothing
    CloseAll
'End If
End Sub

Private Sub Form_Resize()
    TreeView1.Move 0, 0, ScaleWidth, ScaleHeight
    Picture1.Left = (Me.width - Picture1.width) / 2
    Picture1.Top = (Me.Height - Picture1.Height) / 2
    Picture2.Left = (Me.width - Picture2.width) / 2
    Picture2.Top = (Me.Height - Picture2.Height) / 2
    Picture3.Left = (Me.width - Picture3.width) / 2
    Picture3.Top = (Me.Height - Picture3.Height) / 2
    Subclass Me, TreeView1
End Sub
Public Sub TreeViewMessage(ByVal hwnd As Long, _
    ByVal wMsg As Long, ByVal wParam As Long, _
    ByVal lParam As Long, RetVal As Long, _
    UseRetVal As Boolean)

'Prevent recursion with this variable
Static InProc As Boolean

Dim ps As PAINTSTRUCT
Dim TVDC As Long, drawDC1 As Long, drawDC2 As Long
Dim oldBMP1 As Long, drawBMP1 As Long
Dim oldBMP2 As Long, drawBMP2 As Long
Dim X As Long, Y As Long, w As Long, h As Long
Dim TVWidth As Long, TVHeight As Long

If wMsg = WM_PAINT Then
    If InProc = True Then
        Exit Sub
    End If
    InProc = True
    'Prepare some variables we'll use
    TVWidth = TreeView1.width \ Screen.TwipsPerPixelX
    TVHeight = TreeView1.Height \ Screen.TwipsPerPixelY

    w = ScaleX(Img.Picture.width, vbHimetric, vbPixels)
    h = ScaleY(Img.Picture.Height, vbHimetric, vbPixels)

    'Begin painting. This API must be called in
    'response to the WM_PAINT message or you'll see
    'some odd visual effects :-)
    Call BeginPaint(hwnd, ps)
    TVDC = ps.hDC

    'Create a few canvases in memory to
    'draw on
    drawDC1 = CreateCompatibleDC(TVDC)
    drawBMP1 = CreateCompatibleBitmap(TVDC, TVWidth, TVHeight)
    oldBMP1 = SelectObject(drawDC1, drawBMP1)

    drawDC2 = CreateCompatibleDC(TVDC)
    drawBMP2 = CreateCompatibleBitmap(TVDC, TVWidth, TVHeight)
    oldBMP2 = SelectObject(drawDC2, drawBMP2)

    'This actually causes the TreeView to paint
    'itself onto our memory DC!
    SendMessage hwnd, WM_PAINT, drawDC1, ByVal 0&
    'Tile the bitmap and draw the TreeView
    'over it transparently
    For Y = 0 To TVHeight Step h
        For X = 0 To TVWidth Step w
            PaintNormalStdPic drawDC2, X, Y, w, h, _
                Img.Picture, 0, 0
        Next
    Next
    PaintTransparentDC drawDC2, 0, 0, TVWidth, TVHeight, _
        drawDC1, 0, 0, TranslateColor(vbWindowBackground)
    'Draw to the target DC
    BitBlt TVDC, 0, 0, TVWidth, TVHeight, _
        drawDC2, 0, 0, vbSrcCopy

    'Cleanup
    SelectObject drawDC1, oldBMP1
    SelectObject drawDC2, oldBMP2
    DeleteObject drawBMP1
    DeleteObject drawBMP2

    EndPaint hwnd, ps

    RetVal = 0
    UseRetVal = True
    InProc = False

ElseIf wMsg = WM_ERASEBKGND Then
    'Return TRUE
    RetVal = 1
    UseRetVal = True

ElseIf wMsg = WM_HSCROLL Or wMsg = WM_VSCROLL Or wMsg = WM_MOUSEWHEEL Then
    'Force a repaint to keep the bitmap
    'tiles lined up
    InvalidateRect hwnd, 0, 0

End If

End Sub

Sub CloseAll()
    On Error Resume Next
    Dim intFrmNum As Integer
    intFrmNum = Forms.Count


    Do Until intFrmNum = 0
        Unload Forms(intFrmNum - 1)
        Set Forms(intFrmNum - 1) = Nothing
        intFrmNum = intFrmNum - 1
    Loop
End Sub

Private Sub Label11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim XXX As String
'Newed
List3.Clear
'List3.AddItem "Cancel"
Open App.Path & "\PatientList.txt" For Input As #1
    Do While Not EOF(1)
        Line Input #1, XXX
        List3.AddItem XXX
    Loop
Close #1
List3.Visible = True
List3.SetFocus
End Sub

Private Sub Label4_Change()
Label12.Caption = Label4.Caption
End Sub

Private Sub List1_DblClick()
Dim n As Node
    TreeView1.Enabled = True
    UnSubclass TreeView1
    TreeView1.Visible = False
    For Each n In TreeView1.Nodes
        'MsgBox n.Text
        If n.Text = List2.List(List1.ListIndex) Then
            'mnuTopp_Click
            n.Expanded = True
            n.ForeColor = &HFF&
        Else
            n.Expanded = False
            n.ForeColor = &H80000012
        End If
    Next
    Picture2.Visible = False
Subclass Me, TreeView1
TreeView1.Visible = True

End Sub

Private Sub List3_DblClick()
If List3.List(List3.ListIndex) = "Cancel" Then
    List3.Visible = False
    Exit Sub
End If
Text4.Text = List3.List(List3.ListIndex)
List3.Visible = False
Command6.Enabled = False
Text5.SetFocus

End Sub

Private Sub mnuAAFP_Click()
On Error Resume Next
StartDoc "http://www.aafp.org/online/en/home.html"
End Sub

Private Sub mnuAddDose_Click()
Dim txt As String
Dim new_node As Node

    txt = InputBox("Text", "Add Node", "")
    txt = txt & " (Rx)"
    If Len(txt) > 0 Then
        If TreeView1.SelectedItem Is Nothing Then
            Set new_node = TreeView1.Nodes.Add(, , , txt)
        Else
            Set new_node = TreeView1.Nodes.Add( _
                TreeView1.SelectedItem, tvwChild, , txt)
        End If
        new_node.EnsureVisible
    End If

End Sub

Private Sub mnuAddFav_Click()
Open App.Path & "\Favorites.txt" For Append As #1
    Print #1, vbTab & TreeView1.SelectedItem.Parent
    Print #1, vbTab & vbTab & TreeView1.SelectedItem.Text
Close #1
End Sub

Private Sub mnuAddPE_Click()
Dim txt As String
Dim new_node As Node

    txt = InputBox("Text", "Add Node", "")
    txt = txt & " (ED)"
    If Len(txt) > 0 Then
        If TreeView1.SelectedItem Is Nothing Then
            Set new_node = TreeView1.Nodes.Add(, , , txt)
        Else
            Set new_node = TreeView1.Nodes.Add( _
                TreeView1.SelectedItem, tvwChild, , txt)
        End If
        new_node.EnsureVisible
    End If

End Sub

Private Sub mnuBlank_Click()

End Sub

Private Sub mnuAllDrugs_Click()
UnSubclass TreeView1
TreeView1.Visible = False
mnuSaveFave.Visible = False
'FileCopy App.Path & "\test.txt", App.Path & "\Favorites.txt"
'FileCopy App.Path & "\test.bak", App.Path & "\test.txt"
Favorites = False
mnuFavList.Enabled = True
mnuAllDrugs = False

    file_name = App.Path
    If Right$(file_name, 1) <> "\" Then file_name = file_name & "\"
    DoEvents: DoEvents
    file_name = file_name & "test.txt"
    LoadTreeViewFromFile file_name, TreeView1
    DoEvents: DoEvents
    mnuCollapse_Click
Subclass Me, TreeView1
TreeView1.Visible = True
mnuTopp_Click
End Sub

Private Sub mnuAnnalsFP_Click()
On Error Resume Next
StartDoc "http://www.annfammed.org/"

End Sub

Private Sub mnuArchives_Click()
On Error Resume Next
StartDoc "http://archinte.ama-assn.org/"
End Sub

Private Sub mnuATF_Click()
    mnuAddFav_Click
End Sub

Private Sub mnuBlankMultiRx_Click()
    Newed
    Option1(1).Value = True
    WriteBlank = True
    WriteBlankSingle = False
    WriteBlankMultiple = True
    Option1(0).Enabled = False
    Option1(1).Enabled = False
    Form1.TreeView1.Enabled = False
    mnuPrint.Enabled = True
    Load Form2
    mnuSaveRx.Enabled = False
    Form2.mnuWriteRx_Click
    Form2.Show
    Picture1.Visible = False
    TreeView1.Enabled = False
    Me.Visible = False
    mnuNewRx.Enabled = True

End Sub

Private Sub mnuBlankSingle_Click()
    Newed
    Option1(0).Value = True
    WriteBlank = True
    WriteBlankSingle = True
    WriteBlankMultiple = False
    Option1(0).Enabled = False
    Option1(1).Enabled = False
    Form1.TreeView1.Enabled = False
    mnuPrint.Enabled = True
    Load RxForm2
    mnuSaveRx.Enabled = False
    RxForm2.mnuWriteRx_Click
    TreeView1.Enabled = False
    RxForm2.Show
    Picture1.Visible = False
    mnuNewRx.Enabled = True
    Me.Visible = False
End Sub

Private Sub mnuBottomm_Click()
    TreeView1.SetFocus
    SendKeys "{END}"

End Sub

Private Sub mnuCochrane_Click()
On Error Resume Next
StartDoc "http://www.thecochranelibrary.com/view/0/index.html"
End Sub

Public Sub mnuCollapse_Click()
UnSubclass TreeView1
TreeView1.Visible = False
Dim n As Node
    For Each n In TreeView1.Nodes
        If n.Expanded Then n.Expanded = False
        n.ForeColor = &H80000012
    Next
Subclass Me, TreeView1
TreeView1.Visible = True
End Sub

Private Sub mnuEpocrates_Click()
On Error Resume Next
StartDoc "https://online.epocrates.com"

End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuExpand_Click()
UnSubclass TreeView1
TreeView1.Visible = False
Dim n As Node
    For Each n In TreeView1.Nodes
        n.Expanded = True
        n.ForeColor = &H80000012
    Next
mnuTopp_Click
Subclass Me, TreeView1
TreeView1.Visible = True
End Sub

Private Sub mnuExternal_Click()
mnuUndoRx.Visible = True
FileCopy App.Path & "\test.txt", App.Path & "\Previoustest.txt"
StartDoc App.Path & "\test.txt"
End Sub

Private Sub mnuExtPatients_Click()
mnuUndoPt.Visible = True
FileCopy App.Path & "\PatientList.txt", App.Path & "\PreviousPatientList.txt"
StartDoc App.Path & "\PatientList.txt"
End Sub

Private Sub mnuFavList_Click()
UnSubclass TreeView1
mnuSaveFave.Visible = True
TreeView1.Visible = False
    FileCopy App.Path & "\test.txt", App.Path & "\test.bak"
    'FileCopy App.Path & "\Favorites.txt", App.Path & "\test.txt"
    mnuFavList.Enabled = False
    mnuAllDrugs = True
    Favorites = True
    file_name = App.Path
    If Right$(file_name, 1) <> "\" Then file_name = file_name & "\"
    DoEvents: DoEvents
    file_name = file_name & "Favorites.txt"
    LoadTreeViewFromFile file_name, TreeView1
    DoEvents: DoEvents
    'mnuCollapse_Click
Subclass Me, TreeView1
TreeView1.Visible = True
mnuTopp_Click
End Sub

Public Sub mnuFile_Click()

End Sub

Private Sub mnuInstruct_Click()
On Error Resume Next
StartDoc App.Path & "\Help\Help.htm"

End Sub

Private Sub mnuInteract_Click()
On Error Resume Next
StartDoc "http://reference.medscape.com/drug-interactionchecker"

End Sub

Private Sub mnuJAMA_Click()
On Error Resume Next
StartDoc "http://jama.ama-assn.org/"
End Sub

Private Sub mnuMayo_Click()
On Error Resume Next
StartDoc "http://www.mayoclinicproceedings.org/"

End Sub

Private Sub mnuMD_Click()
UnSubclass TreeView1
mnuSaveFave.Visible = True
TreeView1.Visible = False
    ShiftNode TreeView1, TreeView1.SelectedItem, CLng(2)
Subclass Me, TreeView1
TreeView1.Visible = True
End Sub

Private Sub mnuMD1_Click()
On Error Resume Next
StartDoc "http://search.medscape.com/reference-search?newSearchHeader=1&queryText=" & PatientEducation

End Sub

Private Sub mnuMedscape_Click()
On Error Resume Next
StartDoc "http://search.medscape.com/reference-search?newSearchHeader=1&queryText=" & PatientEducation

End Sub

Private Sub mnuMedscape1_Click()
On Error Resume Next
StartDoc "http://reference.medscape.com/drugs"

End Sub

Private Sub mnuMoveDown_Click()
    ShiftNode TreeView1, TreeView1.SelectedItem, CLng(2)

End Sub

Private Sub mnuMoveUp_Click()
    ShiftNode TreeView1, TreeView1.SelectedItem, CLng(3)

End Sub

Private Sub mnuMU_Click()
UnSubclass TreeView1
mnuSaveFave.Visible = True
TreeView1.Visible = False
    ShiftNode TreeView1, TreeView1.SelectedItem, CLng(3)
Subclass Me, TreeView1
TreeView1.Visible = True
End Sub

Private Sub mnuNEJM_Click()
On Error Resume Next
StartDoc "http://www.nejm.org/"

End Sub

Private Sub mnuNewPatient_Click()
Dim intsave As Integer
If Saved = False Then 'Or PrintedFlag = False Then
    intsave = MsgBox("Do you want to Save the Rx?", _
                     vbYesNoCancel + vbExclamation)
    Select Case intsave
        Case vbYes
            mnuSaveRx_Click
            Exit Sub
    End Select
End If
NewedPt
End Sub

Private Sub mnuNewRx_Click()
Dim intsave As Integer
If Saved = False Then 'Or PrintedFlag = False Then
    intsave = MsgBox("Do you want to Save the Rx?", _
                     vbYesNoCancel + vbExclamation)
    Select Case intsave
        Case vbYes
            mnuSaveRx_Click
            Exit Sub
    End Select
End If
Newed
End Sub
Sub Newed()
    RxNumber = 0
    WriteBlank = False
    WriteBlankSingle = False
    WriteBlankMultiple = False
    'Clear all data
    Unload Form2
    Set Form2 = Nothing
    Unload RxForm2
    Set RxForm2 = Nothing
    mnuCollapse_Click
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    'Text4.Text = ""
    'Text5.Text = ""
    mnuNewRx.Enabled = True
    mnuNewPatient.Enabled = True
    mnuSaveRx.Enabled = False
    mnuPtEd1.Enabled = False
    Saved = True
    OpenFlag = False
    PrintedFlag = False
    TreeView1.Enabled = True
    WriteOnFlag = False
    Command3.Enabled = True
    Option1(1).Enabled = True
    Option1(0).Enabled = True
    Form1.mnuRx.Enabled = True
    Form1.mnuEdit.Enabled = True
    Form1.mnuView.Enabled = True
    MergedRx = False
End Sub
Sub NewedPt()
    RxNumber = 0
    WriteBlank = False
    WriteBlankSingle = False
    WriteBlankMultiple = False
    'Clear all data
    Unload Form2
    Set Form2 = Nothing
    Unload RxForm2
    Set RxForm2 = Nothing
    mnuNewRx.Enabled = True
    mnuNewPatient.Enabled = True
    mnuCollapse_Click
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    mnuSaveRx.Enabled = False
    mnuPtEd1.Enabled = False
    Saved = True
    OpenFlag = False
    PrintedFlag = False
    TreeView1.Enabled = True
    WriteOnFlag = False
    Form1.mnuRx.Enabled = True
    Form1.mnuEdit.Enabled = True
    Form1.mnuView.Enabled = True
    Command3.Enabled = True
    Option1(1).Enabled = True
    Option1(0).Enabled = True
    MergedRx = False
End Sub
Private Sub mnuOpenRx_Click()
    On Error GoTo ErrHandler
    Dim ReturnValuePath As String, i As Integer, XX As String, ii As Integer
    Dim intsave As Integer
    If Saved = False Then   'Or PrintedFlag = False Then
        If RxNumber <> 0 Then
            intsave = MsgBox("Do you want to Save/Print the Rx?", _
                             vbYesNoCancel + vbExclamation)
            Select Case intsave
                Case vbYes
                    mnuSaveRx_Click
                    Exit Sub
            End Select
        End If
    End If
'    Unload Form2
'    Unload RxForm2
'    mnuCollapse_Click
'    Text1.Text = ""
'    Text2.Text = ""
'    Text3.Text = ""
'    Text4.Text = ""
'    Text5.Text = ""
'    Text6.Text = ""
'    Command3.Enabled = False
'    WriteBlank = False
'    TreeView1.Enabled = False
'    Saved = False
'    OpenFlag = False
'    PrintedFlag = False
'    mnuNewRx.Enabled = True
'    Option1(1).Enabled = False
'    Option1(0).Enabled = False
'    Command3.Enabled = False
'    WriteOnFlag = False
    NewedPt
    SetTopMostWindow Me.hwnd, False
    RxNumber = 0
    cdlOpen.CancelError = True
    cdlOpen.Flags = cdlOFNHideReadOnly
    cdlOpen.Filter = "Rx Files  (*.rx;*.rx1)|*.rx;*.rx1"

    cdlOpen.InitDir = App.Path
    'cdlOpen.FilterIndex = 2
    cdlOpen.ShowOpen
    If Trim(cdlOpen.FileName) <> "" Then
        Open cdlOpen.FileName For Input As #1
            Select Case LCase(Right(cdlOpen.FileName, 3))
                Case "rx1"  'Single
                    oRx = False
                    Option1(0).Value = True
                    'Option1(1).Enabled = False
                    Do While Not EOF(1)
                        DoEvents: DoEvents
                       Saved = True
                        Line Input #1, XX: Text4.Text = XX
                        RxForm2.Named.Caption = Trim(Text4.Text)
                        Line Input #1, XX: Text5.Text = XX
                        RxForm2.Aged.Caption = Trim(Text5.Text)
                        Line Input #1, XX: Text6.Text = XX
                        RxForm2.Dated.Caption = Trim(Text6.Text)
                        Line Input #1, XX: RxForm2.Drug1(0).Text = XX
                        Line Input #1, XX: RxForm2.Sig1(0).Text = XX
                        Line Input #1, XX: RxForm2.Comment.Text = XX
                        Line Input #1, XX: RxForm2.Num1(0).Caption = XX: Text1.Text = XX
                        Line Input #1, XX: RxForm2.Ref1(0).Caption = XX: Text2.Text = XX
                        Line Input #1, XX: RxForm2.Sub1(0).Caption = XX
                        If Trim(RxForm2.Drug1(0).Text) <> "" Then
                            Command1.Enabled = True
                            mnuPrint.Enabled = True
                            'mnuSaveRx.Enabled = True
                            RxNumber = RxNumber + 1
                            Label4.Caption = RxForm2.Drug1(0).Text & " " & RxForm2.Sig1(0).Text
                            If RxForm2.Sub1(0) = "Yes" Then
                                Check1.Value = 0
                            Else
                               Check1.Value = 1
                            End If
                        End If
                    Loop
                Case ".rx"   'Multi
                    oRx = True
                    Option1(0).Value = False
                    Option1(1).Enabled = False
                    Saved = True
                    Do While Not EOF(1)
                        Line Input #1, XX: Text4.Text = XX
                        Form2.Named.Caption = Trim(Text4.Text)
                        Line Input #1, XX: Text5.Text = XX
                        Form2.Aged.Caption = Trim(Text5.Text)
                        Line Input #1, XX: Text6.Text = XX
                        Form2.Dated.Caption = Trim(Text6.Text)
                        Line Input #1, XX: Form2.Comment.Text = XX: Text3.Text = XX
                        For ii = 0 To 7
                            Line Input #1, XX: Form2.Drug1(ii).Text = XX
                            Line Input #1, XX: Form2.Sig1(ii).Caption = XX
                            Line Input #1, XX: Form2.Num1(ii).Caption = XX: Text1.Text = XX
                            Line Input #1, XX: Form2.Ref1(ii).Caption = XX: Text2.Text = XX
                            Line Input #1, XX: Form2.Num1(ii).Caption = XX
                            Line Input #1, XX: Form2.Sub1(ii).Caption = XX
                            If Trim(Form2.Drug1(ii).Text) <> "" Then
                                Command1.Enabled = True
                                mnuPrint.Enabled = True
                                mnuSaveRx.Enabled = False
                                RxNumber = RxNumber + 1
                                Label4.Caption = Form2.Drug1(ii).Text & " " & Form2.Sig1(ii).Caption
                                If Form2.Sub1(ii) = "Yes" Then
                                    Check1.Value = 0
                                Else
                                   Check1.Value = 1
                                End If
                            End If
                        Next
                    Loop

            End Select
        Close #1
        If LCase(Right(cdlOpen.FileName, 4)) = ".rx1" Then
            Option1(1).Enabled = False
            Option1(0) = False
        Else
            Option1(1).Enabled = False
            Option1(0) = False
        End If
        Saved = True
        'mnuPrint_Click
        mnuSaveRx.Enabled = False
        If Option1(0).Value = True Or oRx = False Then
            Load RxForm2
            RxForm2.Show
        Else
            Load Form2
            Form2.Show
        End If
        OpenFlag = True
        Command1_Click
    Else
        MsgBox "Please Select an Rx!"
    End If

If mnuTop.Checked = True Then SetTopMostWindow Me.hwnd, True

Exit Sub
ErrHandler:
Close #1
If mnuTop.Checked = True Then SetTopMostWindow Me.hwnd, True
Saved = True
'mnuPrint_Click

mnuSaveRx.Enabled = False
If Option1(0).Value = True Or oRx = False Then
    Load RxForm2
    RxForm2.Show
Else
    Load Form2
    Form2.Show
End If
OpenFlag = True
Command1_Click

End Sub

Private Sub mnuPopupAddNode_Click()
Dim txt As String
Dim new_node As Node

    txt = InputBox("Text", "Add Node", "")
    If Len(txt) > 0 Then
        If TreeView1.SelectedItem Is Nothing Then
            Set new_node = TreeView1.Nodes.Add(, , , txt)
        Else
            Set new_node = TreeView1.Nodes.Add( _
                TreeView1.SelectedItem, tvwChild, , txt)
        End If
        new_node.EnsureVisible
    End If
End Sub

Private Sub mnuPopupDeleteNode_Click()
    TreeView1.Nodes.Remove TreeView1.SelectedItem.Index
End Sub


Private Sub mnuPrint_Click()
Command1.Caption = "Print"
mnuCollapse_Click
Picture1.Visible = True
Picture2.Visible = False
Picture1.SetFocus
TreeView1.Enabled = False
If WriteBlankSingle = True Then
    Option1(0).Enabled = False
    Option1(1).Enabled = False
    Form1.TreeView1.Enabled = False
    RxForm2.Show
    Exit Sub
End If
If WriteBlankMultiple = True Then
    Option1(0).Enabled = False
    Option1(1).Enabled = False
    Form1.TreeView1.Enabled = False
    Form2.Show
    Exit Sub
End If
'MsgBox oRx
If oRx = True And oRx1 = False Then
    Option1(0).Value = 1
    Option1(1).Value = 0
    'Option1(0).Enabled = True
    Option1(1).Enabled = False
End If
If oRx = False And oRx1 = True Then
    Option1(0).Value = 0
    Option1(1).Value = 1
    'Option1(1).Enabled = True
    Option1(0).Enabled = False
End If
End Sub

Private Sub mnuPrintED_Click()

End Sub

Private Sub mnuPRx_Click()
Command1.Caption = "Print"
Picture1.Visible = True
Picture1.SetFocus

End Sub

Private Sub mnuPtEd1_Click()
On Error Resume Next
Dim PtEdd As String
PtEdd = App.Path & "\Patient Education\" & PatientEducation & ".pdf"
StartDoc PtEdd


End Sub

Private Sub mnuPtEdFold_Click()
On Error Resume Next

      If shlShell Is Nothing Then
          Set shlShell = New shell32.Shell
      End If
    shlShell.Explore ("C:\1down\OfficeVisitorPSC\PatientEducation\")
End Sub

Private Sub mnuPubmed_Click()
On Error Resume Next
StartDoc "http://www.ncbi.nlm.nih.gov/pubmed/"

End Sub

Private Sub mnuRefresh_Click()
Dim file_name As String

UnSubclass TreeView1
TreeView1.Visible = False
Dim n As Node
    'file_name = App.Path
    'If Right$(file_name, 1) <> "\" Then file_name = file_name & "\"
    DoEvents: DoEvents
    'file_name = file_name & "test.txt"
    LoadTreeViewFromFile file_name, TreeView1
    DoEvents: DoEvents
    DoEvents: DoEvents
    For Each n In TreeView1.Nodes
        If n.Expanded Then n.Expanded = False
    Next
    Text6.Text = Date
    DoEvents: DoEvents
Subclass Me, TreeView1
TreeView1.Visible = True
mnuTopp_Click
End Sub

Private Sub mnuSage_Click()
On Error Resume Next
StartDoc "http://online.sagepub.com/"

End Sub

Private Sub mnuSaveFave_Click()
    file_name = App.Path
    If Right$(file_name, 1) <> "\" Then file_name = file_name & "\"
    file_name = file_name & "Favorites.txt"
    SaveTreeViewIntoFile file_name, TreeView1

End Sub

Private Sub mnuSaveRx_Click()
    On Error GoTo ErrHandler
    Dim ReturnValuePath As String, i As Integer, Filer As String
    Dim intsave As Integer
    SetTopMostWindow Me.hwnd, False
    
    
    cdlOpen.CancelError = True
    cdlOpen.Flags = cdlOFNHideReadOnly
    cdlOpen.Filter = "MultiRx Files (*.Rx)|*.Rx|Single Rx Files(*.rx1)|*.rx1"
    If Option1(0).Value = True Then
        cdlOpen.FilterIndex = 2
    Else
        cdlOpen.FilterIndex = 1
    End If
    cdlOpen.InitDir = App.Path
here:
    cdlOpen.ShowSave
    'MsgBox cdlOpen.FileName
    If Option1(0).Value = True And LCase(Right(cdlOpen.FileName, 4)) <> ".rx1" Then
        MsgBox "You MUST save this file with an .rx1 extension!"
        Exit Sub
    End If
    If Option1(0).Value = False And LCase(Right(cdlOpen.FileName, 3)) <> ".rx" Then
        MsgBox "You MUST save this file with an .rx extension!"
        Exit Sub
    End If
    
    If Dir(cdlOpen.FileName) <> "" Then
        If Saved = False Then   'Or PrintedFlag = False Then
                intsave = MsgBox("Do you want to Overwrite " & Filer & "?", _
                                 vbYesNoCancel + vbExclamation)
                Select Case intsave
                    Case vbNo
                        GoTo here
                    Case vbCancel
                        Exit Sub
                End Select
        End If
    End If
    If Trim(cdlOpen.FileName) <> "" Then
        Select Case Option1(0).Value
            Case False   'Multi .rx
                'If Dir(App.Path & "\Savedit") <> "" Then SaveSub
                FileCopy App.Path & "\Savedit", cdlOpen.FileName
                Kill App.Path & "\Savedit"
                Saved = True
            Case True   'single .rx1
                'If Dir(App.Path & "\Savedit") <> "" Then SaveSub
                FileCopy App.Path & "\Savedit", cdlOpen.FileName
                Kill App.Path & "\Savedit"
                Saved = True
        End Select
           
    Else
        MsgBox "Please Provide a Name for the Rx!"
    End If
  

Exit Sub
ErrHandler:
Close #1
    'mnuCollapse_Click

'Command1.Caption = "Save"
'Picture1.Visible = True
'Picture1.SetFocus
End Sub

Public Sub SaveSub()
Dim i As Integer
        Select Case Option1(0).Value
            Case False   'Multi .rx
                Open App.Path & "\Savedit" For Output As #1
                    Print #1, Text4.Text
                    Print #1, Text5.Text
                    Print #1, Text6.Text
                    Print #1, Form2.Comment.Text
                    For i = 0 To 7
                        Print #1, Form2.Drug1(i).Text
                        Print #1, Form2.Sig1(i).Caption
                        Print #1, Form2.Num1(i).Caption
                        Print #1, Form2.Ref1(i).Caption
                        Print #1, Form2.Num1(i).Caption
                        Print #1, Form2.Sub1(i).Caption
                    Next
                Close #1
            Case True   'single .rx1
                 Open App.Path & "\Savedit" For Output As #1
                    Print #1, Text4.Text
                    Print #1, Text5.Text
                    Print #1, Text6.Text
                    Print #1, RxForm2.Drug1(0).Text
                    Print #1, RxForm2.Sig1(0).Text
                    Print #1, RxForm2.Comment.Text
                    Print #1, RxForm2.Ref1(0).Caption
                    Print #1, RxForm2.Num1(0).Caption
                    Print #1, RxForm2.Sub1(0).Caption
                Close #1
        End Select
End Sub
Private Sub mnuSearch_Click()
mnuSrch_Click
End Sub

Private Sub mnuSearchRx_Click()
    mnuSrch_Click
End Sub

Private Sub mnuSendEmail_Click()
On Error Resume Next
Dim Message, Title, Default, MyValue
Message = "An Email Address"   ' Set prompt.
Title = "Send Email"   ' Set title.
Default = "anyone@anywhere.com"   ' Set default.
' Display message, title, and default value.
MyValue = InputBox(Message, Title, Default)

'MyValue = InputBox(Message, Title, Default, 100, 100)

StartDoc "mailto:" & MyValue & "?SUBJECT=VIM: Dr. Warren Goff " & Date

'     World", Normal


End Sub

Private Sub mnuSignedRx_Click()
    Newed
    UnSignedRx = False
    SignedRx = True
End Sub

Private Sub mnuSrch_Click()
Picture2.Visible = True
Picture1.Visible = False
Text7.Text = "Search Term"
TreeView1.Enabled = False
End Sub

Private Sub mnuSRx1_Click()
mnuSrch_Click
End Sub

Private Sub mnuTop_Click()
UnSubclass TreeView1
mnuSaveFave.Visible = True
TreeView1.Visible = False
    If mnuTop.Checked = False Then
        mnuTop.Checked = True
        SetTopMostWindow Me.hwnd, True
    Else
        mnuTop.Checked = False
        SetTopMostWindow Me.hwnd, False
    End If
UnSubclass TreeView1
mnuSaveFave.Visible = True
TreeView1.Visible = False
End Sub

Private Sub mnuTopp_Click()
    TreeView1.SetFocus
    SendKeys "^{HOME}"
    SendKeys "{HOME}"

End Sub

Private Sub mnuUndoPt_Click()
mnuUndoPt.Visible = False
FileCopy App.Path & "\PreviousPatientList.txt", App.Path & "\PatientList.txt"

End Sub

Private Sub mnuUndoRx_Click()
mnuUndoRx.Visible = False
FileCopy App.Path & "\Previoustest.txt", App.Path & "\test.txt"

End Sub

Private Sub mnuUnsignedRx_Click()
    Newed
    UnSignedRx = True
    SignedRx = False

End Sub

Private Sub mnuWalmartOTC_Click()
On Error Resume Next
StartDoc App.Path & "\Walmart OTC.pdf"
End Sub

Private Sub mnuWalmartRx_Click()
On Error Resume Next
StartDoc App.Path & "\Walmart Prescription medications.pdf"
End Sub

Private Sub mnuWebsite_Click()
On Error Resume Next
StartDoc "http://www.warrengoff.com"

End Sub

Private Sub Picture1_LostFocus()
'Picture1.Visible = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = IIf(Not KeyAscii = 8 And Not Val((Chr(KeyAscii))) > 0, 0, KeyAscii)

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = IIf(Not KeyAscii = 8 And Not Val((Chr(KeyAscii))) > 0, 0, KeyAscii)

End Sub

Private Sub Text4_Change()
If Trim(Text4.Text) <> "" Then
    Command6.Enabled = True
Else
    Command6.Enabled = False
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
KeyAscii = IIf(Not KeyAscii = 8 And Not Val((Chr(KeyAscii))) > 0, 0, KeyAscii)

End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command4_Click
 
End Sub

Private Sub Text7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Text7.Text = "Search Term" Then Text7.Text = ""
End Sub

Private Sub Text8_Change()
Text8.Text = Replace(Text8.Text, " (Rx)", "")
If Len(Text8) > 33 Then
    Option1(0).Value = True
    Option1(1).Enabled = False
Else
    Option1(0).Enabled = True
    Option1(1).Enabled = True
    Option1(1).Value = True
End If
SigChanged = True
End Sub

Private Sub TreeView1_Click()
'SendKeys "{HOME}"
'DoEvents: DoEvents
If Favorites = True Then mnuAddFav.Enabled = False: Exit Sub
If InStr(TreeView1.SelectedItem.Text, "Rx") > 0 Then
    mnuAddFav.Enabled = True
    mnuATF.Enabled = True
Else
    mnuAddFav.Enabled = False
    mnuATF.Enabled = False
End If
End Sub

Private Sub TreeView1_DblClick()
If InStr(TreeView1.SelectedItem.Text, "Rx") > 0 Then
    Command1.Caption = "Print"
    Saved = False
    TreeView1.Enabled = False
    Picture1.Visible = True
    Picture1.SetFocus
End If
End Sub

Private Sub TreeView1_KeyPress(KeyAscii As Integer)
If KeyAscii = 21 Then  'u 117
    mnuMoveUp_Click
End If
If KeyAscii = 4 Then  'd
    mnuMoveDown_Click
End If
If KeyAscii = 24 Then  'x
    mnuPopupDeleteNode_Click
End If
End Sub

Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Picture1.Visible = False'
'List3.Visible = False
'Picture2.Visible = False
'List1.Clear
'Text7.Text = "Search Term"
mnuMD.Enabled = True
mnuMU.Enabled = True
    If Button = vbRightButton Then
        Set TreeView1.SelectedItem = TreeView1.HitTest(X, Y)
        PopupMenu mnuPopup
    End If
End Sub


Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
'On Error Resume Next
Dim Y As Integer, nam As String
Dim PtEdd As String
mnuPtEd1.Enabled = False
If TreeView1.Nodes.Count > 0 Then 'prevents an error if TV is empty
Y = TreeView1.SelectedItem.Index
'Label1.Caption = Str(y)
    If InStr(TreeView1.SelectedItem.Text, "Rx") > 0 Then
        mnuPrint.Enabled = True
        mnuPRx.Enabled = True
        mnuSaveRx.Enabled = True
        OpenFlag = False
        'mnuPrintED.Enabled = False
        nam = TreeView1.SelectedItem.Parent & ": " & TreeView1.SelectedItem.Text
        nam = Trim(Replace(nam, "(Rx)", ""))
        PatientEducation = LCase(TreeView1.SelectedItem.Parent)
        PtEdd = App.Path & "\Patient Education\" & PatientEducation & ".pdf"
        If Dir(PtEdd) <> "" Then
            mnuPtEd1.Enabled = True
        Else
            mnuPtEd1.Enabled = False
        End If
        Label4.Caption = TreeView1.SelectedItem.Parent
        Text8.Text = TreeView1.SelectedItem.Text
        SigChanged = False
        Me.Caption = nam '& " - " & Str(y)
        If Option1(1).Value = True Then Command3.Enabled = True
        Exit Sub
    Else
        Me.Caption = ""
    End If
    If InStr(TreeView1.SelectedItem.Text, "ED") > 0 Then
        If RxNumber = 0 Then mnuPrint.Enabled = False
        mnuPRx.Enabled = False
        'mnuSaveRx.Enabled = False
        'mnuPrintED.Enabled = True
        nam = TreeView1.SelectedItem.Parent & ": " & TreeView1.SelectedItem.Text
        Me.Caption = Trim(Replace(nam, "(ED)", "")) '& " - " & Str(y)
        Exit Sub
    Else
        Me.Caption = ""
    End If
    If RxNumber = 0 Then mnuPrint.Enabled = False
    mnuPRx.Enabled = False
    'mnuSaveRx.Enabled = False
    'mnuPrintED.Enabled = False
Else
    Set TreeView1.DropHighlight = TreeView1.SelectedItem
End If
End Sub
