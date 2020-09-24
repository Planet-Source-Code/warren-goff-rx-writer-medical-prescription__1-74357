Attribute VB_Name = "Drag"
Option Explicit

Global Dater As String, FirstName As String, LastName As String, Doctor As String

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Public Const HTCAPTION = 2
'Public Const WM_NCLBUTTONDOWN = &HA1
'declare for moving the form
Public Declare Function ReleaseCapture Lib "user32" () As Long
'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Global PrintedFlag As Boolean, WriteOnFlag As Boolean, OpenFlag As Boolean
Global WriteBlank As Boolean, SignedRx As Boolean, UnSignedRx As Boolean
Global WriteBlankSingle As Boolean, WriteBlankMultiple As Boolean
Global MergedRx As Boolean, PatientEducation As String, FM As Integer
Global ExitFlag As Boolean
Public Sub Delay(HowLong As Date)
Dim TempTime As String
TempTime = DateAdd("s", HowLong, Now)
While TempTime > Now
DoEvents 'Allows windows to handle other stuff
Wend
End Sub

