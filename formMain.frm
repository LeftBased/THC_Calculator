VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form formMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "THC Calculator"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5685
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   5685
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cmnDiag 
      Left            =   240
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Save"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   2295
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   120
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Text            =   "25"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Calculate"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "THC%:"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "formMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Sub MakeTopMost(hwnd As Long)
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub
Private Sub Command1_Click()
On Error Resume Next
Dim xVal, resTHCperc, oneG, halfG, qrterG, oneTenth, HalfATenth, aHalfZ, anOz, TwoTenths As Double
Dim Output1 As String

xVal = Text1.Text
resTHCperc = xVal / 0.1
oneG = resTHCperc
halfG = oneG / 2
qrterG = oneG / 4
oneTenth = oneG / 10
TwoTenths = oneG / 5
HalfATenth = oneG / 20
aHalfZ = oneG * 14
anOz = oneG * 28
aqrter = oneG * 7
an8th = oneG * 3.5

Output1 = "THC Results:" + vbCrLf + "An Ounce (28g = 28000mg): " + CStr(anOz) + " mg" + vbCrLf + _
"A Half Ounce (14g = 14000mg): " + CStr(aHalfZ) + " mg" + vbCrLf + "A quarter ounce (7g = 7000mg): " + CStr(aqrter) + " mg" + vbCrLf & _
"An 8th (3.5g = 3500mg): " + CStr(an8th) + " mg" + vbCrLf + "One gram (1g = 1000mg): " + CStr(oneG) + " mg" + vbCrLf & _
"Half Gram (0.5g = 500mg): " + CStr(halfG) + " mg" + vbCrLf + "Quarter Gram (0.25g = 250mg): " + CStr(qrterG) + " mg" + vbCrLf + _
"One Tenth (0.1g = 100mg): " + CStr(oneTenth) + " mg" + vbCrLf + "Two Tenths (0.2g = 200 mg): " + CStr(TwoTenths) + " mg" + vbCrLf + _
"Half a Tenth (0.05g = 50mg): " + CStr(HalfATenth) + " mg"
Text2.Text = Output1
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim xFilename As String
cmnDiag.InitDir = App.Path()
cmnDiag.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
cmnDiag.ShowSave
If cmnDiag.FileName = "" Then
Else
xFilename = cmnDiag.FileName
Open xFilename For Output As #1
Print #1, Text2.Text 'CStr(Replace(Text2.Text, Chr(34), ""))
Close #1
MsgBox "File saved to: " + xFilename, vbCritical, ""
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
MakeTopMost Me.hwnd
End Sub
