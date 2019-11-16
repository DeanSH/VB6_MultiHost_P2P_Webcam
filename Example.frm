VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cam Usage Example"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   3615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Open And View Thier Cam?"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open And Host Your Cam?"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   3375
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "anthony"
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "deano"
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "127.0.0.1"
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "127.0.0.1"
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "<< Thier UserName"
      Height          =   255
      Index           =   3
      Left            =   2040
      TabIndex        =   7
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "<< Your UserName"
      Height          =   255
      Index           =   2
      Left            =   2040
      TabIndex        =   6
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "<< Thier IP Address"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   5
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "<< Your IP Address"
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long



Private Function GetINI(Key As String) As String
Dim Ret As String, NC As Long
  
  Ret = String(600, 0)
  NC = GetPrivateProfileString("P2PWebcam", Key, Key, Ret, 600, App.Path & "\Config.ini")
  If NC <> 0 Then Ret = Left$(Ret, NC)
  If Ret = Key Or Len(Ret) = 600 Then Ret = ""
  GetINI = Ret

End Function
'Read from INI

Private Sub WriteINI(ByVal Key As String, Value As String)
  
  WritePrivateProfileString "P2PWebcam", Key, Value, App.Path & "\Config.ini"

End Sub
'Write to INI
Private Sub Command1_Click()
On Error GoTo Error
WriteINI "HostPassword", Text1
WriteINI "ConnectTo", Text2
WriteINI "ViewingPassword2", Text3
WriteINI "ViewingPassword", Text4
WriteINI "Driver", "0"
DoEvents
Shell App.Path & "\CamHost.exe", vbNormalFocus
Error:
End Sub

Private Sub Command2_Click()
On Error GoTo Error
WriteINI "HostPassword", Text1
WriteINI "ConnectTo", Text2
WriteINI "ViewingPassword2", Text3
WriteINI "ViewingPassword", Text4
DoEvents
Shell App.Path & "\CamView.exe", vbNormalFocus
Error:
End Sub

Private Sub Form_Load()
Text1 = GetINI("HostPassword")
Text2 = GetINI("ConnectTo")
Text3 = GetINI("ViewingPassword2")
Text4 = GetINI("ViewingPassword")
End Sub
