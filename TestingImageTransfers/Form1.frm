VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3375
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   3720
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3240
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Image Image2 
      Height          =   2415
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   2415
      Left            =   120
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private TheSize As Long
Private AllData As Variant

Private Sub Command1_Click()
On Error Resume Next
Winsock1.Close
Winsock1.LocalPort = "801"
Winsock1.Listen
Winsock2.Close
Winsock2.Connect "127.0.0.1", 801
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Winsock1.Close
Winsock1.Accept requestID
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim Data As Variant
Dim PropBag2 As New PropertyBag  'property bag to store the data
Dim byteArr2() As Byte
Winsock1.GetData Data
If InStr(1, Data, "||SIZE||") > 0 Then
TheSize = Split(Data, "||SIZE||")(0)
AllData = Split(Data, "||SIZE||")(1)
Else
AllData = AllData & Data
End If
If Len(AllData) >= TheSize Then
byteArr2 = AllData
Debug.Print "Received: " & AllData
PropBag2.Contents = byteArr2
Image2.Picture = PropBag2.ReadProperty("Pic")
Winsock1.Close
End If
End Sub

Private Sub Winsock2_Connect()
Dim PropBag As New PropertyBag  'property bag to store the data
Dim TheData As Variant
Dim byteArr() As Byte
Dim Lengh As Long
PropBag.WriteProperty "Pic", Image1.Picture
byteArr = PropBag.Contents
TheData = byteArr
Lengh = Len(TheData)
TheData = Lengh & "||SIZE||" & TheData
Debug.Print TheData
Winsock2.SendData TheData
End Sub

