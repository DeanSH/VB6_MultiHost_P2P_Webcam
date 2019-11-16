VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl CamSocket 
   BackColor       =   &H00FF0000&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1800
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1200
      Top             =   480
   End
   Begin MSWinsockLib.Winsock wsL 
      Left            =   240
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "CamSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event Stats(Data As String)
Private sFile As String
Private WhoIsThis As String
Private IsSending As Boolean

Private Sub Timer1_Timer()
On Error Resume Next
Timer1 = False
If WhoIsThis = "" Then
RaiseEvent Stats("Host: User Connection Timed Out!")
wsL.Close
End If
End Sub

Private Sub Timer2_Timer()
Timer2 = False
IsSending = False
ImageReady = False
End Sub

Private Sub wsL_Close()
Timer1 = False
If WhoIsThis = "" Then Exit Sub
RaiseEvent Stats("Host: " & WhoIsThis & " Stopped Viewing!")
RemoveList WhoIsThis, HostList.List1
DoEvents
WhoIsThis = ""
IsSending = False
ImageReady = False
Host.Caption = "WebCam (" & HostList.List1.ListCount & " Viewing)"
End Sub

Public Sub ForceStop()
On Error Resume Next
IsSending = False
ImageReady = False
wsL.Close
Timer1 = False
If WhoIsThis = "" Then
Host.Caption = "WebCam (" & HostList.List1.ListCount & " Viewing)"
Else
RemoveList WhoIsThis, HostList.List1
DoEvents
WhoIsThis = ""
Host.Caption = "WebCam (" & HostList.List1.ListCount & " Viewing)"
End If
End Sub

Public Function TheName() As String
TheName = WhoIsThis
End Function

Public Sub AcceptRequest(ByVal requestID As Long)
On Error Resume Next
RaiseEvent Stats("Host: User Authenticating..")
wsL.Close
Timer1 = True
wsL.Accept requestID
End Sub

Private Sub wsL_DataArrival(ByVal bytesTotal As Long)
On Error GoTo Error
Dim sData As String
Dim sWho As String
  wsL.GetData sData
  
    If Left$(sData, 3) = "ID-" Then
      sData = Mid$(sData, 4)
      Select Case Left$(sData, 4)
      Case "PASS"
      Timer1 = False
      IsSending = False
      ImageReady = False
        If InStr(1, sData, "SS" & WhoAmI & "|~|~|") > 0 Then
        WhoIsThis = Split(sData, WhoAmI & "|~|~|")(1)
        If WhoIsThis = "" Then GoTo Bad
        If InList(WhoIsThis, HostList.List1) = True Then
        RaiseEvent Stats("Host: Blocked Duplicate User Attempt!")
        WhoIsThis = ""
        wsL.Close
        Exit Sub
        End If
        HostList.List1.AddItem WhoIsThis
        RaiseEvent Stats("Host: " & WhoIsThis & " Started Viewing!")
        Host.Caption = "WebCam (" & HostList.List1.ListCount & " Viewing)"
        If Desktop = True Then
        wsL.SendData "ID-RES%" & Sratio
        DoEvents
        Pause "0.5"
        End If
        SendPicture
        Else
Bad:
        RaiseEvent Stats("Host: User Authentication Failed!")
        WhoIsThis = ""
        wsL.Close
        Exit Sub
        End If

      Case "NEXT"
        Pause "0.001"
        SendPicture

      End Select
    End If
Error:
End Sub
'parsing results

Private Sub wsL_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Timer1 = False
ImageReady = False
If WhoIsThis = "" Then Exit Sub
RaiseEvent Stats("Host: " & WhoIsThis & " Stopped Viewing!")
RemoveList WhoIsThis, HostList.List1
DoEvents
WhoIsThis = ""
IsSending = False
ImageReady = False
Host.Caption = "WebCam (" & HostList.List1.ListCount & " Viewing)"
End Sub
'duh

Private Sub SendPicture()
On Error GoTo Err
    'get jpg file
wait:
    'If sFile2 = "" Then GoTo wait
    If ImageReady = True Then Pause "0.001": GoTo wait
    IsSending = True
    ImageReady = True
    sFile = sFile2
    Timer2 = True
    wsL.SendData "ID-SIZE" & Len(sFile) & "FILE" & sFile
Exit Sub
Err:
ImageReady = False
IsSending = False
Debug.Print "SendPic Error"
End Sub

Private Sub wsL_SendComplete()
Timer2 = False
If IsSending = True Then
IsSending = False
ImageReady = False
End If
End Sub

Private Sub Pause(ByVal interval As String)
On Error Resume Next
Dim wait   As Single
  
  wait = Timer
  
  Do While Timer - wait < CSng(interval$)
     DoEvents
 Loop
End Sub
