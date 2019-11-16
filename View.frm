VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form View 
   BackColor       =   &H00000000&
   Caption         =   "CamView"
   ClientHeight    =   4470
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   5295
   FillColor       =   &H00E0E0E0&
   ForeColor       =   &H8000000B&
   Icon            =   "View.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7200
      Top             =   3720
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4215
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   8811
            Text            =   "Viewing Status"
            TextSave        =   "Viewing Status"
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock wsC 
      Left            =   6720
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Image Pic4 
      BorderStyle     =   1  'Fixed Single
      Height          =   4215
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5295
   End
   Begin VB.Menu mnu 
      Caption         =   "Menu"
      Begin VB.Menu paus 
         Caption         =   "Pause"
      End
      Begin VB.Menu jfdh 
         Caption         =   "-"
      End
      Begin VB.Menu clsit 
         Caption         =   "Close"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu setups 
      Caption         =   "Settings"
      Enabled         =   0   'False
   End
   Begin VB.Menu paus2d 
      Caption         =   "Paused!"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "View"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private TheIP As String
Private MyID As String
Private ThereID As String
Private sFile As String
Private lFileSize As Long, sFile2 As String
Private TimeOutCount As Long
'listen for a connection

Private Sub clsit_Click()
On Error Resume Next
  wsC.Close
  If Dir$(App.Path & "\Files\" & sFile & ".jpg") <> "" Then Kill App.Path & "\Files\" & sFile & ".jpg"
  DoEvents
  'deletes viewing
  End
End Sub

Private Sub StartView()
On Error Resume Next
    StatusBar1.Panels(1).Text = "Connecting To " & ThereID
    'setup for stop viewing
    wsC.Close
    wsC.Connect TheIP, 5905
    'connects
  'start viewing a cam
End Sub

Private Sub Form_Load()
On Error GoTo Error
    sLastSendOut = "N/A"
    paus.Enabled = False
    paus2d.Visible = False
    Pic4.Visible = False
    Wsize = Screen.Width / 15
    Hsize = Screen.Height / 15
    Debug.Print Wsize
    Debug.Print Hsize
    'Sratio = (100 / Wsize) * Hsize
    'Form_Resize
    TheIP = GetINI("ConnectTo")
    ThereID = GetINI("ViewingPassword")
    MyID = GetINI("ViewingPassword2")
    Me.Caption = ThereID
    sFile = RandomGen2("0123456789abcdefghijklmnopqrstuvwxyz", 9)
    WriteINI "ConnectTo", "127.0.0.1"
    Me.Show
    DoEvents
    StartView
Exit Sub
Error:
On Error Resume Next
End
End Sub

Private Sub Form_Resize()
On Error GoTo Error
If Sratio = "" Then
If Me.Width <= "4399" Then Me.Width = "4400": Me.Height = "4200": Exit Sub
If Me.Height <= "7499" Then Me.Height = Me.Width / 100 * 96: GoTo Skip
If Me.Height <= "9999" Then Me.Height = Me.Width / 100 * 92: GoTo Skip
If Me.Height <= "12499" Then Me.Height = Me.Width / 100 * 88: GoTo Skip
If Me.Height > "12499" Then Me.Height = Me.Width / 100 * 86
Skip:
If Me.Height <= "4199" Then Me.Width = "4400": Me.Height = "4200": Exit Sub
Pic4.Height = Me.Height - StatusBar1.Height - 870
Pic4.Width = Me.Width - 240
If Me.Height > "4250" And Me.Width > "4450" And Pic4.Left > "0" Then Pic4.Left = "0"
Else
If Me.Width <= "4399" Then Me.Width = "4400"
Me.Height = (Me.Width / 100) * Sratio
Pic4.Width = (Me.Width - 240)
Pic4.Height = (Me.Height - StatusBar1.Height - 870)
If Me.Width > "4399" And Pic4.Left > "0" Then Pic4.Left = "0"
End If
Exit Sub
Error:
Form_Resize2
End Sub

Private Sub Form_Resize2()
On Error Resume Next
If Sratio = "" Then
Pic4.Height = Me.Height - StatusBar1.Height - 870
Pic4.Width = Me.Height / 85 * 100
If ((Me.Width - Pic4.Width - 200) / 2) < 0 Then
Pic4.Left = "0"
Else
Pic4.Left = (Me.Width - Pic4.Width - 200) / 2
End If
Else
Pic4.Height = (Me.Height - StatusBar1.Height - 870)
Pic4.Width = (Pic4.Height / Sratio) * 100
If ((Me.Width - Pic4.Width - 200) / 2) < 0 Then
Pic4.Left = "0"
Else
Pic4.Left = (Me.Width - Pic4.Width - 200) / 2
End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
  wsC.Close
  If Dir$(App.Path & "\Files\" & sFile & ".jpg") <> "" Then Kill App.Path & "\Files\" & sFile & ".jpg"
  DoEvents
  'deletes viewing
  End
End Sub

Private Sub paus_Click()
If paus.Checked = True Then
paus.Checked = False
paus2d.Visible = False
Else
paus2d.Visible = True
paus.Checked = True
End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Timer1 = False
If wsC.State = 7 Then
TimeOutCount = TimeOutCount + 1
If TimeOutCount >= 5 Then
TimeOuntCOunt = 0
wsC.SendData "ID-NEXT"
End If
Timer1 = True
End If
End Sub

Private Sub wsC_Close()
  StatusBar1.Panels(1).Text = "Connection Closed. Session Ended!"
  If Dir$(App.Path & "\Files\" & sFile & ".jpg") <> "" Then Kill App.Path & "\Files\" & sFile & ".jpg"
  Pic4.Visible = False
  paus.Checked = False
  paus.Enabled = False
  paus2d.Visible = False
End Sub
'viewer closed

Private Sub wsC_Connect()
  StatusBar1.Panels(1).Text = "Connected! Authenticating.."
  Pic4.Visible = True
  paus.Enabled = True
  wsC.SendData "ID-PASS" & ThereID & "|~|~|" & MyID
End Sub
'viewer connect send pass

Private Sub wsC_DataArrival(ByVal bytesTotal As Long)
Dim sData As String
Static iPictures As Long, sLastRecieved As String
On Error GoTo Err

  wsC.GetData sData
  
    If Left$(sData, 3) = "ID-" Then
      sData = Mid$(sData, 4)
      Select Case Left$(sData, 4)
      Case "RES%"
      Sratio = Split(sData, "%")(1)
      Form_Resize
      Exit Sub
      
      Case "SIZE"
        TimeOutCount = 0
        If Timer1 = False Then Timer1 = True
        lFileSize = Split(Mid$(sData, 5), "FILE")(0)
        sFile2 = Split(sData, lFileSize & "FILE")(1)
        If Len(sFile2) = lFileSize Then
        If Dir$(App.Path & "\Files\" & sFile & ".jpg") <> "" Then Kill App.Path & "\Files\" & sFile & ".jpg"
          sLastRecieved = Time$
          'pic count + time
          StatusBar1.Panels(1).Text = "Last Image At " & sLastRecieved & " [" & lFileSize & "bytes]"
          'status
          Open App.Path & "\Files\" & sFile & ".jpg" For Binary As #2
          Put #2, , sFile2
          Close #2
          'prints to file
          If paus.Checked = False Then Pic4.Picture = LoadPicture(App.Path & "\Files\" & sFile & ".jpg")
          'show picture
SaveNext:
          Pause "0.1"
          wsC.SendData "ID-NEXT"
          Exit Sub
        'if file complete
        Else
          Exit Sub
        End If
      End Select
      
      Else
      
      sFile2 = sFile2 & sData
      
        If Len(sFile2) = lFileSize Then
        If Dir$(App.Path & "\Files\" & sFile & ".jpg") <> "" Then Kill App.Path & "\Files\" & sFile & ".jpg"
          sLastRecieved = Time$
          'pic count + time
          StatusBar1.Panels(1).Text = "Last Image At " & sLastRecieved & " [" & lFileSize & "bytes]"
          'status
          Open App.Path & "\Files\" & sFile & ".jpg" For Binary As #2
          Put #2, , sFile2
          Close #2
          'prints to file
          If paus.Checked = False Then Pic4.Picture = LoadPicture(App.Path & "\Files\" & sFile & ".jpg")
          'show picture
          Pause "0.1"
          wsC.SendData "ID-NEXT"
        'if file complete
        Exit Sub
        Else
        Exit Sub
        End If
      
    End If
    
Err:
  If Err.Number = 75 Then GoTo SaveNext
  'picture saving error, could not make directory
End Sub


Private Sub wsC_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  StatusBar1.Panels(1).Text = "Connection Error. Session Ended!"
  If Dir$(App.Path & "\Files\" & sFile & ".jpg") <> "" Then Kill App.Path & "\Files\" & sFile & ".jpg"
  Pic4.Visible = False
  paus.Checked = False
  paus.Enabled = False
  paus2d.Visible = False
End Sub
'error for whatever reason. error is generally could not connect or random happenings.
'close is usually when they kick you. so on error reconnect automatically

