VERSION 5.00
Begin VB.Form HostList 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Viewers List"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2670
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   2670
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Double Click To Kick A Viewer"
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   15
      Width           =   2655
   End
End
Attribute VB_Name = "HostList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Me.Icon = Host.Icon
End Sub

Private Sub List1_DblClick()
On Error Resume Next
If List1.ListCount = 0 Then Exit Sub
Dim I As Integer
Dim WhoKick As String
WhoKick = List1
For I = 1 To 50
If Host.CamSocket(I).TheName = WhoKick Then
Host.CamSocket(I).ForceStop
GoTo Done
End If
Next I
DoEvents
Done:
List1.RemoveItem List1.ListIndex
DoEvents
Host.Caption = "WebCam (" & List1.ListCount & " Viewing)"
End Sub
