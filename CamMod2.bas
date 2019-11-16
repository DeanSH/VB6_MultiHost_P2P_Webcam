Attribute VB_Name = "CamMod2"
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Wsize As String
Public Hsize As String
Public Sratio As String

Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INTERNET_OPEN_TYPE_PROXY = 3

Private Const scUserAgent = "VB Project"
Private Const INTERNET_FLAG_RELOAD = &H80000000

Private Declare Function InternetOpen Lib "wininet.dll" _
  Alias "InternetOpenA" (ByVal sAgent As String, _
  ByVal lAccessType As Long, ByVal sProxyName As String, _
  ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

Private Declare Function InternetOpenUrl Lib "wininet.dll" _
  Alias "InternetOpenUrlA" (ByVal hOpen As Long, _
  ByVal sUrl As String, ByVal sHeaders As String, _
  ByVal lLength As Long, ByVal lFlags As Long, _
  ByVal lContext As Long) As Long

Private Declare Function InternetReadFile Lib "wininet.dll" _
  (ByVal hFile As Long, ByVal sBuffer As String, _
   ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) _
  As Integer

Private Declare Function InternetCloseHandle _
   Lib "wininet.dll" (ByVal hInet As Long) As Integer





Type imgdes
    ibuff As Long
    stx As Long
    sty As Long
    endx As Long
    endy As Long
    buffwidth As Long
    palette As Long
    colors As Long
    imgtype As Long
    bmh As Long
    hBitmap As Long
    End Type


Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
    End Type


Declare Function bmpinfo Lib "VIC32.DLL" (ByVal Fname As String, bdat As BITMAPINFOHEADER) As Long
Declare Function allocimage Lib "VIC32.DLL" (image As imgdes, ByVal wid As Long, ByVal leng As Long, ByVal BPPixel As Long) As Long
Declare Function loadbmp Lib "VIC32.DLL" (ByVal Fname As String, desimg As imgdes) As Long
Declare Sub freeimage Lib "VIC32.DLL" (image As imgdes)
Declare Function convert1bitto8bit Lib "VIC32.DLL" (srcimg As imgdes, desimg As imgdes) As Long
Declare Sub copyimgdes Lib "VIC32.DLL" (srcimg As imgdes, desimg As imgdes)
Declare Function savejpg Lib "VIC32.DLL" (ByVal Fname As String, srcimg As imgdes, ByVal Quality As Long) As Long
    'end declarations
    'the sub

Public Sub BMPtoJPG(Thebmp As String, Thejpg As String, Quality As Long)
    Dim tmpimage As imgdes ' Image descriptors
    Dim tmp2image As imgdes
    Dim rcode As Long
    Dim vbitcount As Long
    Dim bdat As BITMAPINFOHEADER ' Reserve space For BMP struct
    Dim bmp_fname As String
    Dim jpg_fname As String
    bmp_fname = Thebmp
    jpg_fname = Thejpg
    ' Get info on the file we're to load
    rcode = bmpinfo(bmp_fname, bdat)


    If (rcode <> NO_ERROR) Then
        'cannot find file!
        Exit Sub
    End If
    vbitcount = bdat.biBitCount


    If (vbitcount >= 16) Then ' 16-, 24-, or 32-bit image is loaded into 24-bit buffer
        vbitcount = 24
    End If
    ' Allocate space for an image
    rcode = allocimage(tmpimage, bdat.biWidth, bdat.biHeight, vbitcount)


    If (rcode <> NO_ERROR) Then
        'not enuf memory!
        Exit Sub
    End If
    ' Load image
    rcode = loadbmp(bmp_fname, tmpimage)


    If (rcode <> NO_ERROR) Then
        freeimage tmpimage ' Free image On Error
        'cannot load file
        Exit Sub
    End If


    If (vbitcount = 1) Then ' If we loaded a 1-bit image, convert To 8-bit grayscale
        ' because jpeg only supports 8-bit grays
        '     cale or 24-bit color images
        rcode = allocimage(tmp2image, bdat.biWidth, bdat.biHeight, 8)


        If (rcode = NO_ERROR) Then
            rcode = convert1bitto8bit(tmpimage, tmp2image)
            freeimage tmpimage ' Replace 1-bit image With grayscale image
            copyimgdes tmp2image, tmpimage
        End If
    End If
    ' Save image
    rcode = savejpg(jpg_fname, tmpimage, Quality)
    freeimage tmpimage
End Sub


Public Function GetINI(Key As String) As String
Dim Ret As String, NC As Long
  
  Ret = String(600, 0)
  NC = GetPrivateProfileString("P2PWebcam", Key, Key, Ret, 600, App.Path & "\Config.ini")
  If NC <> 0 Then Ret = Left$(Ret, NC)
  If Ret = Key Or Len(Ret) = 600 Then Ret = ""
  GetINI = Ret

End Function
'Read from INI

Public Sub WriteINI(ByVal Key As String, Value As String)
  
  WritePrivateProfileString "P2PWebcam", Key, Value, App.Path & "\Config.ini"

End Sub
'Write to INI


Public Function OpenURL(ByVal sUrl As String) As String
On Error GoTo Err
    Dim hOpen               As Long
    Dim hOpenUrl            As Long
    Dim bDoLoop             As Boolean
    Dim bRet                As Boolean
    Dim sReadBuffer         As String * 2048
    Dim lNumberOfBytesRead  As Long
    Dim sBuffer             As String

DoEvents
hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PRECONFIG, _
    vbNullString, vbNullString, 0)

DoEvents
hOpenUrl = InternetOpenUrl(hOpen, sUrl, vbNullString, 0, _
   INTERNET_FLAG_RELOAD, 0)

    bDoLoop = True
    While bDoLoop
        DoEvents
        sReadBuffer = vbNullString
        bRet = InternetReadFile(hOpenUrl, sReadBuffer, _
           Len(sReadBuffer), lNumberOfBytesRead)
        sBuffer = sBuffer & Left$(sReadBuffer, _
             lNumberOfBytesRead)
        If Not CBool(lNumberOfBytesRead) Then bDoLoop = False
    Wend

    If hOpenUrl <> 0 Then InternetCloseHandle (hOpenUrl)
    If hOpen <> 0 Then InternetCloseHandle (hOpen)
    OpenURL = sBuffer
Exit Function
Err:
 OpenURL = "Couldn't Retrieve"
End Function
'opens an internet url. WAY better alternative than inet. will grab full url data and what not.

Public Function RandomGen2(rChars As String, rCount As Integer) As String
On Error Resume Next
  Dim tmpStr As String, x As Integer
    Randomize
      Do Until Len(tmpStr) = rCount
        x = Len(rChars) * Rnd + 1
        tmpStr = tmpStr & (Mid$(rChars, x, 1))
      Loop
        RandomGen2 = tmpStr
End Function

Public Sub Pause(ByVal interval As String)
On Error Resume Next
Dim wait   As Single
  
  wait = Timer
  
  Do While Timer - wait < CSng(interval$)
     DoEvents
 Loop
End Sub

  
