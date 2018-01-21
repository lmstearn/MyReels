Attribute VB_Name = "auxrout1"
Option Explicit
Global gwsMainWS     As Workspace    'main workspace object
Global gdbCurrentDB  As Database    'main database object
Private gsDBName         As String   'current database name
Private gsDataType       As String   'data backend = connect string

Public pct As Long, intscatternumber As Long, intscattervec(2, 2) As Long, intscattertotal As Long
Public VOGchg As Long, prepspin As Boolean, currheld As Long, procend As Boolean, chtr As Long, zhiddnstatus As Long, baseno As Long, waitimmarker As Long, BmpIX As Long, justcached As Boolean
Public spinz(4) As Long, picnum(4, 3) As Long, hreel(4) As Boolean, symbolselect As Long, dirofspin As Long, dirspin(4) As Long, substitute(14, 14) As Boolean
Public gamespinsymbol(3) As Long, disablegamespintabs(1) As Boolean, spinsettings(1, 15) As Long, reelcheck(14, 5) As Boolean
'reelcheck(pct,0) will be true if scatter is chosen through cornfig, false otherwise
Public gt(200) As Long, initvec(14) As Long, freegamesettings(1, 9) As Long, sst(14, 10) As Long, gamespinkeep(3) As Long
Public Stringvars(100) As String, hq(2, 6) As Long, quotestring(3) As String, loaddirectory As String, olddirectory As String, textwidthratio As Single, decsep As String
Public genoptsgen As Boolean, resX As Single, resY As Single, RR As Boolean, resCheck As Boolean, runAdmin As Boolean, isMainActive As Boolean, hfile As Integer, changecalculated As Boolean, midisNum As Long, midiDevTot As Integer
Public dbsCurrent As Database, sDatabaseName As String, rectemp As Recordset
Private ct As Long, ct1 As Long, ct2 As Long, temp As Long, response As Long, thumbsize As Long
Public Const INVALID_HANDLE_VALUE = -1
Public Const LOGPIXELSX = 88
Public Const LOGPIXELSY = 90
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_TRANSPARENT = &H20&
Public Const Alias = "tune1"
Const RASTERCAPS As Long = 38
Const DEFAULT_PITCH = 0
Const FF_DONTCARE = 0    'Don't care or don't know.
Const PROOF_QUALITY = 2
Const CLIP_DEFAULT_PRECIS = 0
Const CLIP_TT_ALWAYS = 32
Const OUT_DEFAULT_PRECIS = 0
Const OUT_TT_ONLY_PRECIS = 7
Const OUT_TT_PRECIS = 4
Const LF_FACESIZE = 32
Const OEM_CHARSET = 255
Const ANSI_CHARSET = 0


Const cash1 = "(This is my hard earned cash, you know)"
Const cash2 = "(Its probably worth the extra investment)"
Const cash3 = "(What else is my money good for?)"
Const return1 = "Maximum expected % return (hit me with a +6 studded flail)"
Const return2 = "Maximum expected % return (nudge me with a nylon broom)"
Const return3 = "Maximum expected % return (flick me with a matchstick)"

Const LOCALE_USER_DEFAULT& = &H400
Const LOCALE_SDECIMAL& = &HE


Type FileTime
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type LOGFONT_TYPE
          lfHeight As Long
          lfWidth As Long
          lfEscapement As Long
          lfOrientation As Long
          lfWeight As Long
          lfItalic As Byte
          lfUnderline As Byte
          lfStrikeOut As Byte
          lfCharSet As Byte
          lfOutPrecision As Byte
          lfClipPrecision As Byte
          lfQuality As Byte
          lfPitchAndFamily As Byte
          lffacename As String * LF_FACESIZE
    End Type

Private Type GUID
Data1 As Long
Data2 As Integer
Data3 As Integer
Data4(7) As Byte
End Type


Private Type PicBmp
Size As Long
Type As Long
hBmp As Long
hPal As Long
Reserved As Long
End Type


Public Type BrowseInfo
         hWndOwner      As Long
         pIDLRoot       As Long
         pszDisplayName As String
         lpszTitle      As String
         ulFlags        As Long
         lpfnCallback   As Long
         lParam         As Long
         iImage         As Long
      End Type


'Browse constants
Const BFFM_INITIALIZED = 1
Public Const WM_USER = &H400
Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)

Const MaxLFNPath = 260

Private Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FileTime
        ftLastAccessTime As FileTime
        ftLastWriteTime As FileTime
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MaxLFNPath
        cShortFileName As String * 14
End Type


'Free Space
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long 'C Bool
Private Declare Function GetDiskFreeSpaceExAsCurrency Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpDirectoryName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long       'C Bool

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" _
                        (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
' Rtns True (non zero) on success, False on failure
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" _
                        (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
' Rtns True (non zero) on success, False on failure
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

'Browse APIs
Private Declare Function lstrcpyA Lib "kernel32" (lpString1 As Any, lpString2 As Any) As Long
Private Declare Function lstrlenA Lib "kernel32" (lpString As Any) As Long


Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetDeviceCaps Lib "GDI32" (ByVal hDC As Long, ByVal nIndex As Long) As Long

'used for file browse
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function CreateFontIndirect Lib "GDI32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT_TYPE) As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

Private Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function BitBlt Lib "GDI32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long


'Winhelp API
Declare Function Winhelp Lib "user32" Alias "WinHelpA" (ByVal hWnd As Long, ByVal IpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
'Descimal separator
Private Declare Function GetLocaleInfo& Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long)
Private Sub DecSepar()
    Dim rtmp As Long, stmp As String
    stmp = String(10, "a")
    rtmp = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDECIMAL, stmp, 10)
    decsep = Left$(stmp, rtmp - 1) 'removing Null terminator
End Sub
Private Function getdrvname(drvdir As String)
Dim loctmp As Long, drvname As String
drvname = drvdir

loctmp = InStr(drvname, "\")


If loctmp > 0 Then
drvname = Left$(drvname, loctmp)
Else
drvname = drvname & "\"
End If

loctmp = InStr(drvname, " ")

'Is there a space ? Strip off the vol name if so
If loctmp > 0 Then drvname = Left$(drvname, loctmp - 1)

getdrvname = drvname


End Function

'* FUNCTION vbGetAvailableKBytesAsString()
'* ===============
'* This returns a VB string containing the number of '* free kilobytes on the drive 'pointed to by sPath.  This function will '* correctly call either GetDiskFreeSpace() or 'GetDiskFreeSpaceEx() '* from Win32 API as appropriate
'*
'* INPUTS:
'*  (Optional) sPath (see notes in vbGetAvailableBytesAsString function) '*
'* RETURNS:
'*  on error returns vbNullChar ("")
'*  else returns number of free (available) kilobytes as string '*  (rounded to nearest 'kbyte)
'****************************************************************
Public Function vbGetAvailableKBytesAsString(Optional ByVal sPath As String = "") As String
Dim bytes As Currency, kBytes As Currency, stmp As String
stmp = getdrvname(sPath)

stmp = vbGetAvailableBytesAsString(stmp)
bytes = CCur(stmp)
' If commented out by JRK - gives empty string instead of "0"     ' If bytes Then 'avoid divide by 0 errors
kBytes = bytes / 1024
kBytes = Fix(kBytes)
vbGetAvailableKBytesAsString = CStr(kBytes)
' End If
End Function
'* FUNCTION vbGetAvailableBytesAsString()
'* ===============
'* This routine will return a VB string containing the number of '* free bytes on the drive 'pointed to by sPath.  This function will '* correctly call either GetDiskFreeSpace() or 'GetDiskFreeSpaceEx() '* from the Win32 API as appropriate
'*
'* INPUTS:
'*  (Optional) sPath
'* -notes from MSDN documentation-
'* Pointer to a null-terminated string that specifies the root directory '* of the disk to return 'information about. If lpRootPathName is omitted, '* the function uses the root of the current 'directory. If this parameter '* is a UNC name, you must follow it with an additional backslash. 'For '* example, you would specify \\MyServer\MyShare as \\MyServer\MyShare\. '* Windows 95: The 'initial release of Windows 95 does not support UNC paths '* for the lpszRootPathName parameter. 'To query the free disk space using a '* UNC path, temporarily map the UNC path to a drive 'letter, 'query the free '* disk space on the drive, then remove the temporary mapping. Windows 95 '* OSR2 'and later: UNC paths are supported.
'*
'* RETURNS:
'*  on error returns vbNullChar ("")
'*  otherwise returns number of free (available) bytes as string '****************************************************************
Private Function vbGetAvailableBytesAsString(Optional ByVal sPath As String = "") As String
Dim lo As Long, hi As Long
Dim sOut As String
If ExistGetDiskFreeSpaceEx() Then
sOut = vbGetAvailableBytesEx(sPath)
Else
sOut = CStr(vbGetAvailableBytes(sPath))
End If
vbGetAvailableBytesAsString = sOut
End Function
'* FUNCTION ExistGetDiskFreeSpaceEx()
'* ===============
'* This routine used the Microsoft-recommended way to determine if '* the Win32 API function *GetDiskFreeSpaceEx() exists on the current '* OS platform. (should be available on all Win32 *systems after OSr.2) '*
'* INPUTS: none
'*
'* RETURNS:
'*  TRUE - if the GetDiskFreeSpaceEx() function is available
'*  FALSE - if the GetDiskFreeSpaceEx() function is available
'*          in this case you should call the older GetDiskFreeSpace() '****************************************************************
Private Function ExistGetDiskFreeSpaceEx() As Boolean
Dim hInst As Long
Dim procAddress As Long
hInst = LoadLibrary("kernel32.dll")
If hInst Then
procAddress = GetProcAddress(hInst, "GetDiskFreeSpaceExA")
Call FreeLibrary(hInst)
End If
ExistGetDiskFreeSpaceEx = CBool(procAddress)
End Function
'* FUNCTION vbGetAvailableBytes()
'* ===============
'* This routine will return the number of free bytes on the
'* specified drive (does not handle drive partitions over
'* 2GB)
'*
'* INPUTS:
'* sPath - (see notes in vbGetAvailableBytesAsString function)
'*
'* RETURNS:
'* Long - free disk space in bytes
'****************************************************************
Private Function vbGetAvailableBytes(ByVal sPath As String) As Long
Dim lSpc As Long 'sectors per cluster
Dim lBps As Long 'bytes per sector
Dim lNfc As Long 'number of free clusters
Dim lTnc As Long 'total number of clusters
Call GetDiskFreeSpace(sPath, lSpc, lBps, lNfc, lTnc)
vbGetAvailableBytes = lSpc * lBps * lNfc
End Function
'****************************************************************
'* FUNCTION vbGetAvailableBytesEx()
'* ===============
'* This routine will return a String containing the the available
'* bytes as reported by the GetDiskFreeSpaceEX() API
'*
'* This function will correctly return values for large disk partitions (i.e., Fat32)
'*
'* INPUTS:
'* sPath - (see notes in vbGetAvailableBytesAsString function)
'*
'* RETURNS:
'*
'* String - Available bytes on disk pointed to by sPath
'****************************************************************
Private Function vbGetAvailableBytesEx(ByVal sPath As String) As String
    Dim BytesAvailable As Currency
    Dim TotalBytes As Currency
    Dim TotalFreeBytes As Currency
    Dim tmp As Currency

    On Error GoTo APIfailed
    If "" = sPath Then
        Call GetDiskFreeSpaceExAsCurrency(vbNullString, BytesAvailable, TotalBytes, TotalFreeBytes)
    Else
        Call GetDiskFreeSpaceExAsCurrency(sPath, BytesAvailable, TotalBytes, TotalFreeBytes)
    End If

    'If BytesAvailable Then
        BytesAvailable = BytesAvailable * 10000
        vbGetAvailableBytesEx = CStr(BytesAvailable)
    'End If
    Exit Function
APIfailed:
    'returns false
    Debug.Print "GetDiskFreeSpaceEx() API Failed!"
End Function
Private Function vbGetTotalBytesEx(ByVal sPath As String) As String
Dim BytesAvailable As Currency
Dim TotalBytes As Currency
Dim TotalFreeBytes As Currency
On Error GoTo APIfailed
If "" = sPath Then
Call GetDiskFreeSpaceExAsCurrency(vbNullString, BytesAvailable, TotalBytes, TotalFreeBytes)
Else
Call GetDiskFreeSpaceExAsCurrency(sPath, BytesAvailable, TotalBytes, TotalFreeBytes)
End If
'If TotalBytes Then
TotalBytes = TotalBytes * 10000
vbGetTotalBytesEx = CStr(TotalBytes)
'End If
Exit Function
APIfailed:
'returns false
MsgBox "GetDiskFreeSpaceEx Failed!", vbOKOnly
End Function
'****************************************************************
Public Function findafile(curpath$, filecriteria As String) As Single
'Findafile returns a non zero size of file it has found
Dim WFD As WIN32_FIND_DATA, xfile As String
Const KB& = 1024
findafile = 0
With WFD
If Right$(curpath$, 1) <> "\" Then curpath$ = curpath$ & "\"
    
xfile = FindFirstFile(curpath$ & filecriteria, WFD)
If xfile <> INVALID_HANDLE_VALUE Then
Do
If Len(Left$(.cFileName, InStr(.cFileName, vbNullChar) - 1)) > 0 Then
    If .nFileSizeHigh > 0 Then
    MsgBox "File too big !"
    findafile = 0
    Exit Function
    Else
    Select Case .nFileSizeLow \ KB
    Case 0
    Case Is < 10
    findafile = CSng(Format(.nFileSizeLow / KB, "0.00"))
    Case Is < 100
    findafile = CSng(Format(.nFileSizeLow / KB, "0.0"))
    Case Else
    findafile = CSng(Format(.nFileSizeLow / KB, "0"))
    End Select
    End If
End If
' Get the next file matching the FileSpec$,loadnow is true if Slotdata.s$t exists
Loop While FindNextFile(xfile, WFD)
Call FindClose(xfile)
End If
End With
End Function
Public Function BrowseCallbackProcStr(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
                                       
'Callback for the Browse STRING method.
 
'On initialization, set the dialog's
'pre-selected folder from the pointer
'to the path allocated as bi.lParam,
'passed back to the callback as lpData param.
 
   Select Case uMsg
      Case BFFM_INITIALIZED
      
         Call SendMessage(hWnd, BFFM_SETSELECTIONA, True, ByVal StrFromPtrA(lpData))
                          
         Case Else:
         
   End Select
          
End Function
Public Function FARPROC(pfn As Long) As Long
'A dummy procedure that receives and returns
'the value of the AddressOf operator.
'Obtain and set the address of the callback
'This workaround is needed as you can't assign
'AddressOf directly to a member of a user-
'defined type, but you can assign it to another
'long and use that (as returned here)
 
FARPROC = pfn

End Function
Private Function StrFromPtrA(lpszA As Long) As String

'Returns an ANSI string from a pointer to an ANSI string.
   
Dim sRtn As String
sRtn = String$(lstrlenA(ByVal lpszA), 0)
Call lstrcpyA(ByVal sRtn, ByVal lpszA)
StrFromPtrA = sRtn

End Function
Public Function randwheelvec(thumbsize As Long, wheelvec() As Long, wheelorder() As Long)
Dim order100(100) As Long, neworder24(24) As Long, tempwheelorder(4, 24) As Long
Dim boolescape As Boolean, testval As Long, firsttotal As Long, secondtotal As Long, intreel As Long
'The next is not strictly a randwheelvec routine
'but has to be done somewhere at this stage

randwheelvec = True

zeroscatter ' zero scatter vars

For pct = 1 To thumbsize

'Initialize intscatternumber and intscattervec while we are here
Scatterinit wheelvec(1, pct), pct, True

testval = 0
For intreel = 1 To 5
If wheelvec(intreel, pct) > 0 Then testval = testval + 1
Next
sst(pct, 0) = 0

        If testval < 5 Then

        'zero current prefs
        For ct = 1 To 5
        sst(pct, ct) = 0
        Next

        'reset any old spin/game settings for pct, resolves old lt 5 problems
        For ct = 0 To 3

        If gamespinsymbol(ct) = pct Then
        Select Case ct
        Case 0
        For ct1 = 1 To 9
        freegamesettings(0, ct1) = 0
        Next
        Case 1
        For ct1 = 1 To 9
        freegamesettings(1, ct1) = 0
        Next
        Case 2
        For ct1 = 1 To 15
        spinsettings(0, ct1) = 0
        Next
        Case 3
        For ct1 = 1 To 15
        spinsettings(1, ct1) = 0
        Next
        End Select
        gamespinsymbol(ct) = 0
        End If

        Next
        End If

'Testval = 5, if sstab(0) > 0 reset to 0
If testval = 4 Then
    sst(pct, 6) = 0    'have to zero top prize
    If wheelvec(5, pct) = 0 Then
    sst(pct, 0) = 13
    sst(pct, 1) = 1
    ElseIf wheelvec(1, pct) = 0 Then
    sst(pct, 0) = 14
    sst(pct, 3) = 1
    sst(pct, 1) = 1
    ElseIf wheelvec(3, pct) = 0 Then
    sst(pct, 0) = 15
    sst(pct, 5) = 1
    ElseIf wheelvec(4, pct) = 0 Then
    sst(pct, 0) = 16
    sst(pct, 5) = 1
    Else
    sst(pct, 0) = 17
    sst(pct, 5) = 1
    End If
ElseIf testval = 3 Then
    sst(pct, 6) = 0
    sst(pct, 7) = 0
    If wheelvec(4, pct) = 0 And wheelvec(5, pct) = 0 Then
    sst(pct, 0) = 6
    sst(pct, 1) = 1
    ElseIf wheelvec(1, pct) = 0 And wheelvec(2, pct) = 0 Then
    sst(pct, 0) = 7
    sst(pct, 3) = 1
    sst(pct, 1) = 1
    ElseIf wheelvec(1, pct) = 0 And wheelvec(5, pct) = 0 Then
    sst(pct, 0) = 8
    sst(pct, 4) = 1
    sst(pct, 1) = 1
    Else
        If wheelvec(1, pct) > 0 And wheelvec(2, pct) > 0 Then
        sst(pct, 0) = 9
        ElseIf wheelvec(4, pct) > 0 And wheelvec(5, pct) > 0 Then
        sst(pct, 0) = 10
        ElseIf wheelvec(2, pct) > 0 And wheelvec(3, pct) > 0 Then
        sst(pct, 0) = 11
        ElseIf wheelvec(3, pct) > 0 And wheelvec(4, pct) > 0 Then
        sst(pct, 0) = 11
        Else
        sst(pct, 0) = 12
        End If
    sst(pct, 5) = 1
    End If
ElseIf testval = 2 Then
    For ct = 6 To 8
    sst(pct, ct) = 0
    Next
    If wheelvec(1, pct) > 0 And wheelvec(2, pct) > 0 Then
    sst(pct, 0) = 2
    sst(pct, 1) = 1
    ElseIf wheelvec(4, pct) > 0 And wheelvec(5, pct) > 0 Then
    sst(pct, 0) = 3
    sst(pct, 3) = 1
    sst(pct, 1) = 1
    ElseIf wheelvec(2, pct) > 0 And wheelvec(3, pct) > 0 Then
    sst(pct, 0) = 4
    sst(pct, 5) = 1
    ElseIf wheelvec(3, pct) > 0 And wheelvec(4, pct) > 0 Then
    sst(pct, 0) = 4
    sst(pct, 5) = 1
    Else
    sst(pct, 0) = 5
    sst(pct, 5) = 1
    End If
ElseIf testval = 1 Then
    For ct = 6 To 9
    sst(pct, ct) = 0
    Next
    sst(pct, 0) = 1
    sst(pct, 5) = 1
End If
Next

'End thumbsize loop, now loop through reels
For intreel = 0 To 4

Starthere:
For ct = 1 To 100
order100(ct) = Int((24 - intscattertotal) * Rnd)
Next

ct1 = 1
neworder24(1) = order100(1)
For ct = 2 To 100
boolescape = False
testval = order100(ct)
    For ct2 = 1 To ct1
    If testval = neworder24(ct2) Then boolescape = True
    Next
    If boolescape = False Then
    ct1 = ct1 + 1
    neworder24(ct1) = testval
    If ct1 = 24 - intscattertotal Then Exit For
    End If
Next
If ct1 < 24 - intscattertotal Then GoTo Starthere

'Now populate tempwheelorder
temp = 0
testval = 0
'Summing through all non - scatters
For pct = 1 To thumbsize
    If sst(pct, 2) = 0 Then 'no scatters here
    testval = wheelvec(intreel + 1, pct)
    For ct = temp To temp + testval
    tempwheelorder(intreel, ct) = pct
    Next
    temp = testval + temp

    End If 'scatter condition
Next


For ct = 1 To 24 - intscattertotal
wheelorder(intreel, ct) = tempwheelorder(intreel, neworder24(ct))
Next

If intscatternumber > 0 Then    'now for scatters

    'Scramble tempwheelorder before distributing scatters
    For ct = 1 To 24 - intscattertotal
    tempwheelorder(intreel, ct) = wheelorder(intreel, ct)
    Next

    'First generate new randoms (order100 re-used) to distribute randoms evenly
    firsttotal = intscattervec(1, 1)
    secondtotal = intscattervec(2, 1)

    For ct = 1 To firsttotal

    If intscatternumber = 1 Then    'only 1 scatter
        If intscattertotal = 4 Then
        order100(ct) = Int(3 * Rnd + 4)
        Else
        order100(ct) = Int(2 * Rnd + 7)
        End If

    Else    '2 scatters
    order100(ct) = Int(2 * Rnd + 4)
    End If

    Next

    If intscatternumber = 2 Then    '2 scatters
    For ct = firsttotal + 1 To intscattertotal
    If firsttotal = 4 Then
        If secondtotal = 4 Then
        order100(ct) = Int(2 * Rnd + 1)
        Else
        order100(ct) = order100(2 * (ct - firsttotal)) + Int(2 * Rnd + 1)
        End If
    Else    'intscattervec(1, 1) <= 2
    order100(ct) = order100(ct - 1) + Int(2 * Rnd + 1)
    End If

    Next
    End If


    ct1 = 0
    ct2 = 0
    ct = 1
    
    'Work the last segment of neworder
    While ct + ct2 <= intscattertotal
    ct1 = ct1 + order100(ct)
    neworder24(24 - intscattertotal + ct) = ct1
    ct = ct + 1
    If ct2 < secondtotal Then
    ct2 = ct2 + 1
    neworder24(24 - secondtotal + ct2) = ct1 + order100(firsttotal + ct2)
    End If
    Wend

    For ct = 24 - intscattertotal + 1 To 24
    'Allocate the first scatter symbol Id to tempwheelorder in the swap
    temp = tempwheelorder(intreel, neworder24(ct))
    If ct <= 24 - secondtotal Then
    tempwheelorder(intreel, neworder24(ct)) = intscattervec(1, 2)
    Else
    tempwheelorder(intreel, neworder24(ct)) = intscattervec(2, 2)
    End If
    tempwheelorder(intreel, ct) = temp
    Next
    
    For ct = 1 To 24
    wheelorder(intreel, ct) = tempwheelorder(intreel, ct)
    If wheelorder(intreel, ct) = 0 Then
    randwheelvec = False
    Exit Function
    End If
    Next

End If 'Intscatter conditon

ct2 = 0

'Prevent doubles
For ct1 = 1 To 24
If wheelorder(intreel, ct1) = wheelorder(intreel, Advanz(ct1, -1)) Or wheelorder(intreel, ct1) = wheelorder(intreel, Advanz(ct1, 1)) Then
ct = ct2
temp = 0
'Don't want last value

Do While True
ct2 = Int(24 * Rnd)
temp = temp + 1
    If ct2 <> ct Then
        If ct1 = ct2 Then
        ElseIf wheelorder(intreel, ct1) = intscattervec(1, 2) Or wheelorder(intreel, ct1) = intscattervec(2, 2) Then
        ElseIf wheelorder(intreel, ct2) = intscattervec(1, 2) Or wheelorder(intreel, ct2) = intscattervec(2, 2) Then
        ElseIf wheelorder(intreel, ct2) = wheelorder(intreel, Advanz(ct1, -1)) Or wheelorder(intreel, ct2) = wheelorder(intreel, Advanz(ct1, 1)) Then
        ElseIf wheelorder(intreel, ct1) = wheelorder(intreel, Advanz(ct2, -1)) Or wheelorder(intreel, ct1) = wheelorder(intreel, Advanz(ct2, 1)) Then
        Else
        Exit Do
        End If
    End If
If temp > 24 Then Exit Do
Loop

If temp < 25 Then
temp = wheelorder(intreel, ct1)
wheelorder(intreel, ct1) = wheelorder(intreel, ct2)
wheelorder(intreel, ct2) = temp
End If
End If

Next

'End intreel loop
Next

End Function
Public Sub initwheelvec()
'set baseno
Select Case thumbsize
Case 3 To 8
baseno = 12 - thumbsize
Case 9 To 14
baseno = Int(53 / thumbsize) - 1
End Select
Select Case thumbsize
Case 3 To 7
initvec(thumbsize) = 13 - thumbsize
randominitvec thumbsize - 1, 24 - 13 + thumbsize
Case 8
initvec(8) = 1
initvec(7) = 5
randominitvec 6, 18, 7
Case 9
initvec(7) = 1
initvec(8) = 1
initvec(9) = 5
randominitvec 6, 17, 7
Case 10
initvec(6) = 2
initvec(7) = 1
initvec(8) = 1
initvec(9) = 1
initvec(10) = 5
randominitvec 5, 14, 6
Case 11
initvec(10) = 1
initvec(11) = 4
randominitvec 9, 19, 10
Case 12
initvec(10) = 1
initvec(11) = 1
initvec(12) = 4
randominitvec 9, 18, 10
Case 13
initvec(10) = 1
initvec(11) = 1
initvec(12) = 1
initvec(13) = 4
randominitvec 9, 17, 10
Case 14
initvec(10) = 1
initvec(11) = 1
initvec(12) = 1
initvec(13) = 1
initvec(14) = 3
randominitvec 9, 17, 10
End Select
If gt(1) = 1 Then ShellSort initvec, thumbsize
End Sub
Private Sub randominitvec(stepno As Long, limmit As Long, Optional fixedbit As Long = 0)
Dim randvec(100) As Long
sstart:
For ct1 = 1 To 100
randvec(ct1) = Int(baseno * Rnd + 1)
Next
For ct1 = 2 To stepno
initvec(ct1) = randvec(ct1 - 1)
Next
temp = stepno - 1
For ct1 = 1 To 100 - stepno
ct = 0
'cycle the values
For ct2 = 1 To temp
initvec(ct2) = initvec(ct2 + 1)
ct = ct + initvec(ct2)
Next
initvec(stepno) = randvec(ct1 + temp)
If (ct + initvec(stepno) = limmit) Then GoTo endd
Next
GoTo sstart
endd:
'now mix remaining nos
If fixedbit > 0 Then
For ct = fixedbit To thumbsize
ct1 = initvec(ct)
temp = CLng(CInt(fixedbit * Rnd + 1))
initvec(ct) = initvec(temp)
initvec(temp) = ct1
Next
End If
End Sub
Public Sub Scatterinit(scatfreq As Long, pics As Long, pleaseswap As Boolean)
'Intscatternumber = number of scatters
'intscattervec(intscatternumber, piccounter)

If sst(pics, 2) = 1 Then
intscatternumber = intscatternumber + 1
intscattervec(intscatternumber, 1) = scatfreq
intscattervec(intscatternumber, 2) = pics
Select Case intscatternumber
Case 1
intscattertotal = intscattervec(1, 1)
Case 2
    If intscattervec(2, 1) > intscattervec(1, 1) And pleaseswap = True Then 'swap if necessary
    temp = intscattervec(1, 1)
    intscattervec(1, 1) = intscattervec(2, 1)
    intscattervec(2, 1) = temp
    temp = intscattervec(1, 2)
    intscattervec(1, 2) = intscattervec(2, 2)
    intscattervec(2, 2) = temp
    End If
intscattertotal = intscattervec(1, 1) + intscattervec(2, 1)

End Select
End If
End Sub
Public Sub chgchtr(Optional makepoz As Boolean = False)
If makepoz = True Then
Select Case chtr
Case Is = 0, 1
chtr = thumbsize / 2
Case Is < 0
chtr = -chtr
Case Else
chtr = chtr - 1
End Select
Else    'toggle for next time
If chtr > 0 Then
chtr = -chtr
Else
Exit Sub
End If
Select Case chtr
Case -1
chtr = -thumbsize / 2
Case Else
chtr = chtr + 1
End Select
End If
End Sub
Public Sub setinauxrouts(piccount As Long)
'This sub gets thumbsize
thumbsize = piccount
End Sub
Public Sub resetprize(rlz As Long, symbl As Long)
'Uses fact of maxno of substitutes = 2 (later comment max is 4 in SSTAB)
temp = -1
If symbl < thumbsize Then   'take prize candidates of lower rank
    For ct1 = symbl + 1 To thumbsize
        If sortprizes(rlz, ct1, temp) = True Then
        temp = ct1
        Exit For
        End If
    Next
End If
If temp = -1 And symbl > 1 Then   'if above fails take prize candidates of higher rank
    For ct1 = symbl - 1 To 1 Step -1
        If sortprizes(rlz, ct1, temp) = True Then
        temp = ct1
        Exit For
        End If
    Next
End If

If temp = -1 Then
'Last resort
sst(symbl, rlz) = Int(10 ^ (9 - rlz))
ElseIf rlz < 9 And sst(temp, rlz) = 0 Then 'Everything is lt 5
sst(symbl, rlz) = Int(10 ^ (9 - rlz))
Else
sst(symbl, rlz) = sst(temp, rlz)
End If
End Sub
Public Function chkprizelevel(goingup As Boolean, prizelevel As Long)
Dim testval As Long
'symbolselect is Public
If sst(symbolselect, 2) = 1 Then   'Not dealing with scatters
chkprizelevel = True
Exit Function
End If
temp = sst(symbolselect, prizelevel)
testval = -1
chkprizelevel = False

If goingup = True Then
'First need to compare within symbols own prize levels

If prizelevel > 6 Then  'not needed at top level

For ct = prizelevel - 1 To 6 Step -1
        If sortprizes(ct, symbolselect, testval) = True Then
                If temp >= sst(symbolselect, ct) Then
                Exit Function
                Else
                Exit For
                End If
        End If
Next
End If

        If symbolselect > 1 Then
        testval = -1
        For ct = symbolselect - 1 To 1 Step -1
        If sortprizes(prizelevel, ct, testval) = True Then
            If temp >= sst(ct, prizelevel) Then
            Exit Function
            Else
            chkprizelevel = True
            Exit Function
            End If
        End If
        Next
        End If
        'gets here if all higher ranking symbols < prizelevel are 0 and or scatters (testval -1) or Symbolselect = 1, no limit
Else    'going down

    'First need to compare within symbols own prize levels

    If prizelevel < 10 Then 'not needed at top level

    For ct = prizelevel + 1 To 10   'sortprizes not necessary
        If temp <= sst(symbolselect, ct) Then
        Exit Function
        Else
        Exit For
        End If
    Next
    End If

    If symbolselect < thumbsize Then
        testval = -1
        For ct = symbolselect + 1 To thumbsize
        If sortprizes(prizelevel, ct, testval) = True Then
            If temp <= sst(ct, prizelevel) Then
            Exit Function
            Else
            chkprizelevel = True
            Exit Function
            End If
        End If
        Next
    End If
    'gets here if all lower ranking symbols < prizelevel are 0 and or scatters or Symbolselect = thumbsize, no limit going down
End If
chkprizelevel = True
End Function
Public Function sortprizes(rlz As Long, c As Long, testval As Long)
'Examines LT situations as well
If sst(c, 2) = 0 Then
        Select Case sst(c, 0)
        Case 0
        If rlz < 9 Then
        '(for new game cfg) must have non zero prize here
        If sst(c, rlz) > 0 Then testval = c
        Else
        testval = c
        End If
        Case 1
        If rlz = 10 Then testval = c
        Case 2 To 5
        If rlz > 8 Then testval = c
        Case 6 To 12
        If rlz > 7 Then testval = c
        Case Else
        If rlz > 6 Then testval = c
        End Select
End If
If testval = -1 Then    'no scatters or outside lt range
sortprizes = False
Else
sortprizes = True
End If
End Function
Public Function findhcmplusone(foo As Integer, multfoo As Integer)
Dim fooresult As Integer
fooresult = Int(foo / multfoo)
If fooresult = foo / multfoo Then
findhcmplusone = foo
Else
findhcmplusone = (fooresult + 1) * multfoo
End If
End Function
Public Function Advanz(objekt As Long, inkrement As Long)
objekt = objekt + inkrement
If objekt > 24 Then objekt = objekt - 24
If objekt < 1 Then objekt = 24 + objekt
Advanz = objekt
End Function
Public Function Cycle(j As Long, n As Long)
If dirofspin = 1 Then
If j < n Then
Cycle = n - j - 1
Else
Cycle = 3 + n - j
End If
Else
If j < n Then
Cycle = 4 - n + j
Else
Cycle = j - n
End If
End If
End Function
Public Function maxY(X As Long, Y As Long)
If X > Y Then
maxY = X
Else
maxY = Y
End If
End Function
Public Sub randomspinvec(pw As Long, zmedfastslowmove() As Long, special As Long, spinzstart As Long)
Dim quotRnd As Long, temperr As String, pwScale As Long
temperr = "Quotes.s$t corrupt or missing or not configured correctly, please reconfigure ..."


pwScale = fixpw(1, pw)



For ct = 0 To 24
ct1 = Int(2 * Rnd + 2)
zmedfastslowmove(1, ct) = pwScale / ct1
ct1 = Int(2 * Rnd + 1)
zmedfastslowmove(2, ct) = pwScale / ct1
ct1 = Int(2 * Rnd + 3)
zmedfastslowmove(3, ct) = pwScale / ct1
ct1 = Int(2 * Rnd + 4)
zmedfastslowmove(4, ct) = pwScale / ct1
Next

For ct = 0 To 4
ct1 = Int(2 * Rnd + 4)
zmedfastslowmove(0, ct) = ct1
Next

spinzstart = zmedfastslowmove(0, special + Int(3 * Rnd) + 1) 'seed spinzstart
'total spinz on the total for this spin


'Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
'The 0'th,0'th element is the random number for the midi playlist
midisNum = 0
For ct = 44 To 50
If Stringvars(ct) <> "" Then midisNum = midisNum + 1
Next



If gt(187) < 0 Then

'keep randomness
ct = Int(Rnd)   'Create dummy random gen as to change random process
zmedfastslowmove(0, 5) = Int(midisNum * Rnd + 44)
Else
Select Case midisNum
Case 0
zmedfastslowmove(0, 5) = -1
Case Else
zmedfastslowmove(0, 5) = gt(187) + 43
End Select
ct = Int(Rnd)
End If

'Always clear before spinning
For ct = 0 To 3
quotestring(ct) = ""
Next



'The 0'th,1'th element generates random quote selections
If Stringvars(3) = "" Then
gt(191) = 0
quotRnd = Int(Rnd + 1)    'Create dummy random gen as to change random process
Else
    If gt(159) = 1 Then
    quotRnd = Int(Rnd + 1)
    Else
        'Get quotestring here
        sDatabaseName = Stringvars(3) & "Quotes.s$t"
        If Openquotes("dummy.txt", False, quotRnd) = True Then
        zmedfastslowmove(0, 6) = quotRnd
        Else
        MsgBox temperr, vbOKOnly
        Stringvars(3) = ""
        gt(191) = 0
        quotRnd = Int(Rnd + 1)
        End If
    End If
End If

End Sub
Public Sub LoadFrmSplsh(heightval As Long)
Load frmSplsh
With frmSplsh

If heightval <> 3000 Then
.lblconfigload.Visible = True
.Frame1.Visible = False
End If

Select Case heightval
Case 430
.lblconfigload.Caption = Space(10) & "Loading Thumbnail DB ......"
Case 440
.lblconfigload.Caption = Space(19) & "Loading Game ......"
Case 450
.lblconfigload.Caption = Space(20) & "Please Wait ......"
Case 460
Unload Zhidden
Set Zhidden = Nothing
End Select

.Height = heightval

.AutoRedraw = True
.Show
Sleep 250
.AutoRedraw = False
End With
End Sub
Public Sub firstgametypeload()

DecSepar 'Get decimal sepeartor
zhiddnstatus = 0

Select Case gt(158)
Case 0
'initialise colours as - 1
gt(161) = -1
gt(162) = -1
For ct = 1 To 190
gt(ct) = 0
Next
Case 1
For ct = 11 To 156
Select Case ct
Case 11, 13, 16 To 18, 26 To 34, 48 To 150, 154, 156
gt(ct) = 0
End Select
Next
Case 2
'initialise colours as - 1
gt(161) = -1
gt(162) = -1
For ct = 1 To 157
gt(ct) = 0
Next
For ct = 159 To 190
gt(ct) = 0
Next
End Select

'Leave gt(191) quotes total alone

For ct = 0 To 1
disablegamespintabs(ct) = False
For ct1 = 1 To 9
freegamesettings(ct, ct1) = 0
Next
For ct1 = 1 To 15
spinsettings(ct, ct1) = 0
Next
Next

For ct = 0 To 3
gamespinsymbol(ct) = 0
gamespinkeep(ct) = 0
Next

For ct = 1 To 14

sst(ct, 0) = 0
sst(ct, 2) = 0

reelcheck(ct, 0) = False
For ct1 = 1 To 5
reelcheck(ct, ct1) = True
Next

For ct1 = 1 To 14
substitute(ct, ct1) = False
Next
Next


End Sub
Public Sub zeroscatter()
intscattertotal = 0
intscatternumber = 0
For ct = 1 To 2
For temp = 1 To 2
intscattervec(ct, temp) = 0
Next
Next
End Sub
Public Function testscatter(a As Long, B As Long, c As Long, D As Long, e As Long)
'Screen scatter candidates
Select Case a
Case 1, 2, 4
If a = B And B = c And c = D And D = e Then
testscatter = True
Else
testscatter = False
End If
Case Else
testscatter = False
End Select
End Function
Public Function XdownI(X As Long, i As Long)
temp = 1
If i = 0 Or X = 0 Then
XdownI = 1
Else
For ct = 1 To i
temp = (X - ct + 1) * temp
Next
XdownI = temp
End If
End Function
Public Sub ShellSort(testvec() As Long, intlength As Long)
Dim nSwaps%, nComp%, iOffset%, iLimit%, iSwitch%, i%, a$
'Perform Shell Sort on testvec
    'Set comparison offset to half the number of items.
    nSwaps% = 0
    nComp% = 0
 
  iOffset% = (intlength) / 2

  While iOffset%
    ' Loop until offset gets to zero.
    iLimit% = intlength - iOffset% + 1 ' - 1
    Do
      iSwitch% = False
      For i% = 2 To iLimit%
               nComp% = nComp% + 1
        If testvec(i% - 1) > testvec(i% - 1 + iOffset%) Then
          a$ = testvec(i% - 1)
          testvec(i% - 1) = testvec(i% - 1 + iOffset%)
          testvec(i% - 1 + iOffset%) = a$
          nSwaps% = nSwaps% + 1
          iSwitch% = i%
        End If
             Next i%
      ' Sort on next pass only to where last switch was made.
      iLimit% = iSwitch% - iOffset%
    Loop While iSwitch%
    ' No switches at last offset, try one half as big.
    iOffset% = iOffset% / 2
  Wend
End Sub
Public Sub Reduceint(count1, inttracker)
For ct = 1 To 14
If inttracker(ct) > inttracker(count1) Then inttracker(ct) = inttracker(ct) - 1
Next
End Sub
Public Sub TextCircle(obj As Object, txt As String, X As Long, Y As Long, radius As Long, ByVal startdegree As Double, objfontsize As Integer)
Dim foo As Long, TxtX As Long, TxtY As Long
Dim twipsperdegree As Long, wrktxt As String, wrklet As String, degreexy As Double, degree As Double
twipsperdegree = (radius * 3.14159 * 2) / 360
    
 startdegree = startdegree + Int(360 - (((obj.TextWidth(txt)) / twipsperdegree) / 2))
    
For foo = 1 To Len(txt)
    wrklet = Mid$(txt, foo, 1)
    degreexy = (obj.TextWidth(wrktxt)) / twipsperdegree + startdegree
    DegreesToXY X, Y, degreexy, radius, radius, TxtX, TxtY
    degree = (obj.TextWidth(wrktxt) + 0.5 * obj.TextWidth(wrklet)) / twipsperdegree + startdegree
    RotateText 360 - degree, obj, Stringvars(9), CSng(objfontsize), (TxtX), (TxtY), wrklet
    wrktxt = wrktxt & wrklet
Next foo
End Sub
Public Sub DegreesToXY(CenterX As Long, CenterY As Long, degree As Double, radiusX As Long, radiusY As Long, X As Long, Y As Long)
Dim convert As Double

    convert = 3.141593 / 180
    X = CenterX - (Sin(-degree * convert) * radiusX)
    Y = CenterY - (Sin((90 + (degree)) * convert) * radiusY)

End Sub
Public Sub RotateText(Degrees As Long, obj As Object, fontname As String, Fontsize As Single, X As Long, Y As Long, Caption As String)
Dim RotateFont As LOGFONT_TYPE, hWnd As Long, hDC As Long
Dim CurFont As Long, rFont As Long, foo As Long, inc1 As Long, inc2 As Long

hWnd = GetDesktopWindow
hDC = GetDC(hWnd)


With RotateFont
      
      
' All but two properties are very straight-forward,
' even with rotation, and map directly.
      '
.lfHeight = -(Fontsize * GetDeviceCaps(hDC, LOGPIXELSY)) / 72
.lfWidth = 0
.lfEscapement = Degrees * 10
.lfOrientation = .lfEscapement
.lfClipPrecision = CLIP_TT_ALWAYS
.lfQuality = PROOF_QUALITY
.lfPitchAndFamily = DEFAULT_PITCH Or FF_DONTCARE
.lffacename = fontname & Chr$(0)


If .lfCharSet = OEM_CHARSET And Degrees <> 0 Then RotateFont.lfCharSet = ANSI_CHARSET

.lfOutPrecision = OUT_TT_PRECIS

' Only TrueType fonts can rotate, so we must

If obj.FontBold Then
    .lfWeight = 800
Else
    .lfWeight = 300
End If

End With

rFont = CreateFontIndirect(RotateFont)
CurFont = SelectObject(obj.hDC, rFont)

inc1 = X + (Fontsize * GetDeviceCaps(hDC, LOGPIXELSX)) / 72
inc2 = Y + (Fontsize * GetDeviceCaps(hDC, LOGPIXELSY)) / 72


obj.ForeColor = gt(176)
obj.CurrentX = 15 + inc1
obj.CurrentY = 15 + inc2
obj.Print Caption

'Echoing
obj.ForeColor = gt(164)
obj.CurrentX = inc1
obj.CurrentY = inc2
obj.Print Caption
'Restore
foo = SelectObject(obj.hDC, CurFont)
foo = DeleteObject(rFont)
Call ReleaseDC(hWnd, hDC)

End Sub
Public Sub Randomisethem(seedchanged As Boolean)
seedchanged = False

Select Case gt(36)
Case 0
If gt(38) <> Second(Time) Then
gt(38) = Second(Time)
ElseIf gt(39) <> Minute(Time) Then
gt(39) = Minute(Time)
ElseIf gt(40) <> Hour(Time) Then
gt(40) = Hour(Time)
ElseIf gt(41) <> Day(Date) Then
gt(41) = Day(Date)
ElseIf gt(42) <> Month(Date) Then
gt(42) = Month(Date)
ElseIf gt(43) <> Year(Date) Then
gt(43) = Year(Date)
Else
GoTo Nochg
End If
Case 1
If gt(39) <> Minute(Time) Then
gt(39) = Minute(Time)
ElseIf gt(40) <> Hour(Time) Then
gt(40) = Hour(Time)
ElseIf gt(41) <> Day(Date) Then
gt(41) = Day(Date)
ElseIf gt(42) <> Month(Date) Then
gt(42) = Month(Date)
ElseIf gt(43) <> Year(Date) Then
gt(43) = Year(Date)
Else
GoTo Nochg
End If
Case 2
If gt(40) <> Hour(Time) Then
gt(40) = Hour(Time)
ElseIf gt(41) <> Day(Date) Then
gt(41) = Day(Date)
ElseIf gt(42) <> Month(Date) Then
gt(42) = Month(Date)
ElseIf gt(43) <> Year(Date) Then
gt(43) = Year(Date)
Else
GoTo Nochg
End If
Case 3
If gt(41) <> Day(Date) Then
gt(41) = Day(Date)
ElseIf gt(42) <> Month(Date) Then
gt(42) = Month(Date)
ElseIf gt(43) <> Year(Date) Then
gt(43) = Year(Date)
Else
GoTo Nochg
End If
Case 4
If gt(42) <> Month(Date) Then
gt(42) = Month(Date)
ElseIf gt(43) <> Year(Date) Then
gt(43) = Year(Date)
Else
GoTo Nochg
End If
Case 5
If gt(43) <> Year(Date) Then
gt(43) = Year(Date)
Else
GoTo Nochg
End If
End Select

genrandomseed
seedchanged = True
Nochg:
Rnd -1     'reset seeding
Randomize (gt(35))
End Sub
Public Sub genrandomseed(Optional Zgenoptsgen As Boolean = False)
Dim newrandomseed As Long
Do While True
Randomize CSng((1 + 2 * CSng(Time)) * CSng(Date) ^ (1 + 2 * CSng(Time)))  'Need a unique seed for each change
'Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
newrandomseed = CLng(Int(50000000 * Rnd + 1))
If newrandomseed <> gt(35) And newrandomseed > 0 Then Exit Do
Loop
gt(35) = newrandomseed
If Zgenoptsgen = True Then genoptsgen = True
End Sub
Public Sub outputforformat(argu As Single, gtindex As Long, zerosindex As Long, Optional showLTone As Boolean = False)
Dim lft As Long, rht As Long, flowalert As Boolean, atlimit As Boolean
lft = gt(gtindex)
rht = gt(gtindex + 1)
gt(zerosindex) = 0
flowalert = False
atlimit = False
'2 loops to ensure no overlapping when adding roundoff
'Assign right, left values in gt

For ct = 1 To 2
If argu >= 999.99 Then
atlimit = True
lft = 999
rht = 99
ElseIf argu >= 100 Then
lft = Val(Left(argu, 3))
rht = Val(Mid(argu, 5, 2))
    If Val(Mid(CStr(argu), 7, 1)) >= 5 Then
    If rht = 99 Then
    rht = 0
    lft = lft + 1
    argu = Val(CStr(lft) & decsep & CStr(rht))
    flowalert = True
    Else
    rht = rht + 1
    If lft = 999 And rht = 99 Then atlimit = True
    End If
    End If
ElseIf argu >= 10 Then
lft = Val(Left(argu, 2))
rht = Val(Mid(argu, 4, 3))
    If Val(Mid(CStr(argu), 7, 1)) >= 5 Then
    If rht = 999 Then
    rht = 0
    lft = lft + 1
    argu = Val(CStr(lft) & decsep & CStr(rht))
    flowalert = True
    Else
    rht = rht + 1
    If lft = 99 And rht = 999 Then atlimit = True
    End If
    End If
ElseIf argu >= 1 Then
lft = Val(Left(argu, 1))
rht = Val(Mid(argu, 3, 4))
    If Val(Mid(CStr(argu), 7, 1)) >= 5 Then
    If rht = 9999 Then
    rht = 0
    lft = lft + 1
    argu = Val(CStr(lft) & decsep & CStr(rht))
    flowalert = True
    Else
    rht = rht + 1
    If lft = 9 And rht = 9999 Then atlimit = True
    End If
    End If
ElseIf showLTone = False Then
atlimit = True
lft = 0
rht = 9999
Else
lft = 0
rht = Val(Mid(argu, 3, 4))
    If Val(Mid(CStr(argu), 7, 1)) >= 5 Then
    If rht = 99999 Then
    rht = 0
    lft = lft + 1
    argu = Val(CStr(lft) & decsep & CStr(rht))
    flowalert = True
    Else
    rht = rht + 1
    If rht = 99999 Then atlimit = True
    End If
    End If
End If
If flowalert = False Then Exit For
Next



If atlimit = False Then
flowalert = False
'Now if rht of argu has leading zeros; flowalert used for different purpose
For ct = 1 To 5
If flowalert = True Then
If Mid(CStr(argu), ct, 1) = "0" Then
gt(zerosindex) = gt(zerosindex) + 1
Else
Exit For
End If

End If

If Mid(CStr(argu), ct, 1) = "." Then flowalert = True
Next

End If


If Len(CStr(lft)) + gt(zerosindex) + Len(CStr(rht)) > 5 Then rht = Left(CStr(rht), Len(CStr(rht)) - 1)


gt(gtindex) = lft
gt(gtindex + 1) = rht

If Len(CStr(lft)) + Len(CStr(rht)) > 4 Then gt(zerosindex) = 0

End Sub
Public Function inputformat(gtindex As Long) As String
Dim tmpzeros As String
tmpzeros = ""
'Sort out leading zeros mess
Select Case gtindex
Case 29
For ct = 1 To gt(137)
tmpzeros = tmpzeros & "0"
Next
Case 31
For ct = 1 To gt(138)
tmpzeros = tmpzeros & "0"
Next
Case 33
For ct = 1 To gt(139)
tmpzeros = tmpzeros & "0"
Next
Case 104
For ct = 1 To gt(140)
tmpzeros = tmpzeros & "0"
Next
Case 106
For ct = 1 To gt(141)
tmpzeros = tmpzeros & "0"
Next
Case 108
For ct = 1 To gt(142)
tmpzeros = tmpzeros & "0"
Next
Case 110
For ct = 1 To gt(143)
tmpzeros = tmpzeros & "0"
Next
Case 112
For ct = 1 To gt(144)
tmpzeros = tmpzeros & "0"
Next
Case 114
For ct = 1 To gt(145)
tmpzeros = tmpzeros & "0"
Next
Case 116
For ct = 1 To gt(146)
tmpzeros = tmpzeros & "0"
Next
Case 118
For ct = 1 To gt(147)
tmpzeros = tmpzeros & "0"
Next
Case 120
For ct = 1 To gt(148)
tmpzeros = tmpzeros & "0"
Next
Case 198
For ct = 1 To gt(197)
tmpzeros = tmpzeros & "0"
Next
End Select

If gt(gtindex + 1) = 0 Then
inputformat = CStr(gt(gtindex))
Else
inputformat = CStr(gt(gtindex)) & decsep & tmpzeros & CStr(gt(gtindex + 1))
End If

End Function
Public Sub Restoredefaults(fromscratch As Boolean, Optional oldverzion As Boolean = False)
Dim tempstr As String

Select Case gt(158)
Case 0  'saved settings
Restoresaved fromscratch
'non zero after restore
Case 2     'generic

If oldverzion = False Then
gt(1) = 1       'cornfig ordered
gt(3) = 2       'degree of title
gt(4) = 1       'subst multiplier
gt(5) = 0       'nats multiplier
gt(6) = 1       'spindir %
gt(7) = 2       'spindur %
gt(8) = 2       'spinspeed %
gt(9) = 1       'All wins pay
gt(10) = 10     'spins before MB
gt(12) = 1      'MB after this
gt(14) = 0      'RJ active
gt(20) = 1      'current bet
gt(21) = 1      'bet mult toggle
gt(22) = 2      '"
gt(23) = 5      '"
gt(24) = 0      '"
gt(25) = 0      '"
gt(36) = 0      'seed change
gt(37) = 1      'special changes
gt(44) = 0      'crops not square in cfgthumb
gt(45) = 0      'new game after gametype
gt(46) = 0      'view about
gt(47) = 100000000
End If  'oldverzion

If fromscratch = True Or oldverzion = True Then gt(152) = 0 'MonteCarlo
gt(153) = 0     'number of lines played

If oldverzion = False Then
gt(155) = 250   'Starting cash on hand
gt(157) = 1     'FSFG times
gt(158) = 1     'Default status - current
gt(159) = 0     'Small pictures
gt(160) = 2     'pz flashes

gt(161) = -1    'wallpaper colour
gt(162) = -1    'Title colour
gt(163) = &H933EA2         'wallpaper forecolour
gt(164) = &HC26D74         'title forecolour

gt(165) = &HF093CD
gt(166) = &HF0B96F
gt(167) = &HA6F15A
gt(168) = &H86ABCA
gt(169) = &HC9C7B6
gt(170) = &H7891D
gt(171) = &H4040&
gt(172) = &H808080
gt(173) = &HFFC0C0
gt(174) = &H555044


gt(175) = &H40C0&             'Prize Money colour
gt(176) = &HFFFF&             'Highlight colour
gt(177) = &HFF80FF            'Highlight Text
gt(178) = &H8080FF            'Winning Lines colour
gt(179) = &H4040&             'Spin button forecolour
gt(180) = 6                   'Spin button style
gt(181) = gt(179)             'Bet button forecolour
gt(182) = 6                   'Bet button style
gt(183) = 0                   'List only currently assigned sounds
gt(184) = 0                   'FSFG bonus all paylines
gt(187) = -1                  'Randomise background midis
gt(186) = 3                   'Wave Silence
gt(189) = 0                   'No Textfile write after Configure Thumbnails
gt(193) = 1                   'Use Basedir on new profile
gt(195) = 1                   'Randomise Quotes
End If  'oldverzion

gt(190) = 1                   'shortcut on new config

If oldverzion = False Then

Stringvars(1) = App.Path & "\news.jpg"  'Default bitmap
Stringvars(2) = Stringvars(1)
If gt(200) = 1 Then   'installation
Stringvars(3) = App.Path & "\"
Else
gt(185) = 0                   'Midi port
If Stringvars(3) <> "" Then Stringvars(3) = App.Path & "\"
End If
Stringvars(4) = App.Path & "\"

Stringvars(5) = "MyReels_Game!"
Stringvars(6) = "Spin A $Win"
Stringvars(7) = "Multiply $Win"
Stringvars(8) = "MyReels_Player"
For ct = 9 To 12
Stringvars(ct) = "Times New Roman"
Next

Stringvars(13) = "Welcome, MyReels_Player!"
Stringvars(14) = Space$(1024)
GetComputerName Stringvars(14), Len(Stringvars(14))

For ct = 1 To Len(Stringvars(14))
    If Namevalid(False, Mid(Stringvars(14), ct, 1)) = True Then
    tempstr = tempstr & Mid(Stringvars(14), ct, 1)
    Else
    Stringvars(14) = tempstr
    Exit For
    End If
Next
Stringvars(14) = "Go, " & Stringvars(14) & "!"
Stringvars(15) = "And Rightly So!"
Stringvars(16) = "Hidden Treasure!"
Stringvars(17) = "Reward Of Potential!"
Stringvars(18) = "Reward Of True Potential!"
Stringvars(19) = "Reward Of Highest Potential!"
Stringvars(20) = "Reward of Plenty!"
Stringvars(21) = "Reward of Great Plenty!"
Stringvars(22) = "Award of Abundance!"
Stringvars(23) = "Award of Great Abundance!"
Stringvars(24) = "Jackpot Awarded!"
Stringvars(25) = "Grand Bonanza!"
SetSndDef
SetMusDef
ElseIf gt(192) < 3 Then
Stringvars(3) = ""
End If
End Select
End Sub
Public Function SetSndDef() As Boolean
Stringvars(26) = App.Path & "\spin.wav" 'Spin
Stringvars(27) = App.Path & "\chgbet.wav"  'Change bet
Stringvars(28) = App.Path & "\MB.wav"  'MB
Stringvars(29) = App.Path & "\RJ.wav"  'RJ
Stringvars(30) = App.Path & "\1to4.wav" '1 - 4
Stringvars(31) = App.Path & "\5to9.wav"  '5 - 9
Stringvars(32) = App.Path & "\10to24.wav"  '10 - 24
Stringvars(33) = App.Path & "\25to49.wav"  '25 - 49
Stringvars(34) = App.Path & "\50to99.wav"  '50 - 99
Stringvars(35) = App.Path & "\100to249.wav"  '100  - 249
Stringvars(36) = App.Path & "\250to999.wav"  '250 - 999
Stringvars(37) = App.Path & "\1000to4999.wav"  '1000 - 4999
Stringvars(38) = App.Path & "\5000.wav"  '5000 +
SetSndDef = FileExists(Stringvars(26))
End Function
Public Function SetMusDef() As Boolean
Stringvars(39) = App.Path & "\intro.mid"    'intro
Stringvars(40) = App.Path & "\win1.mid"     'pz 25 - 99
Stringvars(41) = App.Path & "\win2.mid"     'pz 100 - 249
Stringvars(42) = App.Path & "\win3.mid"     'pz 250 - 999
Stringvars(43) = App.Path & "\win4.mid"     'pz 1000 +
Stringvars(44) = App.Path & "\bkgd1.mid"
Stringvars(45) = App.Path & "\bkgd2.mid"
Stringvars(46) = App.Path & "\bkgd3.mid"
Stringvars(47) = App.Path & "\bkgd4.mid"
Stringvars(48) = App.Path & "\bkgd5.mid"
Stringvars(49) = App.Path & "\bkgd6.mid"
Stringvars(50) = App.Path & "\bkgd7.mid"
SetMusDef = FileExists(Stringvars(39))
End Function
Public Function FileExists(ByVal sFileName As String) As Boolean
On Error Resume Next
FileExists = (GetAttr(sFileName) And vbDirectory) <> vbDirectory
On Error GoTo 0
End Function
Public Sub cashup(ztemp As Long, zcolour As Long, captval As String)
captval = ""
Select Case ztemp
Case 50, 100, 150
captval = cash1
zcolour = &HFFFF00
Case 200, 250, 300, 350
captval = cash2
zcolour = vbButtonFace
Case 400, 450, 500
captval = cash3
zcolour = &H80FF&
End Select
captval = "             ====> New Game with this starting cash ====>           " & captval
End Sub
Public Sub valuecomments(argu1 As Long, argu2 As Long, zcolour As Long, captval As String)
If gt(10) > 0 Then
argu1 = 31
argu2 = 32
Else
argu1 = 29
argu2 = 30
End If
Select Case gt(argu1)
Case 0 To 79
captval = return1
zcolour = &HFFFF00
Case 80 To 99
captval = return2
zcolour = vbButtonText
Case Else
captval = return3
zcolour = &H80FF&
End Select
End Sub
Public Function OpenDb(sDatabaseName As String, opentype As Long)

Dim sConnect As String, dbTemp As Database, Halltemp As String

OpenDb = False

On Error GoTo OpenError

If opentype = 1 Then
Screen.MousePointer = vbHourglass
If sDatabaseName = "" Then
If Len(Stringvars(4)) = 0 Then Stringvars(4) = loaddirectory
sDatabaseName = Stringvars(4) & "Hallfame.s$t"
End If
End If

gsDBName = sDatabaseName


'set the connect string
gsDataType = "Microsoft Access"
sConnect = ";pwd=h$tsl$ts"
  

'setup the DBEngine
DBEngine.DefaultUser = "NewUser"
DBEngine.DefaultPassword = vbNullString



OneMoreTry:

Set gwsMainWS = DBEngine.CreateWorkspace("MainWS", "admin", vbNullString)


Set dbTemp = gwsMainWS.OpenDatabase(sDatabaseName, True, False, sConnect)


'success
Set gdbCurrentDB = dbTemp

OpenDb = True
If opentype = 1 Then Screen.MousePointer = vbDefault
Exit Function

AttemptRepair:
Screen.MousePointer = vbHourglass
DBEngine.RepairDatabase gsDBName
Screen.MousePointer = vbDefault
GoTo OneMoreTry

OpenError:

Select Case Err
Case 55, 70
DoEvents
If MsgBox(Err.Description & vbCrLf & vbCrLf & "Retry File Open?", 4 + 48) = vbYes Then Resume OneMoreTry
Case 3049
If MsgBox(Err.Description & vbCrLf & vbCrLf & "Attempt to Repair it?", 4 + 48) = vbYes Then
Resume AttemptRepair
End If
Case 3031
'password protected database
MsgBox "Password Protected", vbOKOnly
End Select
gsDBName = vbNullString
gsDataType = vbNullString

'check for common dialog cancelled
If Err <> 32755 And Err <> 3049 Then killdb sDatabaseName
End Function
Public Sub killdb(sDatabaseName)
On Error GoTo DBCloseErr
If Err.Number = 0 Then
Set gdbCurrentDB = Nothing
Set gdbCurrentDB = gwsMainWS.OpenDatabase(sDatabaseName, True, False, ";pwd=h$tsl$ts")
'THIS IS WHERE WE CHANGE PASSWORD
'gdbCurrentDB.NewPassword "test", "h$tsl$ts"
End If
Set gdbCurrentDB = Nothing
Set gwsMainWS = Nothing
sDatabaseName = ""
Exit Sub
DBCloseErr:
sDatabaseName = ""
ShowError
End Sub
Public Function compactdb(dbtype As Long)
Dim sOldName As String, sNewName As String, sNewName2 As String, workpath As String, cpactfname As String

compactdb = False



Select Case dbtype
Case 1
workpath = Stringvars(4)
Screen.MousePointer = vbHourglass
cpactfname = "Hallfame.s$t"
Case 2
workpath = Left(sDatabaseName, Len(sDatabaseName) - 10)
cpactfname = "Quotes.s$t"
Case 3
workpath = loaddirectory
cpactfname = "Slotdata.s$t"
End Select


If Len(workpath) > 0 Then
sOldName = workpath
sNewName = sOldName & "sl@t1.s$t"
sNewName2 = sOldName & "sl@t2.s$t"
sOldName = workpath & cpactfname
Else
sNewName = loaddirectory & "sl@t1.s$t"
sNewName2 = loaddirectory & "sl@t2.s$t"
sOldName = loaddirectory & cpactfname
End If

On Error GoTo CopyError
FileCopy sOldName, sNewName

If OpenDb(sNewName, dbtype) = False Then
Screen.MousePointer = vbDefault
Kill sNewName
killdb sOldName
Exit Function
End If

DBEngine.CompactDatabase sOldName, sNewName2, dbLangGeneral, dbVersion30, ";pwd=h$tsl$ts"


Kill sOldName
Name sNewName2 As sOldName 'rename the new one to the original name
compactdb = True
killdb sOldName
Kill sNewName

If dbtype = 1 Then Screen.MousePointer = vbDefault


Exit Function

CopyError:
ShowError
Screen.MousePointer = vbDefault
End Function
Public Sub ShowError()
Dim stmp As String

If Err.Number = 0 Then Exit Sub
Screen.MousePointer = vbDefault
stmp = "The following Error occurred:" & vbCrLf & vbCrLf
'add the error string
stmp = stmp & Err.Description & vbCrLf
'add the error number
stmp = stmp & "Number: " & Err
Beep
response = MsgBox(stmp, vbOKOnly)
Err.Clear
End Sub
Public Sub Maketransparent(objframe As Frame)
'Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Doesn't work with picbox

Dim style As Long
'style = GetWindowLong(objframe.hWnd, GWL_EXSTYLE)
style = style Or WS_EX_TRANSPARENT
'style = SetWindowLong(objframe.hWnd, GWL_EXSTYLE, style)
objframe.Refresh
End Sub
Public Function CaptureWindow(ByVal hWndSrc As Long, ByVal Client As Boolean, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Picture
Dim hDCMemory As Long, hBmp As Long, hBmpPrev As Long, r As Long, hDCSrc As Long, hPal As Long, RasterCapsScrn As Long, pic As PicBmp
'IPicture requires a reference to "Standard OLE Types."
Dim IPic As IPicture, IID_IDispatch As GUID



hDCSrc = GetWindowDC(hWndSrc) ' Get device context for entire
' Create a memory device context for the copy process.
hDCMemory = CreateCompatibleDC(hDCSrc)
' Create a bitmap and place it in the memory DC.
hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
hBmpPrev = SelectObject(hDCMemory, hBmp)

' Get screen properties.
RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster capabilities

' Copy the on-screen image into the memory DC.
r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy)

' Remove the new copy of the  on-screen image.
hBmp = SelectObject(hDCMemory, hBmpPrev)

' Release the device context resources back to the system.
r = DeleteDC(hDCMemory)
r = ReleaseDC(hWndSrc, hDCSrc)
           

' Fill in with IDispatch Interface ID.
With IID_IDispatch
.Data1 = &H20400
.Data4(0) = &HC0
.Data4(7) = &H46
End With
' Fill Pic with necessary parts.
With pic
.Size = Len(pic)          ' Length of structure.
.Type = vbPicTypeBitmap   ' Type of Picture (bitmap).
.hBmp = hBmp              ' Handle to bitmap.
.hPal = hPal              ' Handle to palette (may be null).
End With
' Create Picture object.
r = OleCreatePictureIndirect(pic, IID_IDispatch, 1, IPic)
' Return the new Picture object.
           
           
Set CaptureWindow = IPic
End Function
Public Function Openquotes(textfiletoopen As String, opennow As Boolean, Optional zposition As Long, Optional justdelete As Boolean)
Dim nextline As String, quotetemp As String, quotelength As Long


Openquotes = False

On Error GoTo quoteserr


If OpenDb(sDatabaseName, 2) = False Then Exit Function

Set dbsCurrent = gdbCurrentDB

  If opennow = True Then
  Set rectemp = dbsCurrent.OpenRecordset("Quotez")
  Else
  Set rectemp = dbsCurrent.OpenRecordset("Quotez", dbOpenForwardOnly)
  End If



If opennow = True Then 'empty Quotes.s$t
  If textfiletoopen = "Dummy.txt" Then    'basedir option

    With rectemp
    ct = 0
    On Error GoTo NillRecordz
    .MoveFirst

    Do Until .EOF
    ct = ct + 1
    .MoveNext
    Loop
    gt(191) = ct
    gt(188) = ct - 1

    End With

    Else

    With rectemp
      If Not (.EOF And .BOF) Then
      .MoveLast
      .MoveFirst
      Do Until .EOF
      .DELETE
      .MoveNext
      Loop
      End If
    .Close
    End With
    Set rectemp = Nothing
    Set dbsCurrent = Nothing
    killdb sDatabaseName
    sDatabaseName = loaddirectory & "Quotes.s$t"

    'Now compact it
    If compactdb(2) = False Then Exit Function

      If OpenDb(sDatabaseName, 2) = False Then Exit Function
      Set dbsCurrent = gdbCurrentDB
      Set rectemp = dbsCurrent.OpenRecordset("Quotez")
      With rectemp
      .Index = "Qorder"

        If justdelete = False Then
        hfile = FreeFile
        Open textfiletoopen For Input As #hfile
        ct = 0
          If LOF(hfile) > 0 Then

          Do Until EOF(hfile)
          Line Input #hfile, nextline

          If Len(nextline) > 172 Then nextline = Left(nextline, 172)
            If nextline <> "" Then
            If ct > 32767 Then Exit Do
            ct = ct + 1
            .AddNew
            ![Bmpindex] = 0
            ![Quotestr] = nextline
            ![Bmpfile] = Null
            .Update
            End If
          Loop
          Close #hfile
          End If
          If ct > 0 Then
          gt(191) = ct
          gt(188) = ct - 1
          Stringvars(3) = loaddirectory   'Success!
          Else
          Stringvars(3) = ""
          End If
        End If  'justdelete
        End With
    End If
Else    'opennow = False
'On load only



'Seek record
  If gt(200) = 1 Then 'first load
  With rectemp
  'set gt191
  ct = 0
  if isMainActive = True then 
  Do Until .EOF
  If ![Quotestr] = "If these words, bye and bye, Lose appeal  to the eye 'Tis simple work to change     the text: <Enter>, General, Text Tab next, And a filename you supply." Then
  zposition = ct
  End If
  ct = ct + 1
  .MoveNext
  Loop
  else
  Do Until .EOF
  ct = ct + 1
  .MoveNext
  Loop
  zposition = CLng(Int(Rnd * ct))
  end if

  End With
  Set rectemp = Nothing
  Set rectemp = dbsCurrent.OpenRecordset("Quotez")
  With rectemp
  .Index = gdbCurrentDB.TableDefs(.Name).Indexes(0).Name
  End With

  Set rectemp = Nothing
  Set rectemp = dbsCurrent.OpenRecordset("Quotez")

  gt(191) = ct
  gt(188) = ct - 1

  Else
    If gt(195) = 1 Then
    zposition = CLng(Int(Rnd * gt(191)))
    Else
    gt(188) = gt(188) + 1
    If gt(188) = gt(191) Then gt(188) = 0
    zposition = gt(188)
    End If
  End If



With rectemp
.Move zposition
quotetemp = ![Quotestr]
quotelength = Len(quotetemp)

For ct = 0 To Int(quotelength / 44)
splitquote ct, quotelength, quotetemp
Next
If quotetemp <> "" And ct < 4 Then quotestring(ct) = quotetemp

BmpIX = ![Bmpindex]
    
End With
    
    
  If BmpIX > 0 Then
    
  Set rectemp = Nothing
  Set rectemp = dbsCurrent.OpenRecordset("Quotez")

    With rectemp
    .MoveFirst
    .Move zposition
    If IZempty = True Then
    .MoveFirst
    For ct = 0 To gt(191) - 1
    If ![Bmpindex] = BmpIX Then
    If IZempty = False Then Exit For 'found bmp here
    End If
    .MoveNext
    'No pic
    Next
    End If
    

  'Open source file.

  hfile = FreeFile


  'Get size of field.
  ct = ![Bmpfile].FieldSize()

  End With

  If ct = 0 Then
  MsgBox "Quote Thumbnail not found!"
  BmpIX = 0
  Else
        
    If Chunker(rectemp, False) = False Then GoTo quoteserr
    
    Pokemach.Quotez.Visible = True
    Pokemach.Quotez.Picture = LoadPicture(loaddirectory & "q0.bmp")
    For ct = 10 To 13
    Pokemach.lblmisc(ct).Left = resX * 5680
    Next
    End If  'ct
    Else    'ct1
    For ct = 10 To 13
    Pokemach.lblmisc(ct).Left = resX * 6000
    Next
    Pokemach.Quotez.Visible = False
    Set Pokemach.Quotez.Picture = Nothing
  End If  'ct1
    
    
rectemp.Close

End If


Set rectemp = Nothing
Set dbsCurrent = Nothing
killdb sDatabaseName


Openquotes = True
Exit Function

NillRecordz:
If Err.Number = 3021 Then
Openquotes = True
gt(191) = 0
gt(188) = 0
End If


quoteserr:
BmpIX = 0
Close
Set rectemp = Nothing
Set dbsCurrent = Nothing
ShowError
killdb sDatabaseName
End Function
Private Sub splitquote(splitno As Long, quotelength As Long, quoteztr As String)
Dim tempstr As String
tempstr = ""

'Funny ampersand  and _ problem
For ct1 = 44 To 1 Step -1
tempstr = Mid(quoteztr, ct1, 3)
If tempstr = " & " Then quoteztr = Mid(quoteztr, 1, ct1 - 1) & " && " & Mid(quoteztr, ct1 + 3, 175)
Next


If quotelength < 44 Then
quotestring(splitno) = quoteztr
quoteztr = ""
Else
For ct1 = 44 To 1 Step -1
tempstr = Mid(quoteztr, ct1, 1)
  If tempstr = " " Then
  quotestring(splitno) = Left(quoteztr, ct1 - 1)
  quoteztr = Right(quoteztr, Len(quoteztr) - ct1)
  Exit For
  End If
Next


If ct1 = 0 Then  'no spaces
quotestring(splitno) = Left(quoteztr, 43)
quoteztr = Right(quoteztr, Len(quoteztr) - 43)
End If
End If
quotelength = Len(quoteztr)
End Sub
Public Function infofortt(newpicno As Long, pathsel As String)
infofortt = newpicno & " X 43k for pics, 56k for Slotdata, " & findafile(loaddirectory, "quotes.s$t") & "k to copy quotes. Free space : " & vbGetAvailableKBytesAsString(pathsel) & "k"
End Function
