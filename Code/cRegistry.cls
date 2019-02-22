VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================================
' Class:    cRegistry
' Author:   Steve McMahon
' Date  :   21 Feb 1997
'
' A nice class wrapper around the registry functions
' Allows searching,deletion,modification and addition
' of Keys or Values.
'
' Updated 29 April 1998 for VB5.
'   * Fixed GPF in EnumerateValues
'   * Added support for all registry types, not just strings
'   * Put all declares in local class
'   * Added VB5 Enums
'   * Added CreateKey and DeleteKey methods
'
' Updated 2 January 1999
'   * The CreateExeAssociation method failed to set up the
'     association correctly if the optional document icon was not provided.
'   * Added new parameters to CreateExeAssociation to set up other standard handlers: Print, Add, New
'   * Provided the CreateAdditionalEXEAssociations method
'     to allow non-standard menu items to be added (for example,
'     right click on a .VBP file.  VB installs Run and Make menu items).
'
' Updated 8 February 2000
'   * Ensure CreateExeAssociation and related items sets up the
'     registry keys in the
'           HKEY_LOCAL_MACHINE\SOFTWARE\Classes
'     branch as well as the HKEY_CLASSES_ROOT branch.
'
' Updated 23 January 2004
'   * Fixed problem with saving zero length strings. Thanks to Shane Marsden for the fix.
'   * Fixed problem with truncation of binary data.  Thanks to Morten Egelund Rasmussen.

'For RunAsAdmin
Public Enum EShellShowConstants
        essSW_HIDE = 0
        essSW_SHOWNORMAL = 1
        essSW_SHOWMINIMIZED = 2
        essSW_MAXIMIZE = 3
        essSW_SHOWMAXIMIZED = 3
        essSW_SHOWNOACTIVATE = 4
        essSW_SHOW = 5
        essSW_MINIMIZE = 6
        essSW_SHOWMINNOACTIVE = 7
        essSW_SHOWNA = 8
        essSW_RESTORE = 9
        essSW_SHOWDEFAULT = 10
End Enum

Private Type SHELLEXECUTEINFO
        cbSize        As Long
        fMask         As Long
        hWnd          As Long
        lpVerb        As String
        lpFile        As String
        lpParameters  As String
        lpDirectory   As String
        nShow         As Long
        hInstApp      As Long
        lpIDList      As Long     'Optional
        lpClass       As String   'Optional
        hkeyClass     As Long     'Optional
        dwHotKey      As Long     'Optional
        hIcon         As Long     'Optional
        hProcess      As Long     'Optional
End Type

Const TOKEN_QUERY As Long = &H8

Private Enum TOKEN_INFORMATION_CLASS 'Stripped to essentials.
    TokenElevationType = 18
    TokenElevation = 20
End Enum

Private Enum TOKEN_ELEVATION_TYPE
    TokenElevationTypeDefault = 1
    TokenElevationTypeFull ' = 1 etc
    TokenElevationTypeLimited
End Enum


' Check prev instance

Const STANDARD_RIGHTS_REQUIRED              As Long = &HF0000
Const SYNCHRONIZE                           As Long = &H100000
Const MUTANT_QUERY_STATE                    As Long = &H1
Const MUTANT_ALL_ACCESS                     As Long = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or MUTANT_QUERY_STATE)
Const SECURITY_DESCRIPTOR_REVISION          As Long = 1
Const DACL_SECURITY_INFORMATION             As Long = 4
Const App_GUID                              As String = "Global\20b70e57-1c2e-4de9-99e5-20f3961e6912" '20b70e57-1c2e-4de9-99e5-20f3961e6812'

Private m_hCurrentMutex                     As Long


Private Type SECURITY_DESCRIPTOR
    Revision            As Byte
    Sbz1                As Byte
    Control             As Long
    Owner               As Long
    Group               As Long
    pSacl               As Long
    pDacl               As Long
End Type








'Registry Specific Access Rights
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20

'Open/Create Options
Const REG_OPTION_NON_VOLATILE = 0&
Const REG_OPTION_VOLATILE = &H1

'Key creation/open disposition
Const REG_CREATED_NEW_KEY = &H1
Const REG_OPENED_EXISTING_KEY = &H2

'masks for the predefined standard access types
Const STANDARD_RIGHTS_ALL = &H1F0000
Const SPECIFIC_RIGHTS_ALL = &HFFFF

'Define severity codes
Const ERROR_SUCCESS = 0&
Const ERROR_ACCESS_DENIED = 5
Const ERROR_INVALID_DATA = 13&
Const ERROR_MORE_DATA = 234 '  dderror
Const ERROR_NO_MORE_ITEMS = 259

'HTMLHelp
Const HH_DISPLAY_TOPIC = &H0         ' select last opened tab, [display a specified topic] ( or 0 as Long)
Const HH_DISPLAY_TOC = &H1           ' select contents tab, [display a specified topic]
Const HH_HELP_CONTEXT = &HF          ' display mapped numeric value in dwData
Const MAX_PATH = 260


'Structures Needed For Registry Prototypes
Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Boolean
End Type

Private Type FileTime
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type


'Monitors
Private Type MONITORINFO
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
End Type

Private Type tagINITCOMMONCONTROLSEX
    dwSize  As Long
    dwICC   As Long
End Type


Private Declare Function InitShell Lib "shell32" Alias "IsUserAnAdmin" () As Long
Private Declare Sub InitCommonControls Lib "comctl32" ()
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As Any) As Boolean

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Private Declare Function HTMLHelp Lib "HHCtrl.ocx" Alias "HtmlHelpA" (ByVal hWndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Long) As Long

' RunAsAdmin
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function IsUserAnAdmin Lib "shell32" Alias "#680" () As Integer
Private Declare Function ShellExecuteEx Lib "shell32.dll" Alias "ShellExecuteExA" (lpSEI As SHELLEXECUTEINFO) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, ByRef TokenHandle As Long) As Long
Private Declare Function GetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal TokenInformationClass As Long, ByRef TokenInformation As Any, ByVal TokenInformationLength As Long, ByRef ReturnLength As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

'check previnstance
Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (lpMutexAttributes As Any, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Private Declare Function OpenMutex Lib "kernel32" Alias "OpenMutexA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Private Declare Function InitializeSecurityDescriptor Lib "advapi32.dll" (pSecurityDescriptor As Any, ByVal dwRevision As Long) As Long
Private Declare Function SetSecurityDescriptorDacl Lib "advapi32.dll" (pSecurityDescriptor As Any, ByVal bDaclPresent As Long, pDacl As Any, ByVal bDaclDefaulted As Long) As Long
Private Declare Function SetKernelObjectSecurity Lib "advapi32.dll" (ByVal Handle As Long, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long


Private Declare Function GetMonitorInfo Lib "user32.dll" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, ByRef lpmi As MONITORINFO) As Boolean
Private Declare Function MonitorFromWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal dwFlags As Long) As Long

Private Const ICC_STANDARD_CLASSES As Long = &H4000&
Private Const MONITOR_DEFAULTTONULL = &H0       'If the monitor is not found, return 0
Private Const MONITOR_DEFAULTTOPRIMARY = &H1    'If the monitor is not found, return the primary monitor
Private Const MONITOR_DEFAULTTONEAREST = &H2    'If the monitor is not found, return the nearest monitor

' https://support.microsoft.com/en-us/kb/253940
' for GetDeviceCaps replacement of screen properties
Private Const HORZRES = 8
Private Const VERTRES = 10
Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hDC As Long, ByVal nIndex As Long) As Long



'OSinfo stuff

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

'windows-defined type OSVERSIONINFO
Private Type OSVERSIONINFO
  OSVSize         As Long         'size, in bytes, of this data structure
  dwVerMajor      As Long         'ie NT 3.51, dwVerMajor = 3; NT 4.0, dwVerMajor = 4. XP, Server 2003 = 5, Vista+
  dwVerMinor      As Long         'ie NT 3.51, dwVerMinor = 51; NT 4.0, dwVerMinor= 0; 95 = 0 98 = 10; ME= 90
  dwBuildNumber   As Long         'NT: build number of the OS
                                  'Win9x: build number of the OS in low-order word.
                                  '       High-order word contains major & minor ver nos.
  PlatformID      As Long         'Identifies the operating system platform: 1 Win 9X Me, 2 Win 2000 NT
  szCSDVersion    As String * 128 'NT: string, such as "Service Pack 3"
                                  'Win9x: 'arbitrary additional information'
End Type





Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

'For Association
Private Declare Function SHChangeNotify Lib "shell32.dll" (ByVal wEventID As Long, ByVal uFlags As Long, ByVal dwItem1 As Long, ByVal dwItem2 As Long) As Long
Const SHCNE_ASSOCCHANGED = &H8000000
Const SHCNF_IDLIST = &H0

'Registry Function Prototypes
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" _
  (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
  ByVal samDesired As Long, phkResult As Long) As Long

Private Declare Function RegSetValueExStr Lib "advapi32" Alias "RegSetValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
   ByVal dwType As Long, ByVal szData As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32" Alias "RegSetValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
   ByVal dwType As Long, szData As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExByte Lib "advapi32" Alias "RegSetValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
   ByVal dwType As Long, szData As Byte, ByVal cbData As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
   
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" _
  (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, _
   ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, _
   lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, _
   lpdwDisposition As Long) As Long

Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" _
  (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, _
   lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, _
   lpcbClass As Long, lpftLastWriteTime As FileTime) As Long

Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" ( _
    ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, _
    ByVal cbName As Long) As Long

Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" _
  (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
   lpcbValueName As Long, ByVal lpReserved As Long, ByVal lpType As Long, _
   ByVal lpData As Long, ByVal lpcbData As Long) As Long
   
Private Declare Function RegEnumValueLong Lib "advapi32.dll" Alias "RegEnumValueA" _
  (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
   lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, _
   lpData As Long, lpcbData As Long) As Long
Private Declare Function RegEnumValueStr Lib "advapi32.dll" Alias "RegEnumValueA" _
  (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
   lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, _
   ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegEnumValueByte Lib "advapi32.dll" Alias "RegEnumValueA" _
  (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
   lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, _
   lpData As Byte, lpcbData As Long) As Long

Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" _
   (ByVal hKey As Long, ByVal lpClass As String, _
   lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, _
   lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, _
   lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, _
   lpftLastWriteTime As Any) As Long

Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
  (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
  (ByVal hKey As Long, ByVal lpValueName As String) As Long

' Other declares:
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long


Public Enum ERegistryClassConstants
HKEY_CLASSES_ROOT = &H80000000
HKEY_CURRENT_USER = &H80000001
HKEY_LOCAL_MACHINE = &H80000002
HKEY_USERS = &H80000003
End Enum

Public Enum ERegistryValueTypes
'Predefined Value Types
REG_NONE = (0)                         'No value type
REG_SZ = (1)                           'Unicode nul terminated string
REG_EXPAND_SZ = (2)                    'Unicode nul terminated string w/enviornment var
REG_BINARY = (3)                       'Free form binary
REG_DWORD = (4)                        '32-bit number
REG_DWORD_LITTLE_ENDIAN = (4)          '32-bit number (same as REG_DWORD)
REG_DWORD_BIG_ENDIAN = (5)             '32-bit number
REG_LINK = (6)                         'Symbolic Link (unicode)
REG_MULTI_SZ = (7)                     'Multiple Unicode strings
REG_RESOURCE_LIST = (8)                'Resource list in the resource map
REG_FULL_RESOURCE_DESCRIPTOR = (9)     'Resource list in the hardware description
REG_RESOURCE_REQUIREMENTS_LIST = (10)
End Enum


'
' Reg Key Security Options
'
Const DELETE = &H10000
Const READ_CONTROL = &H20000
Const WRITE_DAC = &H40000
Const WRITE_OWNER = &H80000
Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))

Private m_MHZval As String
Private m_hClassKey As Long
Private m_sSectionKey As String
Private m_sValueKey As String
Private m_vValue As Variant
Private m_sSetValue As String
Private m_vDefault As Variant
Private m_eValueType As ERegistryValueTypes
Private s As String
Private HWNDStored As Long
Private HDCStored As Long
Private response As Long
Private responseOnce As Boolean
Public Property Get KeyExists() As Boolean
    'KeyExists = bCheckKeyExists( m_hClassKey, m_sSectionKey)
Dim hKey As Long
    If RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, 1, hKey) = ERROR_SUCCESS Then
        KeyExists = True
        RegCloseKey hKey
    Else
        KeyExists = False
    End If
    
End Property
Public Sub CallHelp(Optional ByVal menuClk As Boolean = False)
Dim GetDesktophWnd As Long
If menuClk Then
s = HTMLHelp(GetDesktophWnd, App.HelpFile, HH_HELP_CONTEXT, 1)
Else
s = HTMLHelp(GetDesktophWnd, App.HelpFile, HH_DISPLAY_TOPIC, 0)
End If
End Sub

Public Function CreateKey() As Boolean
Dim tSA As SECURITY_ATTRIBUTES, hKey As Long, lCreate As Long, e As Long

    'Open or Create the key
    e = RegCreateKeyEx(m_hClassKey, m_sSectionKey, 0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, tSA, hKey, lCreate)
    If e Then
        Err.Raise 26001, App.EXEName & ".cRegistry", "Failed to create registry Key! Was MyReels Uninstalled or OS upgraded? Code: '" & m_sSectionKey
        Err.Clear
    Else
        CreateKey = (e = ERROR_SUCCESS)
        'Close the key
        RegCloseKey hKey
    End If
End Function
Public Function DeleteKey() As Boolean
Dim e As Long
    e = RegDeleteKey(m_hClassKey, m_sSectionKey)
    If e Then
        Err.Raise 26001, App.EXEName & ".cRegistry", "Failed to delete registry Key: '" & m_hClassKey & "',Section: '" & m_sSectionKey
        Err.Clear
    Else
        DeleteKey = (e = ERROR_SUCCESS)
    End If
    
End Function
Public Function DeleteValue() As Boolean
Dim e As Long, hKey As Long

    e = RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, KEY_ALL_ACCESS, hKey)
    If e Then
        Err.Raise 26001, App.EXEName & ".cRegistry", "Failed to open key '" & m_hClassKey & "',Section: '" & m_sSectionKey & "' for delete access"
        Err.Clear
    Else
        e = RegDeleteValue(hKey, m_sValueKey)
        If e Then
            Err.Raise 26001, App.EXEName & ".cRegistry", "Failed to delete registry Key: '" & m_hClassKey & "',Section: '" & m_sSectionKey & "',Key: '" & m_sValueKey
            Err.Clear
        Else
            DeleteValue = (e = ERROR_SUCCESS)
        End If
    End If

End Function
Public Property Get Value() As Variant
s = "*"
Value = GetVal()
End Property
Public Property Get CurrentValue() As Variant
s = "\.Current"
CurrentValue = GetVal()
End Property
Public Property Get DefaultValue() As Variant
s = "\.Default"
DefaultValue = GetVal()
End Property
Private Function GetVal() As Variant
Dim vValue As Variant, bufflen As Long, vkey As String, ordType As Long, sBuffer As String, e As Long, hKey As Long, accessMode As Long
GetVal = ""

If VistaorLater Then
accessMode = KEY_EXECUTE
Else
accessMode = KEY_ALL_ACCESS
End If
If s = "*" Then
e = RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0&, accessMode, hKey)
vkey = m_sValueKey
sBuffer = String$(1024, 0)                             ' Allocate Variable Space
bufflen = Len(sBuffer)
Else
e = RegOpenKeyEx(m_hClassKey, m_sSectionKey & "\" & m_sValueKey & s, 0&, accessMode, hKey)
End If


If e = ERROR_SUCCESS Then

  e = RegQueryValueEx(hKey, vkey, 0&, ordType, ByVal sBuffer, bufflen)

  If Not (e And e <> ERROR_MORE_DATA) Then
         
    If bufflen > 0 Then
      If s <> "*" Then
      sBuffer = Space(bufflen)
        If RegQueryValueEx(hKey, vbNullString, 0&, ordType, ByVal sBuffer, Len(sBuffer)) <> ERROR_SUCCESS Then
        ' Close Registry Key
        e = RegCloseKey(hKey)
        Exit Function
        End If
      End If

      If (Asc(Mid(sBuffer, Len(sBuffer), 1)) = 0) Then  'Win95 Adds Null Terminated String...
        sBuffer = Left(sBuffer, Len(sBuffer) - 1)         'Null Found, Extract From String
        Else                                              'WinNT Does NOT Null Terminate String...
        sBuffer = Left(sBuffer, Len(sBuffer))             'Null Not Found, Extract String Only
      End If

    End If

    'Trim NULL and return successful query!
       
    Select Case ordType                                    ' Search Data Types...
    Case REG_SZ, REG_EXPAND_SZ                             ' String Registry Key Data Type
      vValue = sBuffer                                     ' Copy String Value
    Case REG_DWORD                                         ' Double Word Registry Key Data Type
      For e = Len(sBuffer) To 1 Step -1                    ' Convert Each Bit
      vValue = vValue + Hex(Asc(Mid(Len(sBuffer), e, 1)))  ' Build Value Char. By Char.
      Next
      vValue = Format$("&h" + vValue)                      ' Convert Double Word To String
    End Select
       
  End If
        
e = RegCloseKey(hKey)                                  ' Close Registry Key

End If

GetVal = vValue
End Function
Public Property Let Value(ByVal vValue As Variant)
Dim ordType As Long, c As Long, hKey As Long, e As Long, lCreate As Long, tSA As SECURITY_ATTRIBUTES

    'Open or Create the key
    e = RegCreateKeyEx(m_hClassKey, m_sSectionKey, 0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, tSA, hKey, lCreate)
    
    If e Then
        Err.Raise 26001, App.EXEName & ".cRegistry", "Failed to set registry value Key: '" & m_hClassKey & "',Section: '" & m_sSectionKey & "',Key: '" & m_sValueKey & "' to value: '" & m_vValue & "'"
        Err.Clear
    Else

        Select Case m_eValueType
        Case REG_BINARY
            If (VarType(vValue) = vbArray + vbByte) Then
                Dim ab() As Byte
                ab = vValue
                ordType = REG_BINARY
                c = UBound(ab) - LBound(ab) - 1
                e = RegSetValueExByte(hKey, m_sValueKey, 0&, ordType, ab(0), c)
            Else
                Err.Raise 26001
                Err.Clear
            End If
        Case REG_DWORD, REG_DWORD_BIG_ENDIAN, REG_DWORD_LITTLE_ENDIAN
            If (VarType(vValue) = vbInteger) Or (VarType(vValue) = vbLong) Then
                Dim i As Long
                i = vValue
                ordType = REG_DWORD
                e = RegSetValueExLong(hKey, m_sValueKey, 0&, ordType, i, 4)
            End If
        Case REG_SZ, REG_EXPAND_SZ
            Dim s As String, iPos As Long
            s = vValue
            ordType = REG_SZ
            ' Assume anything with two non-adjacent percents is expanded string
            iPos = InStr(s, "%")
            If iPos Then
                If InStr(iPos + 2, s, "%") Then ordType = REG_EXPAND_SZ
            End If
            c = Len(s) + 1
            s = s & vbNullChar ' Thanks to Shane Marsden
            e = RegSetValueExStr(hKey, m_sValueKey, 0&, ordType, s, c)
            
        ' User should convert to a compatible type before calling
        Case Else
            e = ERROR_INVALID_DATA
            
        End Select
        
        If Not e Then
            m_vValue = vValue
        Else
            Err.Raise vbObjectError + 1048 + 26001, App.EXEName & ".cRegistry", "Failed to set registry value Key: '" & m_hClassKey & "',Section: '" & m_sSectionKey & "',Key: '" & m_sValueKey & "' to value: '" & m_vValue & "'"
            Err.Clear
        End If
        
        'Close the key
        RegCloseKey hKey
    
    End If
    
End Property
Public Function EnumerateValues(ByRef sKeyNames() As String, ByRef iKeyCount As Long) As Boolean
Dim lResult As Long, hKey As Long, sName As String, lNameSize As Long
Dim sData As String, lIndex As Long, cJunk As Long, cNameMax As Long, ft As Currency
   
   ' Log "EnterEnumerateValues"

   iKeyCount = 0
   Erase sKeyNames()
    
   lIndex = 0
   lResult = RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, KEY_QUERY_VALUE, hKey)
   If (lResult = ERROR_SUCCESS) Then
      ' Log "OpenedKey:" & m_hClassKey & "," & m_sSectionKey
      lResult = RegQueryInfoKey(hKey, "", cJunk, 0, _
                               cJunk, cJunk, cJunk, cJunk, _
                               cNameMax, cJunk, cJunk, ft)
       Do While lResult = ERROR_SUCCESS
   
           'Set buffer space
           lNameSize = cNameMax + 1
           sName = String$(lNameSize, 0)
           If (lNameSize = 0) Then lNameSize = 1
           
           ' Log "Requesting Next Value"
         
           'Get value name:
           lResult = RegEnumValue(hKey, lIndex, sName, lNameSize, _
                                  0&, 0&, 0&, 0&)
           ' Log "RegEnumValue returned:" & lResult
           If (lResult = ERROR_SUCCESS) Then
       
                ' Although in theory you can also retrieve the actual
                ' value and type here, I found it always (ultimately) resulted in
                ' a GPF, on Win95 and NT.  Why?  Can anyone help?
       
               sName = Left$(sName, lNameSize)
               ' Log "Enumerated value:" & sName
                 
               iKeyCount = iKeyCount + 1
               ReDim Preserve sKeyNames(1 To iKeyCount) As String
               sKeyNames(iKeyCount) = sName
           End If
           lIndex = lIndex + 1
       Loop
   End If
   If (hKey <> 0) Then
      RegCloseKey hKey
   End If

   ' Log "Exit Enumerate Values"
   EnumerateValues = True
   Exit Function
   
EnumerateValuesError:
   If (hKey <> 0) Then
      RegCloseKey hKey
   End If
   Err.Raise vbObjectError + 1048 + 26003, App.EXEName & ".cRegistry", Err.Description
   Err.Clear
   Exit Function

End Function
Public Function EnumerateSections(ByRef sSect() As String, ByRef iSectCount As Long) As Boolean
Dim lResult As Long, hKey As Long, dwReserved As Long, szBuffer As String
Dim lBuffSize As Long, lIndex As Long, lType As Long, iPos As Long
On Error GoTo EnumerateSectionsError

   iSectCount = 0
   Erase sSect
   lIndex = 0

   lResult = RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, KEY_ENUMERATE_SUB_KEYS, hKey)
   Do While lResult = ERROR_SUCCESS
       'Set buffer space
       szBuffer = String$(255, 0)
       lBuffSize = Len(szBuffer)
      
      'Get next value
       lResult = RegEnumKey(hKey, lIndex, szBuffer, lBuffSize)
                             
       If (lResult = ERROR_SUCCESS) Then
          iSectCount = iSectCount + 1
          ReDim Preserve sSect(1 To iSectCount) As String
           iPos = InStr(szBuffer, Chr$(0))
           If (iPos > 0) Then
              sSect(iSectCount) = Left(szBuffer, iPos - 1)
           Else
              sSect(iSectCount) = Left(szBuffer, lBuffSize)
           End If
       End If
       
       lIndex = lIndex + 1
   Loop
   If (hKey <> 0) Then
      RegCloseKey hKey
   End If
   EnumerateSections = True
   Exit Function

EnumerateSectionsError:
   If (hKey <> 0) Then
      RegCloseKey hKey
   End If
   Err.Raise vbObjectError + 1048 + 26002, App.EXEName & ".cRegistry", Err.Description
   Err.Clear
   Exit Function
End Function
Private Sub pSetClassValue(ByVal sValue As String)
    Dim sSection As String
    
    ClassKey = HKEY_CLASSES_ROOT
    Value = sValue
    sSection = SectionKey
    ClassKey = HKEY_LOCAL_MACHINE
    SectionKey = "SOFTWARE\Classes\" & sSection
    Value = sValue
    SectionKey = sSection
End Sub
Public Function CreateEXEAssociation(ByVal sExePath As String, _
        ByVal sClassName As String, _
        ByVal sClassDescription As String, _
        ByVal sAssociation As String, _
        Optional ByVal sOpenMenuText As String = "&Open", _
        Optional ByVal bSupportPrint As Boolean = False, _
        Optional ByVal sPrintMenuText As String = "&Print", _
        Optional ByVal bSupportNew As Boolean = False, _
        Optional ByVal sNewMenuText As String = "&New", _
        Optional ByVal bSupportInstall As Boolean = False, _
        Optional ByVal sInstallMenuText As String = "", _
        Optional ByVal lDefaultIconIndex As Long = -1) As Boolean
   ' Check if path is wrapped in quotes:
CreateEXEAssociation = True

    If VistaorLater() Then
        If Not IsUserAnAdministrator() Then
            MsgBox "MyReels has to be run as administrator on new config. Click OK to Continue", vbInformation, "New Config Run"
            CreateEXEAssociation = False
            Exit Function
        End If
    End If

 
   On Error GoTo EXEAssocerror
   
   sExePath = Trim$(sExePath)
   If (Left$(sExePath, 1) <> """") Then
      sExePath = """" & sExePath
   End If
   If (Right$(sExePath, 1) <> """") Then
      sExePath = sExePath & """"
   End If
   If (sExePath = """") Then
   response = MsgBox("Something went wrong: sExePath is empty.", vbOKOnly)
   End If
    ' Create the .File to Class association:
   SectionKey = "." & sAssociation
   ValueType = REG_SZ
   ValueKey = ""
   pSetClassValue sClassName
   
   ' Create the Class shell open command:
   SectionKey = sClassName
   pSetClassValue sClassDescription
   
   SectionKey = sClassName & "\shell\open"
   If (sOpenMenuText = "") Then sOpenMenuText = "&Open"
   ValueKey = ""
   pSetClassValue sOpenMenuText
   SectionKey = sClassName & "\shell\open\command"
   ValueKey = ""
   pSetClassValue sExePath & " ""%1"""
   
   If (bSupportPrint) Then
      SectionKey = sClassName & "\shell\print"
      If (sPrintMenuText = "") Then sPrintMenuText = "&Print"
      ValueKey = ""
      pSetClassValue sPrintMenuText
      SectionKey = sClassName & "\shell\print\command"
      ValueKey = ""
      pSetClassValue sExePath & " /p ""%1"""
   End If
   
   If (bSupportInstall) Then
      If (sInstallMenuText = "") Then
         sInstallMenuText = "&Install " & sAssociation
      End If
      SectionKey = sClassName & "\shell\add"
      ValueKey = ""
      pSetClassValue sInstallMenuText
      SectionKey = sClassName & "\shell\add\command"
      ValueKey = ""
      pSetClassValue sExePath & " /a ""%1"""
   End If
   
   If (bSupportNew) Then
      SectionKey = sClassName & "\shell\new"
      ValueKey = ""
      If (sNewMenuText = "") Then sNewMenuText = "&New"
      pSetClassValue sNewMenuText
      SectionKey = sClassName & "\shell\new\command"
      ValueKey = ""
      pSetClassValue sExePath & " /n ""%1"""
   End If
   
   If lDefaultIconIndex > -1 Then
      SectionKey = sClassName & "\DefaultIcon"
      ValueKey = ""
      pSetClassValue sExePath & "," & CStr(lDefaultIconIndex)
   End If

If (XP Or VistaorLater) Then SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0

Exit Function
EXEAssocerror:
' always true even if error
ShowError
End Function
Private Sub CreateAdditionalEXEAssociations(ByVal sClassName As String, ParamArray vItems() As Variant)
Dim iItems As Long, iItem As Long
   
   On Error Resume Next
   iItems = UBound(vItems) + 1
   If (iItems Mod 3) <> 0 Or (Err.Number <> 0) Then
      Err.Raise vbObjectError + 1048 + 26004, App.EXEName & ".cRegistry", "Invalid parameter list passed to CreateAdditionalEXEAssociations - expected Name/Text/Command"
      Err.Clear
   Else
      ' Check if it exists:
      SectionKey = sClassName
      If Not (KeyExists) Then
         Err.Raise vbObjectError + 1048 + 26005, App.EXEName & ".cRegistry", "Error - attempt to create additional associations before class defined."
         Err.Clear
      Else
         For iItem = 0 To iItems - 1 Step 3
            ValueType = REG_SZ
            SectionKey = sClassName & "\shell\" & vItems(iItem)
            ValueKey = ""
            pSetClassValue vItems(iItem + 1)
            SectionKey = sClassName & "\shell\" & vItems(iItem) & "\command"
            ValueKey = ""
            pSetClassValue vItems(iItem + 2)
         Next iItem
      End If
   End If
End Sub
Public Property Get ValueType() As ERegistryValueTypes
ValueType = m_eValueType
End Property
Public Property Let ValueType(ByVal eValueType As ERegistryValueTypes)
m_eValueType = eValueType
End Property
Public Property Get ClassKey() As ERegistryClassConstants
ClassKey = m_hClassKey
End Property
Public Property Let ClassKey(ByVal eKey As ERegistryClassConstants)
m_hClassKey = eKey
End Property
Public Property Get SectionKey() As String
SectionKey = m_sSectionKey
End Property
Public Property Let SectionKey(ByVal sSectionKey As String)
m_sSectionKey = sSectionKey
End Property
Public Property Get ValueKey() As String
ValueKey = m_sValueKey
End Property
Public Property Let ValueKey(ByVal sValueKey As String)
m_sValueKey = sValueKey
End Property
Public Property Get Default() As Variant
Default = m_vDefault
End Property
Public Property Let Default(ByVal vDefault As Variant)
m_vDefault = vDefault
End Property
Private Function SwapEndian(ByVal dw As Long) As Long
CopyMemory ByVal VarPtr(SwapEndian) + 3, dw, 1
CopyMemory ByVal VarPtr(SwapEndian) + 2, ByVal VarPtr(dw) + 1, 1
CopyMemory ByVal VarPtr(SwapEndian) + 1, ByVal VarPtr(dw) + 2, 1
CopyMemory SwapEndian, ByVal VarPtr(dw) + 3, 1
End Function
Private Function ExpandEnvStr(sData As String) As String
Dim c As Long, s As String
' Get the length
s = "" ' Needed to get around Windows 95 limitation
c = ExpandEnvironmentStrings(sData, s, c)
' Expand the string
s = String$(c - 1, 0)
c = ExpandEnvironmentStrings(sData, s, c)
ExpandEnvStr = s
End Function
Public Function IsWin2000NT() As Boolean

'returns True if running Windows 2000 (NT5)

    #If Win32 Then
    '64 bit sytems are viewed as 32 bit through WOW64

    Dim OSV As OSVERSIONINFO
   
      OSV.OSVSize = Len(OSV)
   
      If GetVersionEx(OSV) = 1 Then
   
      'PlatformId contains a value representing the OS.
      'If VER_PLATFORM_WIN32_NT and dwVerMajor is 5, return true
      IsWin2000NT = ((OSV.PlatformID = VER_PLATFORM_WIN32_NT) And (OSV.dwVerMajor = 5))
      End If

    #End If

End Function
Public Function VistaorLater() As Boolean

'returns True if running Vista or later (NT5)

    #If Win32 Then
      Dim OSV As OSVERSIONINFO
   
      OSV.OSVSize = Len(OSV)
   
      If GetVersionEx(OSV) = 1 Then
      ' DO NOT run VB6.exe in compatibility mode for anything as this and the following 2 functions break!
      VistaorLater = ((OSV.PlatformID = VER_PLATFORM_WIN32_NT) And (OSV.dwVerMajor >= 6))
      End If

    #End If

End Function
Public Function XP() As Boolean

'returns True if running XP or Server 2003

    #If Win32 Then
      Dim OSV As OSVERSIONINFO
   
      OSV.OSVSize = Len(OSV)
   
      If GetVersionEx(OSV) = 1 Then
      XP = ((OSV.PlatformID = VER_PLATFORM_WIN32_NT) And (OSV.dwVerMajor = 5))
      End If

    #End If

End Function
Public Function EightorLater() As Boolean

'returns True if running W8 or later

    #If Win32 Then
      Dim OSV As OSVERSIONINFO
   
      OSV.OSVSize = Len(OSV)
   
      If GetVersionEx(OSV) = 1 Then
      EightorLater = ((OSV.PlatformID = VER_PLATFORM_WIN32_NT) And (OSV.dwVerMajor >= 8))
      End If

    #End If

End Function
Private Function GetCPUSpeed() As Long
Dim hKey As Long, lClockSpeed As Long, e As Long, strKey As String

If IsWin2000NT = True Or VistaorLater = True Then

strKey = "HARDWARE\DESCRIPTION\System\CentralProcessor\0"
e = RegOpenKeyEx(HKEY_LOCAL_MACHINE, strKey, 0&, KEY_ALL_ACCESS, hKey)
  If e = ERROR_SUCCESS Then

    If RegQueryValueEx(hKey, "~MHz", 0, 0, lClockSpeed, 4) <> ERROR_SUCCESS Then
    e = RegCloseKey(hKey)                                  ' Close Registry Key
    End If

  End If

e = RegCloseKey(hKey)                                  ' Close Registry Key
GetCPUSpeed = CLng(lClockSpeed)
'CLng("&H" & lClockSpeed) doesn't parse the &H

If e <> ERROR_SUCCESS Or GetCPUSpeed = 0 Then response = MsgBox("Problem with obtaining CPU Speed from Registry!", vbOKOnly)


Else
GetCPUSpeed = 4000 'assume super fast for new Win gen.
End If
End Function
Public Property Get RealMHZ() As Long
RealMHZ = GetCPUSpeed
End Property
Public Property Get MHZ() As Long
Dim OSV As OSVERSIONINFO
MHZ = GetCPUSpeed
Select Case GetCPUSpeed
Case 0
'Win9x or error
    
    If OSV.dwVerMinor = 90 Then 'ME
    MHZ = 3
    Else    '0 95, 10 98
    MHZ = 2
    End If


Case Is < 550
MHZ = 2
Case Is < 750
MHZ = 3
Case Is < 1050
MHZ = 5
Case Is < 1650
MHZ = 8
Case Is < 2250
MHZ = 15
Case Is < 2850
MHZ = 23
Case Is < 3250
MHZ = 58
Case Else
MHZ = 72
End Select
End Property
Public Property Get mch2000nt() As Long
mch2000nt = IsWin2000NT
End Property
Public Property Get Left() As Long
Dim MonInf As MONITORINFO
MonInfo MonInf
Left = MonInf.rcMonitor.Left
End Property
Public Property Get Top() As Long
Dim MonInf As MONITORINFO
MonInfo MonInf
Top = MonInf.rcMonitor.Top
End Property
Public Property Get Width() As Long
Dim MonInf As MONITORINFO
MonInfo MonInf
Width = MonInf.rcMonitor.Right - MonInf.rcMonitor.Left
End Property
Public Property Get Height() As Long
Dim MonInf As MONITORINFO
MonInfo MonInf
Height = MonInf.rcMonitor.Bottom - MonInf.rcMonitor.Top
End Property
Public Property Let hWnd(newHwnd As Long)
HWNDStored = newHwnd
End Property
Private Sub MonInfo(MonInfoStruct As MONITORINFO)
Dim hMonitor As Long
hMonitor = MonitorFromWindow(HWNDStored, MONITOR_DEFAULTTONEAREST)
MonInfoStruct.cbSize = Len(MonInfoStruct)
If (GetMonitorInfo(hMonitor, MonInfoStruct) = False) Then
response = MsgBox("GetMonitor Info Failed! Game Screen may not be visible.", vbOKOnly)
End If
End Sub
'Now the Working Directory Attributes
Public Property Get WLeft() As Long
Dim MonInf As MONITORINFO
MonInfo MonInf
WLeft = MonInf.rcWork.Left
End Property
Public Property Get WTop() As Long
Dim MonInf As MONITORINFO
MonInfo MonInf
WTop = MonInf.rcWork.Top
End Property
Public Property Get WWidth() As Long
Dim MonInf As MONITORINFO
MonInfo MonInf
WWidth = MonInf.rcWork.Right - MonInf.rcWork.Left
End Property
Public Property Get WHeight() As Long
Dim MonInf As MONITORINFO
MonInfo MonInf
WHeight = MonInf.rcWork.Bottom - MonInf.rcWork.Top
End Property
Public Property Let hDC(newHdc As Long)
HDCStored = newHdc
End Property
Public Property Get DCWidthTwips() As Long
DCWidthTwips = GetDeviceCaps(HDCStored, HORZRES)
DCWidthTwips = (DCWidthTwips / GetDeviceCaps(HDCStored, LOGPIXELSX)) * 1440
If (Abs(DCWidthTwips - Screen.Width) > 100) Or (Abs(DCHeightTwips - Screen.Height) > 100) Then
'RR global defined in auxrout1
If RR = False Then
response = MsgBox("GetScrDims reports difference between Screen.Width " & Screen.Width & " and GetDeviceCaps Width " & DCWidthTwips & ", or between Screen.Height: " & Screen.Height & " and GetDeviceCaps Height " & DCHeightTwips & ".", vbOKOnly)
RR = True
End If
End If
End Property
Public Property Get DCWidth() As Long
'Note: Useless if form not shown!
DCWidth = GetDeviceCaps(HDCStored, HORZRES)
End Property
Public Property Get DCHeightTwips() As Long
DCHeightTwips = GetDeviceCaps(HDCStored, VERTRES)
DCHeightTwips = (DCHeightTwips / GetDeviceCaps(HDCStored, LOGPIXELSY)) * 1440
End Property
Public Property Get DCHeight() As Long
DCHeight = GetDeviceCaps(HDCStored, VERTRES)
End Property
Public Function RunElevated(ByVal FilePath As String, Optional ShellShowType As EShellShowConstants = essSW_SHOWNORMAL, Optional ByVal hWndOwner As Long = 0, Optional EXEParameters As String = "") As Boolean
    Dim SEI As SHELLEXECUTEINFO
    Const SEE_MASK_DEFAULT = &H0
    
    On Error GoTo Err

    'Fill the SEI structure
    With SEI
        .cbSize = Len(SEI)                  ' Bytes of the structure
        .fMask = SEE_MASK_DEFAULT           ' Check MSDN for more info on Mask
        .lpFile = FilePath                  ' Program Path
        .nShow = ShellShowType              ' How the program will be displayed
        .lpDirectory = App.Path
        .lpParameters = EXEParameters       ' Each parameter must be separated by space. If the lpFile member specifies a document file, lpParameters should be NULL.
        .hWnd = hWndOwner                   ' Owner window handle
        .lpVerb = "runas"
    End With

    RunElevated = ShellExecuteEx(SEI)   ' Execute the program, return success or failure

    Exit Function
Err:
    RunElevated = False
End Function
Public Function ShellOut(ByVal FilePath As String, Optional ShellShowType As EShellShowConstants = essSW_SHOWNORMAL, Optional ByVal hWndOwner As Long = 0, Optional EXEParameters As String = "") As Boolean
    Dim SEI As SHELLEXECUTEINFO
    Const SEE_MASK_DEFAULT = &H0
    
    On Error GoTo Err

    'Fill the SEI structure
    With SEI
        .cbSize = Len(SEI)                  ' Bytes of the structure
        .fMask = SEE_MASK_DEFAULT           ' Check MSDN for more info on Mask
        .lpFile = FilePath                  ' Program Path
        .nShow = ShellShowType              ' How the program will be displayed
        .lpDirectory = ""
        .lpParameters = EXEParameters       ' Each parameter must be separated by space. If the lpFile member specifies a document file, lpParameters should be NULL.
        .hWnd = hWndOwner                   ' Owner window handle
        .lpVerb = ""
    End With

    ShellOut = ShellExecuteEx(SEI)   ' Execute the program, return success or failure

    Exit Function
Err:
    ShellOut = False
End Function
Public Function IsUserAnAdministrator() As Boolean
' dilettante: http://www.vbforums.com/showthread.php?769469-VB6-Elevated-Privileges
    Dim hProcToken As Long
    Dim ElevationType As TOKEN_ELEVATION_TYPE
    Dim TokenIsElevated As Long
    Dim ReturnLength As Long

  On Error GoTo NoCanDo
 
  IsUserAnAdministrator = True

  If IsUserAnAdmin() Then Exit Function

  'If we're on Vista onwards, check for UAC elevation token as we may be an admin but we're not elevated yet, so the IsUserAnAdmin() function will return false

  If VistaorLater() Then

    If OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, hProcToken) Then
      If GetTokenInformation(hProcToken, TokenElevationType, ElevationType, 4, ReturnLength) Then
         If ElevationType = TokenElevationTypeFull Then
            Exit Function
         Else
            If GetTokenInformation(hProcToken, TokenElevation, TokenIsElevated, 4, ReturnLength) Then
               IsUserAnAdministrator = CBool(TokenIsElevated)
               Exit Function
            End If
         End If
      End If
    End If
    Err.Raise &H8004EE00, "IsUserAnAdministrator: ElevatedToken", "System error " & CStr(Err.LastDllError)
        ' Else No UAC
        End If


NoCanDo:
response = MsgBox("IsUserAnAdministrator encountered a problem!", vbOKOnly)
End Function
Public Function MutexChk(Optional MutexInst As Boolean = False)

Dim hMutex As Long, MUTEX_NAME As String, uSec As SECURITY_DESCRIPTOR

  MUTEX_NAME = App_GUID
  MutexChk = False
  hMutex = OpenMutex(MUTANT_ALL_ACCESS, 0, MUTEX_NAME)
  If hMutex <> 0 Then
    MutexChk = True
    Exit Function
  End If
  m_hCurrentMutex = CreateMutex(ByVal 0&, 1, MUTEX_NAME)
  Call InitializeSecurityDescriptor(uSec, SECURITY_DESCRIPTOR_REVISION)
  Call SetSecurityDescriptorDacl(uSec, 1, ByVal 0, 0)
  Call SetKernelObjectSecurity(m_hCurrentMutex, DACL_SECURITY_INFORMATION, uSec)
End Function
Public Sub CloseMutexhandle()
    If m_hCurrentMutex <> 0 Then
        Call CloseHandle(m_hCurrentMutex)
        m_hCurrentMutex = 0
    End If
End Sub
Public Function GetShortName(FileName)
Dim looper As Long, fpiecelength As Long, fpiece As String, testpiece As String
For looper = Len(FileName) To 1 Step -1
testpiece = Mid(FileName, looper, 1)
If testpiece = "." Then
GetShortName = "." & fpiece
fpiece = ""
ElseIf testpiece = "\" Then
If Len(fpiece) > 8 Then fpiece = Left(fpiece, 6) & "~1"
GetShortName = "\" & fpiece & GetShortName
fpiece = ""
fpiecelength = 0
Else
fpiece = testpiece & fpiece
fpiecelength = fpiecelength + 1
End If
Next
End Function
Public Function GetShortFileName(ByVal sFileName As String) As String
    Dim sBuffer As String
    sBuffer = String$(MAX_PATH, vbNullChar)
    Call GetShortPathName(sFileName, sBuffer, MAX_PATH)
    GetShortFileName = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
End Function
' get system32 dir or windir
' OTR (On the run) - Snippet
' ~H7LABS Developement~
' ~haZl0oh~
Public Function SysDir(Optional ByVal SysRt = False) As String
If SysRt Then
SysDir = Environ("SystemRoot")
Else
SysDir = Environ("comspec")
' comspec : length of "cmd.exe" = 7
SysDir = Mid(SysDir, 1, Len(SysDir) - 7)
End If
End Function
'*******************************************************************************
' MODULE:       MStringSortRoutines
' FILENAME:     C:\My Code\vb\Sorting\MStringSortRoutines.bas
' AUTHOR:       Phil Fresle
' CREATED:      01-Dec-1999
' COPYRIGHT:    Copyright 1999 Frez Systems Limited. All Rights Reserved. http://www.frez.co.uk sales@frez.co.uk
'
' DESCRIPTION:
' Generic string sort routines, prefer to use the 'non-pure' Quick Sort unless good
' reason to choose another routine.  Here are the routines in order of efficiency.
'
' Quick Sort     - Fast for large arrays - delegates to insert sort when called with small array and to sort small chunks of large array
'                  This 'non-pure' quick sort is thus quicker by not recursing for small chunks where simple brute-force iteration is quicker.
'Insert Sort    - Fast for small arrays (say less than 60 values)
'
' Due to recursive nature of Quick and Merge sort they are not very efficient for small
' arrays which is why the routines delegate to a more brute force insert sort for small arrays.
'
' MODIFICATION HISTORY:
' 1.0       01-Dec-1999
'           Phil Fresle
'           Initial Version
' 1.1       22-Mar-2000
'           Phil Fresle
'           Added Merge Sort
'*******************************************************************************



' QuickSortStringsStart (SUB)
'
' PARAMETERS:
' (In/Out) - ListArray()    - String  - Array to sort
' (In)     - bAscending     - Boolean - True to sort ascending, false descending
' (In)     - bCaseSensitive - Boolean - True for a case sensitive sort, false
'                                       for an insensitive one
'
' DESCRIPTION:
' User friendly entry point for QuickSortStrings
'*******************************************************************************
' InsertSortStringsStart (SUB)
'
' PARAMETERS:
' (In/Out) - ListArray()    - String  - Array to sort
' (In)     - bAscending     - Boolean - True to sort ascending, false descending
' (In)     - bCaseSensitive - Boolean - True for a case sensitive sort, false for insensitive one
'
' DESCRIPTION:
' User friendly entry point for InsertSortStrings
'*******************************************************************************
Public Sub InsertSortStringsStart(ListArray() As String, Optional ByVal bAscending As Boolean = True, Optional ByVal bCaseSensitive As Boolean = False)

  Dim lMin            As Long
  Dim lMax            As Long
  Dim lOrder          As Long
  Dim lCompareType    As Long

  lMin = LBound(ListArray)
  lMax = gt(196)

  If lMin = lMax Then
  Exit Sub
  End If
  ' Order Ascending or Descending?
  lOrder = IIf(bAscending, 1, -1)

  ' Case sensitive search or not?
  lCompareType = IIf(bCaseSensitive, vbBinaryCompare, vbTextCompare)

  InsertSortStrings ListArray, lMin, lMax, lOrder, lCompareType
End Sub


'*******************************************************************************
' InsertSortStrings (SUB)
'
' PARAMETERS:
' (In/Out) - ListArray()  - String - Array to sort
' (In)     - lMin         - Long   - Start of sorting region within array
' (In)     - lMax         - Long   - End of sorting region within array
' (In)     - lOrder       - Long   - Ascending is -1, Descending is +1, used for comparison in StrComp
' (In)     - lCompareType - Long   - Either vbBinaryCompare or vbTextCompare, used in StrComp function
'
' DESCRIPTION:
' Simple Insert Sort routine for strings, fast for small arrays as no recursion (say less than 60 values)
'*******************************************************************************
Private Sub InsertSortStrings(ListArray() As String, ByVal lMin As Long, ByVal lMax As Long, ByVal lOrder As Long, ByVal lCompareType As Long)

  Dim sValue  As String
  Dim lCount1 As Long
  Dim lCount2 As Long
  ' Loop through array shifting elements down to their correct place
  For lCount1 = lMin + 1 To lMax
      sValue = ListArray(lCount1)


      ' Find place to put it
      For lCount2 = lCount1 - 1 To lMin Step -1
          If StrComp(ListArray(lCount2), sValue, lCompareType) <> lOrder Then
              Exit For
          End If
          ListArray(lCount2 + 1) = ListArray(lCount2)
      Next lCount2

      ' Insert it
      ListArray(lCount2 + 1) = sValue
  Next
End Sub
Public Property Get InitCommonCtrls() As String
    On Error Resume Next
    Dim lr As Boolean
    InitCommonCtrls = ""

    Dim ICC As tagINITCOMMONCONTROLSEX

    With ICC
        .dwSize = Len(ICC)
        .dwICC = ICC_STANDARD_CLASSES
    End With

    InitShell
    IsUserAnAdmin 'dummy to GetProcAddress
    lr = InitCommonControlsEx(ICC)
    If lr = 0 Or Err.Number <> 0 Then
      InitCommonControls ' 9x version
      If Err.Number <> 0 Then InitCommonCtrls = "Critical error: InitCommonCtrls failed!"
    End If
End Property
