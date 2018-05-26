Attribute VB_Name = "modFileEx"

    Option Explicit
     
'    'common to both methods
'    Public Type BROWSEINFO
'     hOwner As Long
'     pidlRoot As Long
'     pszDisplayName As String
'     lpszTitle As String
'     ulFlags As Long
'     lpfn As Long
'     lParam As Long
'     iImage As Long
'    End Type
'
'    Public Declare Function SHBrowseForFolder Lib _
'     "shell32.dll" Alias "SHBrowseForFolderA" _
'     (lpBrowseInfo As BROWSEINFO) As Long
'
'    Public Declare Function SHGetPathFromIDList Lib _
'     "shell32.dll" Alias "SHGetPathFromIDListA" _
'     (ByVal pidl As Long, _
'     ByVal pszPath As String) As Long
'
'    Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
'
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
     (ByVal hWnd As Long, _
     ByVal wMsg As Long, _
     ByVal wParam As Long, _
     lParam As Any) As Long
'
'    Public Declare Sub MoveMemory Lib "kernel32" _
'     Alias "RtlMoveMemory" _
'     (pDest As Any, _
'     pSource As Any, _
'     ByVal dwLength As Long)
'
'    Public Const MAX_PATH = 260
    Private Const WM_USER = &H400
    Private Const BFFM_INITIALIZED = 1
'
'    'Constants ending in 'A' are for Win95 ANSI
'    'calls; those ending in 'W' are the wide Unicode
'    'calls for NT.
'
'    'Sets the status text to the null-terminated
'    'string specified by the lParam parameter.
'    'wParam is ignored and should be set to 0.
'    Public Const BFFM_SETSTATUSTEXTA As Long = (WM_USER + 100)
'    Public Const BFFM_SETSTATUSTEXTW As Long = (WM_USER + 104)
'
'    'If the lParam parameter is non-zero, enables the
'    'OK button, or disables it if lParam is zero.
'    '(docs erroneously said wParam!)
'    'wParam is ignored and should be set to 0.
'    Public Const BFFM_ENABLEOK As Long = (WM_USER + 101)
'
'    'Selects the specified folder. If the wParam
'    'parameter is FALSE, the lParam parameter is the
'    'PIDL of the folder to select , or it is the path
'    'of the folder if wParam is the C value TRUE (or 1).
'    'Note that after this message is sent, the browse
'    'dialog receives a subsequent BFFM_SELECTIONCHANGED
'    'message.
    Private Const BFFM_SETSELECTIONA As Long = (&H400 + 102)
'    Public Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)
'
'
'    'specific to the PIDL method
'    'Undocumented call for the example. IShellFolder's
'    'ParseDisplayName member function should be used instead.
'    Public Declare Function SHSimpleIDListFromPath Lib _
'     "shell32" Alias "#162" _
'     (ByVal szPath As String) As Long
'
'
'    'specific to the STRING method
'    Public Declare Function LocalAlloc Lib "kernel32" _
'     (ByVal uFlags As Long, _
'     ByVal uBytes As Long) As Long
'
'    Public Declare Function LocalFree Lib "kernel32" _
'     (ByVal hMem As Long) As Long
'
    Private Declare Function lstrcpyA Lib "kernel32" (lpString1 As Any, lpString2 As Any) As Long
'
    Private Declare Function lstrlenA Lib "kernel32" (lpString As Any) As Long
'
'    Public Const LMEM_FIXED = &H0
'    Public Const LMEM_ZEROINIT = &H40
'    Public Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)
     
Public Function BrowseCallbackProcStr(ByVal hWnd As Long, _
     ByVal uMsg As Long, _
     ByVal lParam As Long, _
     ByVal lpData As Long) As Long
     
     Select Case uMsg
        Case BFFM_INITIALIZED
        Call SendMessage(hWnd, BFFM_SETSELECTIONA, True, ByVal StrFromPtrA(lpData))
     Case Else:
     End Select
     
End Function
     
     
Public Function BrowseCallbackProc(ByVal hWnd As Long, _
     ByVal uMsg As Long, _
     ByVal lParam As Long, _
     ByVal lpData As Long) As Long
     Select Case uMsg
        Case BFFM_INITIALIZED
        
        Call SendMessage(hWnd, BFFM_SETSELECTIONA, False, ByVal lpData)
        
        Case Else:
     
     End Select
     
    End Function

Public Function FARPROC(pfn As Long) As Long
    FARPROC = pfn
End Function
     

Public Function StrFromPtrA(lpszA As Long) As String
    Dim sRtn As String
    sRtn = String$(lstrlenA(ByVal lpszA), 0)
    Call lstrcpyA(ByVal sRtn, ByVal lpszA)
    StrFromPtrA = sRtn
End Function
