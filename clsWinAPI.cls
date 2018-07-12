VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWinAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'Parameters Mapping values:
Const GHND = &H42
Const CF_TEXT = 1
Const MAXSIZE = 4096
Const GW_HWNDNEXT As Long = 2
Const HWND_TOPMOST = -1
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Const SWP_NOREDRAW = &H8
Const kCapital = 20
Const kNumlock = 144
Const HWND_NOTOPMOST As Long = -2
Const MOUSEEVENTF_MOVE As Long = &H1 ' mouse move
Const MOUSEEVENTF_LEFTDOWN As Long = &H2 ' left button down
Const MOUSEEVENTF_LEFTUP As Long = &H4 ' left button up
Const MOUSEEVENTF_RIGHTDOWN As Long = &H8 ' right button down
Const MOUSEEVENTF_RIGHTUP As Long = &H10 ' right button up
Const MOUSEEVENTF_MIDDLEDOWN As Long = &H20 ' middle button down
Const MOUSEEVENTF_MIDDLEUP As Long = &H40 ' middle button up
Const MOUSEEVENTF_ABSOLUTE As Long = &H8000 ' absolute move
Const MOUSEEVENTF_WHEEL As Long = &H800 ' wheel button rolled
Const sv_click As Long = &HF5&
Const SV_TEXT = &HC
Const SV_CHAR = &H102
Const SV_COMMAND = &H111
Const SV_KEYDOWN = &H290
Const CB_GETCURSEL = &H147
Const CB_SETCURSEL = &H14E
Const CB_SHWODROPDOWN = &H14F
Const CB_SELECTSTRING = &H14D
Const CBN_SELCHANGE = &H1
Const CHILDACTIVATE = &H22
Const TCM_SETCURFOCUS = &H1330&
Const O2S_WINDOW_MAX = 3
Const O2S_WINDOW_NORMAL = 1
Const O2S_WINDOW_MIN = 2
Const HWND_TOPMOST1 = -1
Const SWP_NOSIZE1 = &H1
Const SWP_NOMOVE1 = &H2
Const HWND_TOP1 = 0
Const HWND_BOTTOM1 = 1
Const HWND_NOTOPMOST1 = -2
Const EM_GETSEL = &HB0
Const EM_SETSEL = &HB1
Const EM_GETLINECOUNT = &HBA
Const EM_LINEINDEX = &HBB
Const EM_LINELENGTH = &HC1
Const EM_LINEFROMCHAR = &HC9
Const BM_GETCHECK = &HF0
Const BST_CHECKED = &H1
Const BM_CLICK = &HF5
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_LBUTTONDBLCLK = &H203
Const WM_MOUSEMOVE = &H200
Const WM_KEYDOWN = &H100
Const WM_KEYUP = &H101
Const MOUSEEVENTF_LEFTDOWN1 = &H2
Const MOUSEEVENTF_LEFTUP1 = &H4
Const WM_ACTIVATE = &H6
Const WM_ENABLE = &HA
Const WM_SHOWWINDOW = &H18
Const WM_CHILDACTIVATE = &H22
Const WM_SETCURSOR = &H20
Const WM_SETFOCUS = &H7
Const WM_KILLFOCUS = &H8
Const WM_CLOSE = &H16
Const WM_PASTE = &H302
Const WM_SIZE = &H5
Const WM_GETTEXTLENGTH = &HE
Const WM_GETTEXT = &HD
Const MK_CONTROL = &H8
Const MK_LBUTTON = &H1
Const MK_MBUTTON = &H10
Const MK_RBUTTON = &H2
Const MK_SHIFT = &H4
Const MK_XBUTTON1 = &H20
Const MK_XBUTTON2 = &H40
Const WM_NCHITTEST = &H84
Const VK_CONTROL = &H11
Const VK_0 = &H30
Const VK_1 = &H31
Const VK_2 = &H32
Const VK_3 = &H33
Const VK_4 = &H34
Const VK_5 = &H35
Const VK_6 = &H36
Const VK_7 = &H37
Const VK_8 = &H38
Const VK_9 = &H39
Const VK_A = &H41
Const VK_B = &H42
Const VK_C = &H43
Const VK_D = &H44
Const VK_E = &H45
Const VK_F = &H46
Const VK_G = &H47
Const VK_H = &H48
Const VK_I = &H49
Const VK_J = &H4A
Const VK_K = &H4B
Const VK_L = &H4C
Const VK_M = &H4D
Const VK_N = &H4E
Const VK_O = &H4F
Const VK_P = &H50
Const VK_Q = &H51
Const VK_R = &H52
Const VK_S = &H53
Const VK_T = &H54
Const VK_U = &H55
Const VK_V = &H56
Const VK_W = &H57
Const VK_X = &H58
Const VK_Y = &H59
Const VK_Z = &H5A
Const VK_DOWN = &H28 ' Down for Keyboard Event
Const VK_UP = &H26 'UP key
Const VK_LEFT = &H25 ' Left key
Const VK_RIGHT = &H27 ' Right key
Const VK_NUMLOCK = &H90 'Number key
Const KEYEVENTF_KEYUP = &H2 ' Keyboard constant for Keyboard Event
Const VK_ALT = &H12 ' Left Alt key for keybread event.
Const VK_APPS = &H5D
Const VK_RETURN = &HD 'Enter key
Const VK_SHIFT = &H10 'SHIFT Key
Const VK_TAB = &H9  'TAB Key
Const VK_SPACE = &H20 'SPACEBAR
Const VK_BACK = &H8 'BACK KEY
'Const GHND = &H42
'Const CF_TEXT = 1
'Const MAXSIZE = 4096


'********************************************
'**************KERNAL32 FUNCTIONS************
'********************************************
Private Declare Function OpenClipboard Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "User32" () As Long
Private Declare Function GetClipboardData Lib "User32" (ByVal wFormat As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags&, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Private Declare Function EmptyClipboard Lib "User32" () As Long
Private Declare Function SetClipboardData Lib "User32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GetWindow Lib "User32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
'===================================

Private Declare Function apiShowWindow Lib "User32" Alias "ShowWindow" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

'Get ChildWindows of a main window
Private Declare Function EnumChildWindows Lib "user32.dll" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

'Get Class Name of a specfic window
Private Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
'Access computer memory to find the byte values of a specflic window maybe apply to child window also
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    
Private Declare Function BringWindowToTop Lib "User32" (ByVal lngHWnd As Long) As Long
Private Declare Function IsWindowVisible Lib "User32" (ByVal hWnd As Long) As Boolean
Private Declare Function GetClientRect Lib "User32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetParent Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function ClientToScreen Lib "User32" (ByVal hWnd As Long, ByRef lpPoint As RECT) As Boolean
Private Declare Function MapWindowPoints Lib "User32" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As Any, ByVal cPoints As Long) As Long
Private Declare Function ChildWindowFromPoint Lib "User32" (hWnd As Long, ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function ChildWindowFromPointEx Lib "User32" (ByVal hWnd As Long, ByVal ptx As Long, ByVal pty As Long, ByVal un As Long) As Long
Private Declare Function SetForegroundWindow Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function GetKeyState Lib "User32" (ByVal nVirtKey As Long) As Integer
Private Declare Sub keybd_event Lib "User32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function GetWindowDC Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function GetDC Lib "User32" (ByVal hWnd As Long) As Long

' Access the GetCursorPos function in user32.dll
Private Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long
' Access the GetCursorPos function in user32.dll
Private Declare Function SetCursorPos Lib "User32" (ByVal x As Long, ByVal Y As Long) As Long
      
'We use GetPixel and GetWindowDC functions for checking right page has been loaded for SRI.
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long


      ' GetCursorPos requires a variable declared as a custom data type
      ' that will hold two integers, one for x value and one for y value
        Private Type POINTAPI
           x As Long
           Y As Long
        End Type
      

      
Private Declare Function SetWindowPos Lib "user32.dll" ( _
    ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal x As Long, _
    ByVal Y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) As Long

      
      
Private Declare Function apimouse_event Lib "User32" Alias "mouse_event" ( _
    ByVal dwFlags As Long, _
    ByVal dX As Long, _
    ByVal dY As Long, _
    ByVal cButtons As Long, _
    ByVal dwExtraInfo As Long) As Boolean


'Finded window
'Public Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000

Private Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As Any, _
    ByVal lpWindowName As String) As Long
    
Private Declare Function GetWindowTextW Lib "user32.dll" (ByVal hWnd As Long, ByVal lpString As Long, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowTextLengthW Lib "user32.dll" (ByVal hWnd As Long) As Long

Private Declare Function GetWindowTextLength Lib "User32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
    
'for making a window the topmost
'Private Declare Function SetWindowPos Lib "user32" (ByVal HWND As Long, _
'    ByVal hWndInsertAfter As Long, _
'    ByVal x As Long, _
'    ByVal y As Long, _
'    ByVal cx As Long, _
'    ByVal cy As Long, _
'    ByVal wFlags As Long) As Long


Private Declare Function FindWindowExA Lib "User32" _
(ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, _
ByVal lpsz2 As String) As Long


Private Declare Function EnableWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal fEnable As Boolean) As Long

Private Declare Function GetMenu Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function GetSubMenu Lib "User32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemID Lib "User32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuString Lib "User32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long




'getwindowrect
Private Declare Function GetWindowRect Lib "User32" (ByVal hWnd As Long, RECT As RECT) As Long
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, ByVal Source As Long, ByVal MessageId As Long, ByVal LanguageId As Long, ByVal Buffer As String, ByVal Size As Long, ByVal Arguments As Long) As Long

Private Declare Function GetKeyboardState Lib "User32" (pbKeyState As Byte) As Long
Private Declare Function SetKeyboardState Lib "User32" (kbArray As Byte) As Long
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long








Private Declare Function SendInput Lib "user32.dll" _
 (ByVal nInputs As Long, ByRef pInputs As Any, _
 ByVal cbSize As Long) As Long
 
Private Declare Function VkKeyScan Lib "User32" Alias "VkKeyScanA" _
 (ByVal cChar As Byte) As Integer

'posmouse
Private Declare Function PostMessage Lib "User32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Enum WM_NCHITTEST
    HTBORDER = 18
    HTBOTTOM = 15
    HTBOTTOMLEFT = 16
    HTBOTTOMRIGHT = 17
    HTCAPTION = 2
    HTCLIENT = 1
    HTCLOSE = 20
    HTERROR = -2
    HTGROWBOX = 4
    HTHELP = 21
    HTHSCROLL = 6
    HTLEFT = 10
    HTMENU = 5
    HTMAXBUTTON = 9
    HTMINBUTTON = 8
    HTNOWHERE = 0
    HTREDUCE = HTMINBUTTON
    HTRIGHT = 11
    HTSIZE = HTGROWBOX
    HTSYSMENU = 3
    HTTOP = 12
    HTTOPLEFT = 13
    HTTOPRIGHT = 14
    HTTRANSPARENT = -1
    HTVSCROLL = 7
    HTZOOM = HTMAXBUTTON
End Enum
'===============

Private Declare Function BlockInput Lib "User32" (ByVal fBlock As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'===============================================================================

     Private Declare Function MoveWindow Lib "User32" (ByVal hWnd As Long, _
         ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, _
         ByVal nHeight As Long, ByVal bRepaint As Long) As Long

     Private Declare Function GetDesktopWindow Lib "User32" () As Long

     Private Declare Function EnumWindows Lib "User32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

     Private Declare Function EnumThreadWindows Lib "User32" (ByVal dwThreadId _
         As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long

     Private Declare Function GetWindowThreadProcessId Lib "User32" _
         (ByVal hWnd As Long, lpdwProcessId As Long) As Long

     Private Declare Function GetWindowText Lib "User32" Alias "GetWindowTextA" _
         (ByVal hWnd As Long, ByVal lpString As String, _
         ByVal cch As Long) As Long

      Private TopCount As Integer     ' Number of Top level Windows
      Private ChildCount As Integer   ' Number of Child Windows
      Private ThreadCount As Integer  ' Number of Thread Windows

'     Public Function EnumWinProc(ByVal lhWnd As Long, ByVal lParam As Long) _
'         As Long
'         Dim retVal As Long, ProcessID As Long, ThreadID As Long
'         Dim WinClassBuf As String * 255, WinTitleBuf As String * 255
'         Dim winClass As String, winTitle As String
'
'         retVal = GetClassName(lhWnd, WinClassBuf, 255)
'         winClass = StripNulls(WinClassBuf)  ' remove extra Nulls & spaces
'         retVal = GetWindowText(lhWnd, WinTitleBuf, 255)
'         winTitle = StripNulls(WinTitleBuf)
'         TopCount = TopCount + 1
'         ' see the Windows Class and Title for each top level Window
'         Debug.Print "Top level Class = "; winClass; ", Title = "; winTitle
'         ' Usually either enumerate Child or Thread Windows, not both.
'         ' In this example, EnumThreadWindows may produce a very long list!
'         retVal = EnumChildWindows(lhWnd, AddressOf EnumChildProc, lParam)
'         ThreadID = GetWindowThreadProcessId(lhWnd, ProcessID)
'         retVal = EnumThreadWindows(ThreadID, AddressOf EnumThreadProc, _
'         lParam)
'         EnumWinProc = True
'      End Function

'      Public Function EnumChildProc(ByVal lhWnd As Long, ByVal lParam As Long) _
'         As Long
'         Dim retVal As Long
'         Dim WinClassBuf As String * 255, WinTitleBuf As String * 255
'         Dim winClass As String, winTitle As String
'         Dim WinRect As RECT
'         Dim WinWidth As Long, WinHeight As Long
'         Dim index2(200) As String
'         Dim i As Long
'         retVal = GetClassName(lhWnd, WinClassBuf, 255)
'         winClass = StripNulls(WinClassBuf)  ' remove extra Nulls & spaces
'         retVal = GetWindowText(lhWnd, WinTitleBuf, 255)
'         winTitle = StripNulls(WinTitleBuf)
'         ChildCount = ChildCount + 1
'
'         i = i + 1
'         ' see the Windows Class and Title for each Child Window enumerated
'         index2(i) = winTitle
'          Debug.Print " Child Class = " & winClass & ", Title = " & index2(i)
'
'         ' You can find any type of Window by searching for its WinClass
'         If winClass = "ThunderTextBox" Then    ' TextBox Window
'            retVal = GetWindowRect(lhWnd, WinRect)  ' get current size
'            WinWidth = WinRect.Right - WinRect.Left ' keep current width
'            WinHeight = (WinRect.Bottom - WinRect.Top) * 2 ' double height
'            retVal = MoveWindow(lhWnd, 0, 0, WinWidth, WinHeight, True)
'            EnumChildProc = False
'         Else
'            EnumChildProc = True
'         End If
'      End Function
'
'
'      Public Function EnumThreadProc(ByVal lhWnd As Long, ByVal lParam As Long) _
'         As Long
'         Dim retVal As Long
'         Dim WinClassBuf As String * 255, WinTitleBuf As String * 255
'         Dim winClass As String, winTitle As String
'
'         retVal = GetClassName(lhWnd, WinClassBuf, 255)
'         winClass = StripNulls(WinClassBuf)  ' remove extra Nulls & spaces
'         retVal = GetWindowText(lhWnd, WinTitleBuf, 255)
'         winTitle = StripNulls(WinTitleBuf)
'         ThreadCount = ThreadCount + 1
'         ' see the Windows Class and Title for top level Window
'         Debug.Print "Thread Window Class = "; winClass; ", Title = "; _
'         winTitle
'         EnumThreadProc = True
'      End Function
      

      Public Function StripNulls(OriginalStr As String) As String
         ' This removes the extra Nulls so String comparisons will work
         If (InStr(OriginalStr, Chr(0)) > 0) Then
            OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
         End If
         StripNulls = OriginalStr
      End Function

Public Function ForegroundWindow(ByVal hWnd As Long)
    SetForegroundWindow hWnd
    Sleep 50
End Function

Public Function getDesktopHwnd() As Long
    getDesktopHwnd = GetDesktopWindow
End Function

Public Function GetMenuText(ByVal hWnd As Long, ByVal wIDItem As Long) As String
    Dim C As Long
    Dim S As String
    Dim hMenu As Long
    hMenu = GetMenu(hWnd)
    ' Get the text size
    C = GetMenuString(hMenu, wIDItem, vbNullString, 0, 0)
    ' Allocate the string
    S = String(C + 1, 0)
    ' Get the text
    C = GetMenuString(hMenu, wIDItem, S, C + 1, 0)
    GetMenuText = Trim(S)
End Function

Public Function GetObjectText(hWnd As Long) As String
Dim i As Long
Dim S As String

i = SendMessage(hWnd, WM_GETTEXTLENGTH, ByVal 0&, ByVal 0&)

S = Space$(i + 1)
SendMessage hWnd, WM_GETTEXT, ByVal i + 1, ByVal S

GetObjectText = CleanString(S)

End Function

Public Function getNameFromHwnd(hWnd As Long) As String
    Dim title As String * 255
    Dim tLen As Long
    tLen = GetWindowTextLength(hWnd)
    GetWindowText hWnd, title, 255
    getNameFromHwnd = Left(title, tLen)
End Function

Function ToggleCapsLock(varCase As Long)
    Dim Res As Long
    Dim KBState(0 To 255) As Byte
    Res = GetKeyboardState(KBState(0))
    KBState(&H14) = varCase
    Res = SetKeyboardState(KBState(0))
End Function

Function CapsLockStatus() As Boolean
    Dim Res As Long
    Dim KBState(0 To 255) As Byte
        Sleep 250
    Res = GetKeyboardState(KBState(0))
        Sleep 250
    CapsLockStatus = KBState(&H14) And 1
        Sleep 250
End Function


Public Function CapsLock() As Boolean
    Sleep 250
    CapsLock = KeyState(kCapital)
End Function

Public Function NumLock() As Boolean
    NumLock = KeyState(kNumlock)
End Function

Private Function KeyState(lKey As Long) As Boolean
    Sleep 250
    KeyState = CBool(GetKeyState(lKey))
End Function

Function ProcIDFromWnd(ByVal hWnd As Long) As Long
   Dim idProc As Long

   ' Get PID for this HWnd
   GetWindowThreadProcessId hWnd, idProc
   ProcIDFromWnd = idProc
End Function

Function GetWinHandle(hInstance As Long) As Long
   Dim tempHwnd As Long
     
   ' Grab the first window handle that Windows finds:
   tempHwnd = FindWindow(vbNullString, vbNullString)
  
   ' Loop until you find a match or there are no more window handles:
   Do Until tempHwnd = 0
      ' Check if no parent for this window
      If GetParent(tempHwnd) = 0 Then
         ' Check for PID match
         If hInstance = ProcIDFromWnd(tempHwnd) Then
            ' Return found handle
            GetWinHandle = tempHwnd
            ' Exit search loop
            Exit Do
         End If
      End If
  
      ' Get the next window handle
      tempHwnd = GetWindow(tempHwnd, GW_HWNDNEXT)
   Loop
End Function

Function GetWindowsRect(hHnd As Long) As Variant
    Dim arect As RECT, result(2) As Integer
    GetWindowRect hHnd, arect
    If hHnd = 0 Then GetWindowsRect = Array(0, 0): Exit Function
    GetWindowsRect = Array(arect.Left, arect.Top)
End Function

Sub SetaWindowPos(hHnd As Long, x As Long, Y As Long, cx As Long, cy As Long, wFlags As Long)
    SetWindowPos hHnd, 0&, x, Y, cx, cy, wFlags
End Sub

Function getObjectHwnd(className As String, objectName As String) As Long
    getObjectHwnd = FindWindow(className, objectName)
End Function

Function getChildObjectHwnd(PrimaryHwnd As Long, SubHwnd As Long, className As String, objectName As String) As Long
    getChildObjectHwnd = FindWindowExA(PrimaryHwnd, SubHwnd, className, objectName)
End Function

Function ActivateMenuObject(hWnd As Long, MenuStr As String, mIndex As Integer, SendOrPost As String) As Long
    Dim MenuIndex As Integer, i As Long, MenuText As String
    On Error Resume Next
    For i = mIndex To 1000
        MenuText = CleanString(GetMenuText(hWnd, i))
        MenuText = Replace(MenuText, "&", "")
        If MenuText = MenuStr Then
            Select Case SendOrPost
                Case "sendmsg"
                    ActivateMenuObject = SendMessage(hWnd, SV_COMMAND, i, ByVal 0&)
                Case "postmsg"
                    ActivateMenuObject = PostMessage(hWnd, SV_COMMAND, i, ByVal 0&)
            End Select
            Exit Function
        End If
    Next
End Function

Sub Sleeping(mSec As Long)
    Sleep mSec
End Sub

Sub BkgrndSingleButtonClick(hWnd As Long, iTimes As Integer, x As Integer, Y As Integer)
    Dim i As Integer
    For i = 1 To iTimes
        Sleep 250
        Call SendMessage(hWnd, WM_LBUTTONDOWN, 0&, ByVal MakeDWord(x, Y))
        Call SendMessage(hWnd, WM_LBUTTONUP, 0&, ByVal MakeDWord(x, Y))
        Sleep 250
    Next
End Sub

Sub SendmsgClick(BtnHwnd As Long)
    SendMessage BtnHwnd, sv_click, 0&, 0&
End Sub

Sub PostmsgClick(BtnHwnd As Long)
    PostMessage BtnHwnd, sv_click, 0&, 0&
End Sub

Sub CloseObjWindow(hWnd As Long)
    SendMessage hWnd, &H10, 0&, 0&
End Sub

Sub ActivateATab(hWnd As Long, index As Long)
    SendMessage hWnd, TCM_SETCURFOCUS, index, 0
End Sub

Sub BkgrndDoubleButtonClick(hWnd As Long, x As Integer, Y As Integer)
    Sleep 25
    Call SendMessage(hWnd, WM_LBUTTONDBLCLK, 0&, ByVal MakeDWord(x, Y))
    Sleep 25
End Sub

Sub TickCheckBox(CboxHwnd As Long, CboxSwitch As Boolean)
    Dim Retval As Integer
    Retval = SendMessage(CboxHwnd, BM_GETCHECK, 0&, 0&)
    Select Case CboxSwitch
        Case True
            If Retval = 0 Then
                SendMessage CboxHwnd, BM_CLICK, 0&, 0&
            End If
        Case False
            If Retval = 1 Then
                SendMessage CboxHwnd, BM_CLICK, 0&, 0&
            End If
    End Select
End Sub

Sub SendText(hWnd As Long, aText As String)
    SendMessage hWnd, SV_TEXT, 0, ByVal CStr(aText)
End Sub

Sub ObjSetFocus(hWnd As Long)
    SendMessage hWnd, WM_SETFOCUS, 0, ByVal 0&
End Sub

Sub ObjKillFocus(hWnd As Long)
    SendMessage hWnd, WM_KILLFOCUS, 0, ByVal 0&
End Sub

Sub VKeyDown(hWnd As Long)
    SendMessage hWnd, WM_KEYDOWN, VK_DOWN, ByVal 0&
    SendMessage hWnd, WM_KEYUP, VK_DOWN, ByVal 0&
End Sub

Sub VKeyDownPostWay(hWnd As Long)
    PostMessage hWnd, WM_KEYDOWN, VK_DOWN, ByVal 0&
    PostMessage hWnd, WM_KEYUP, VK_DOWN, ByVal 0&
End Sub

Sub VKeyUp(hWnd As Long)
    SendMessage hWnd, WM_KEYDOWN, VK_UP, ByVal 0&
    SendMessage hWnd, WM_KEYUP, VK_UP, ByVal 0&
End Sub

Sub VKeyBack(hWnd As Long)
    SendMessage hWnd, WM_KEYDOWN, VK_BACK, ByVal 0&
    SendMessage hWnd, WM_KEYUP, VK_UP, ByVal 0&
End Sub

Sub VKeyTab(hWnd As Long)
    SendMessage hWnd, WM_KEYDOWN, VK_TAB, ByVal 0&
    SendMessage hWnd, WM_KEYUP, VK_TAB, ByVal 0&
End Sub

Sub VKeyReturn(hWnd As Long)
    SendMessage hWnd, WM_KEYDOWN, VK_RETURN, ByVal 0&
    SendMessage hWnd, WM_KEYUP, VK_RETURN, ByVal 0&
End Sub

Sub VKeyLft(hWnd As Long)
    SendMessage hWnd, WM_KEYDOWN, VK_LEFT, ByVal 0&
    SendMessage hWnd, WM_KEYUP, VK_LEFT, ByVal 0&
End Sub

Sub VKeyRit(hWnd As Long)
    SendMessage hWnd, WM_KEYDOWN, VK_RIGHT, ByVal 0&
    SendMessage hWnd, WM_KEYUP, VK_RIGHT, ByVal 0&
End Sub

Sub VkeyCtrlC(ByVal hWnd As Long)
    Sleep 250
    keybd_event &H11, 0, 0, 0
    Sleep 50
    PostMessage hWnd, WM_KEYDOWN, VK_C, ByVal 0&
    Sleep 50
    keybd_event &H11, 0, KEYEVENTF_KEYUP, 0
    Sleep 300
End Sub

Sub HitKeyReturn()
    Sleep 50
    keybd_event VK_RETURN, 0, 0, 0
    Sleep 20
    keybd_event VK_RETURN, 0, KEYEVENTF_KEYUP, 0
    Sleep 50
End Sub

Private Function MakeDWord(LoWord As Integer, HiWord As Integer) As Long
    MakeDWord = (HiWord * &H10000) Or (LoWord And &HFFFF&)
End Function

Function ChkboxStatus(ByVal btnHandle As Long) As Boolean
    Dim ret As Integer
    ret = SendMessage(btnHandle, BM_GETCHECK, 0, 0)
    If ret = 1 Then
        ChkboxStatus = True
    Else
        ChkboxStatus = False
    End If
End Function

Sub ChkboxClick(ByVal btnHandle As Long)
    SendMessage btnHandle, BM_CLICK, 0&, 0&
End Sub

Function CleanString(StrIn As String) As String
    ' "Cleans" a string by removing embedded control (non-printable)
    ' characters, including carriage returns and linefeeds.
    ' Does not remove special characters like symbols, international
    ' characters, etc. This function runs recursively, each call
    ' removing one embedded character
    Dim iCh As Integer
    CleanString = StrIn
    For iCh = 1 To Len(StrIn)
        If Asc(Mid(StrIn, iCh, 1)) < 32 Then
            'remove special character
            CleanString = Left(StrIn, iCh - 1) & CleanString(Mid(StrIn, iCh + 1))
            Exit Function
        End If
    Next iCh
End Function

Sub MaxmizeWindow(hWnd As Long)
    apiShowWindow hWnd, O2S_WINDOW_MAX
End Sub

Function IsWinVisible(hWnd As Long) As Boolean
    If IsWindowVisible(hWnd) Then
        IsWinVisible = True
    Else
        IsWinVisible = False
    End If
End Function

Sub SetClipboard(sUniText As String)
    Dim iStrPtr As Long
    Dim iLen As Long
    Dim iLock As Long
    Const GMEM_MOVEABLE As Long = &H2
    Const GMEM_ZEROINIT As Long = &H40
    Const CF_UNICODETEXT As Long = &HD
    OpenClipboard 0&
    EmptyClipboard
    iLen = LenB(sUniText) + 2&
    iStrPtr = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, iLen)
    iLock = GlobalLock(iStrPtr)
    lstrcpy iLock, StrPtr(sUniText)
    GlobalUnlock iStrPtr
    SetClipboardData CF_UNICODETEXT, iStrPtr
    CloseClipboard
End Sub

Sub ClearClipboard()
    EmptyClipboard
End Sub

Function ClipBoard_GetData() As String
    Dim hClipMemory As Long
    Dim lpClipMemory As Long
    Dim MyString As String
    Dim Retval As Long
    
    If OpenClipboard(0&) = 0 Then
        ClipBoard_GetData = "ERR_CB_LOCKED"
        Exit Function
    End If

    ' Obtain the handle to the global memory
    ' block that is referencing the text.
    hClipMemory = GetClipboardData(CF_TEXT)
    If IsNull(hClipMemory) Then
       GoTo OutOfHere
    End If

    ' Lock Clipboard memory so we can reference
    ' the actual data string.
    lpClipMemory = GlobalLock(hClipMemory)

    If Not IsNull(lpClipMemory) Then
       MyString = Space$(MAXSIZE)
       Retval = lstrcpy(MyString, lpClipMemory)
       Retval = GlobalUnlock(hClipMemory)
    
       ' Peel off the null terminating character.
       MyString = Mid(MyString, 1, InStr(1, MyString, Chr$(0), 0) - 1)
    Else
       ClipBoard_GetData = "ERR_CB_LOCKED"
    End If

OutOfHere:

    Retval = CloseClipboard()
    If InStr(ClipBoard_GetData, "ERR") < 1 Then
        ClipBoard_GetData = MyString
    End If
End Function

Function BlockInputs(ByVal TrueorFalse As Boolean)
    BlockInput TrueorFalse
End Function

Function GetWinDC(ByVal hWnd As Long) As Long
    GetWinDC = GetWindowDC(hWnd)
End Function

Function GetaDC(ByVal hWnd As Long) As Long
    GetaDC = GetDC(hWnd)
End Function

Function GetPix(ByVal LdcValue As Long, x As Long, Y As Long) As Long
    GetPix = GetPixel(LdcValue, x, Y)
End Function

Sub KeyTypeIn(ByVal hWnd As Long, key As String, x As Integer, Y As Integer)
    Sleep 5
    PostMessage hWnd, WM_KEYDOWN, CharReader(key), ByVal MakeDWord(x, Y)
    Sleep 5
    PostMessage hWnd, WM_KEYUP, 0&, ByVal MakeDWord(x, Y)
    Sleep 5
End Sub

Private Function CharReader(ByVal Char As String) As Long
    Select Case Char
        Case "a"
            CharReader = &H41
        Case "b"
            CharReader = &H42
        Case "c"
            CharReader = &H43
        Case "d"
            CharReader = &H44
        Case "e"
            CharReader = &H45
        Case "f"
            CharReader = &H46
        Case "g"
            CharReader = &H47
        Case "h"
            CharReader = &H48
        Case "i"
            CharReader = &H49
        Case "j"
            CharReader = &H4A
        Case "k"
            CharReader = &H4B
        Case "l"
            CharReader = &H4C
        Case "m"
            CharReader = &H4D
        Case "n"
            CharReader = &H4E
        Case "o"
            CharReader = &H4F
        Case "p"
            CharReader = &H50
        Case "q"
            CharReader = &H51
        Case "r"
            CharReader = &H52
        Case "s"
            CharReader = &H53
        Case "t"
            CharReader = &H54
        Case "u"
            CharReader = &H55
        Case "v"
            CharReader = &H56
        Case "w"
            CharReader = &H57
        Case "x"
            CharReader = &H58
        Case "y"
            CharReader = &H59
        Case "z"
            CharReader = &H5A
        Case "0"
            CharReader = &H30
        Case "1"
            CharReader = &H31
        Case "2"
            CharReader = &H32
        Case "3"
            CharReader = &H33
        Case "4"
            CharReader = &H34
        Case "5"
            CharReader = &H35
        Case "6"
            CharReader = &H36
        Case "7"
            CharReader = &H37
        Case "8"
            CharReader = &H38
        Case "9"
            CharReader = &H39
        Case "Space"
            CharReader = &H20
        Case " "
            CharReader = &H20
        Case "/"
            CharReader = &HBF
        Case ":"
            CharReader = &HBA
    End Select
End Function
