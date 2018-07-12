VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsICMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'THESE FILES ARE TO CONTROL ICMS AND HAVE BE EDITED FOR THIS PURPOSE IF YOU NEED MORE INFO PLEASE TALK TO JONO FRIESE/JEFF ZHANG
' DO NOT EDIT ANYTHING IN HERE!



Option Explicit
'C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL"

Private Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetWindowWord Lib "User32" (ByVal hWnd As Long, ByVal nIndex As Long) As Integer
Private Declare Function WD_Convert Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal hInstance As Long, ByVal ConvertType As Integer, ROWCOL As ROWCOL, ByVal ShortName As String) As Integer
Private Declare Function WD_AdviseHostUpdate Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal hInstance As Long, ByVal hInstanceMsg As Integer, ByVal msg As Integer) As Integer
Private Declare Function WD_ConnectPS Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal hInstance As Long, ByVal ShortName As String) As Integer
Private Declare Function WD_CopyFieldToString Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal hInstance As Long, ByVal Position As Integer, ByVal Buffer As String, ByVal Length As Integer) As Integer
Private Declare Function WD_CopyOIA Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal hInstance As Long, ByVal OIAData As String, ByVal Length As Integer) As Integer
Private Declare Function WD_CopyPS Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal hInstance As Long, ByVal Buffer As String, ByVal Length As Integer) As Integer
Private Declare Function WD_CopyPSToString Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal hInstance As Long, ByVal Position As Integer, ByVal Buffer As String, ByVal Length As Integer) As Integer
Private Declare Function WD_CopyStringToField Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal hInstance As Long, ByVal Position As Integer, ByVal Buffer As String) As Integer
Private Declare Function WD_CopyStringToPS Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal hInstance As Long, ByVal Position As Integer, ByVal Buffer As String, ByVal Length As Integer) As Integer
Private Declare Function WD_DeletePS Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal hInstance As Long, ByVal ShortName As String) As Integer
Private Declare Function WD_DisconnectPS Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal hInstance As Long) As Integer
Private Declare Function WD_DisplayCursor Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal hInstance As Long, ByVal Position As Integer, ByVal ShortName As String) As Integer
Private Declare Function WD_DisplayPS Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal hInstance As Long, ByVal ShortName As String) As Integer
Private Declare Function WD_FindFieldLength Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal hInstance As Long, Length As Integer, ByVal Position As Integer, ByVal FindData As String) As Integer
Private Declare Function WD_FindFieldPosition Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal hInstance As Long, Location As Integer, ByVal Position As Integer, ByVal FindData As String) As Integer
Private Declare Function WD_GetKey Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal hInstance As Long, ByVal GetKeyData As String) As Integer
Private Declare Function WD_GetSessionHWND Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal hInstance As Long) As Integer
Private Declare Function WD_Pause Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal hInstance As Long, ByVal Length As Integer) As Integer
Private Declare Function WD_PostInterceptStatus Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal hInstance As Long, ByVal ShortName As String) As Integer
Private Declare Function WD_QueryCursorLocation Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal hInstance As Long, Location As Integer) As Integer
Private Declare Function WD_QueryHostUpdate Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal hInstance As Long, ByVal ShortName As String) As Integer
Private Declare Function WD_QuerySessionStatus Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal hInstance As Long, ByVal SessionData As String) As Integer
Private Declare Function WD_QuerySystem Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal hInstance As Long, ByVal SystemData As String) As Integer
Private Declare Function WD_ReceiveFile Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal hInstance As Long, ByVal TransferData As String) As Integer
Private Declare Function WD_Release Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal hInstance As Long) As Integer
Private Declare Function WD_Reserve Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal hInstance As Long) As Integer
Private Declare Function WD_ResetSystem Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal hInstance As Long) As Integer
Private Declare Function WD_RunMacro Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal hInstance As Long, ByVal Buffer As String) As Integer
Private Declare Function WD_RunProfile Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal Profile As String, ByVal SW_Value As Integer) As Integer
Private Declare Function WD_SearchField Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal hInstance As Long, Location As Integer, ByVal Position As Integer, ByVal Buffer As String) As Integer
Private Declare Function WD_SearchPS Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal hInstance As Long, Location As Integer, ByVal Position As Integer, ByVal SearchData As String) As Integer
Private Declare Function WD_SendFile Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal hInstance As Long, ByVal TransferData As String) As Integer
Private Declare Function WD_SendKey Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal hInstance As Long, ByVal KeyData As String) As Integer
Private Declare Function WD_SetCursor Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal hInstance As Long, ByVal Position As Integer) As Integer
Private Declare Function WD_SetSessionParameters Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal hInstance As Long, ByVal SSPData As Long) As Integer
Private Declare Function WD_ShowSession Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal hInstance As Long, ByVal ShowWindow As Integer) As Integer
Private Declare Function WD_StartHostNotification Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal hInstance As Long, ByVal NotifyData As String) As Integer
Private Declare Function WD_StartKeystrokeIntercept Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" Alias "WD_StartKSIntercept" (ByVal hInstance As Long, ByVal SKIData As String) As Integer
Private Declare Function WD_StopHostNotification Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal hInstance As Long, ByVal ShortName As String) As Integer
Private Declare Function WD_StopKeystrokeIntercept Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" Alias "WD_StopKSIntercept" (ByVal hInstance As Long, ByVal ShortName As String) As Integer
Private Declare Function WD_Wait Lib "C:\Program Files (x86)\Micro Focus\RUMBA\System\EHLAPI32.DLL" (ByVal hInstance As Long) As Integer
Private Declare Function GetWindowText Lib "User32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

' Return codes from WD EHLLAPI functions
Const APIOK = 0                  ' Successful
Const APINOTCONNECTED = 1        ' Not Connected To Presentation Space
Const APIPARAMETERERROR = 2      ' Parameter Error/Invalid Function
Const APIFTXCOMPLETE = 3         ' File Transfer Complete
Const APIFTXSEGMENTED = 4        ' File Transfer Complete / segmented
Const APIPSBUSY = 4              ' Presentation Space is Busy
Const APIINHIBITED = 5           ' Inhibited/Keyboard Locked
Const APITRUNCATED = 6           ' Data Truncated
Const APIPOSITIONERROR = 7       ' Invalid Presentation Space Position
Const APINOTAVAILABLE = 8        ' Unavailable Operation
Const APISYSERROR = 9            ' System Error
Const APINOTSUPPORTED = 10       ' Function not supported
Const APIUNAVAILABLE = 11        ' Resource is unavailable
Const APIUNDEFINEDKEY = 20       ' Undefined key combination
Const APIOIAUPDATE = 21          ' OIA updated
Const APIPSUPDATE = 22           ' PS updated
Const APIBOTHUPDATE = 23         ' Both PS And OIA Updated
Const APINOFIELD = 24            ' No Such Field Found
Const APINOKEYSTROKES = 25       ' No Keystrokes are available
Const APIPSCHANGED = 26          ' PS or OIA changed
Const APIINVALIDTYPE = 30        ' Invalid Cursor Type
Const APIKEYOVERFLOW = 31        ' Keystroke overflow
Const APIINVALIDPSID = 9998      ' Invalid Presentation Space ID
Const APIINVALIDRC = 9999        ' Invalid Row or Column Code
Const APIFILENOTFOUND = 302      ' File Not Found
Const APIACCESSDENIED = 305      ' Access Denied
Const APIMEMORY = 308            ' Insufficient Memory

' RumbaSession (ShowWindow) constants
Const SW_HIDE = 0
Const SW_SHOWNORMAL = 1
Const SW_SHOWMINIMIZED = 2
Const SW_SHOWMAXIMIZED = 3
Const SW_SHOWNOACTIVATE = 4
Const SW_SHOW = 5
Const SW_MINIMIZE = 6
Const SW_SHOWMINNOACTIVE = 7
Const SW_SHOWNA = 8
Const SW_RESTORE = 9
Const SW_DELETEPS = 10
Const SW_DEFAULT = 11

'  Window field offsets for GetWindowWord() (same as in Windows.h)
Const GWW_HINSTANCE = (-6)
Const GWW_HWNDPARENT = (-8)
Const GWW_ID = (-12)

' WD_Convert constants
Const CONVERT_POSITION = 80      ' (P) Convert from Position to Row/Col
Const CONVERT_ROW = 82           ' (R) Convert from Row/Col to Position

' WD_FindFieldLength/Position constants
Const FF_CURRENT = "T "
Const FF_NEXT = "N "
Const FF_NEXTPROTECTED = "NP"
Const FF_NEXTUNPROTECTED = "NU"
Const FF_PREVIOUS = "P "
Const FF_PREVIOUSPROTECTED = "PP"
Const FF_PREVIOUSUNPROTECTED = "PU"

' WD_SetSessionParameters definitions                                           */
Const SSP_NEWRET = &H1
Const SSP_OLDRET = &H2
Const SSP_ATTRB = &H4
Const SSP_NOATTRB = &H8
Const SSP_NWAIT = &H10
Const SSP_LWAIT = &H20
Const SSP_TWAIT = &H40
Const SSP_EAB = &H80
Const SSP_NOEAB = &H100
Const SSP_AUTORESET = &H200
Const SSP_NORESET = &H400
Const SSP_SRCHALL = &H1000
Const SSP_SRCHFROM = &H2000
Const SSP_SRCHFRWD = &H4000
Const SSP_SRCHBKWD = &H8000
Const SSP_FPAUSE = &H10000
Const SSP_IPAUSE = &H20000
Const SSP_XLATE = &H40000
Const SSP_NOXLATE = &H80000
'------------------HLLAPI Constants-----------------------
Const RumbaAt                As String = "@@"
Const RumbaAlt               As String = "@A"
Const RumbaAltCursor         As String = "@e@1*"
Const RumbaAttention         As String = "@A@Q"
Const RumbaBackspace         As String = "@<"
Const RumbaBacktab           As String = "@B"
Const RumbaClear             As String = "@C"
Const RumbaCmd               As String = "@A@Y"
Const RumbaCursorDown        As String = "@V"
Const RumbaCursorLeft        As String = "@L"
Const RumbaCursorRight       As String = "@Z"
Const RumbaCursorSelect      As String = "@e@1*"
Const RumbaCursorUp          As String = "@U"
Const RumbaDelete            As String = "@D"
Const RumbaDup               As String = "@S@"
Const RumbaEndKey            As String = "@q"
Const RumbaEnterKey          As String = "@E"
Const RumbaEraseEOF          As String = "@F"
Const RumbaEraseEOL          As String = "@e@d"
Const RumbaEraseInput        As String = "@A@F"
Const RumbaFieldExit         As String = "@A@E"
Const RumbaFieldMark         As String = "@S@y"
Const RumbaFieldMinus        As String = "@A@-"
Const RumbaFieldPlus         As String = "@A@+"
Const RumbaHelp              As String = "@H"
Const RumbaHexadecimal       As String = "@A@"
Const RumbaHome              As String = "@0"
Const RumbaInsert            As String = "@I"
Const RumbaInsertToggle      As String = "@A@I"
Const RumbaLocalHelp         As String = "@e@I"
Const RumbaLocalPrint        As String = "@P"
Const RumbaLocalHelpCursor   As String = "@e@rn*"
Const RumbaLeftTab           As String = "@B"
Const RumbaNewLine           As String = "@e@n*"
Const RumbaPageUp            As String = "@u"
Const RumbaPageDown          As String = "@v"
Const RumbaPrintPC           As String = "@A@t"
Const RumbaPrintScreen       As String = "@A@T"
Const RumbaRecordBackspace   As String = "@A@<"
Const RumbaReset             As String = "@R"
Const RumbaShift             As String = "@S"
Const RumbaSysRequest        As String = "@A@H"
Const RumbaTabKey            As String = "@T"
Const RumbaTest              As String = "@A@C"
Const RumbaUSM               As String = "@e@v"
Const RumbaAPLCUTONOFF       As String = "@e@2"
Const RumbaPA1               As String = "@x"
Const RumbaPA2               As String = "@y"
Const RumbaPA3               As String = "@z"
Const RumbaPA4               As String = "@+"
Const RumbaPA5               As String = "@%"
Const RumbaPA6               As String = "@&"
Const RumbaPA7               As String = "@'"
Const RumbaPA8               As String = "@("
Const RumbaPA9               As String = "@)"
Const RumbaPA10              As String = "@*"
Const RumbaCmd1              As String = "@1"
Const RumbaCmd2              As String = "@2"
Const RumbaCmd3              As String = "@3"
Const RumbaCmd4              As String = "@4"
Const RumbaCmd5              As String = "@5"
Const RumbaCmd6              As String = "@6"
Const RumbaCmd7              As String = "@7"
Const RumbaCmd8              As String = "@8"
Const RumbaCmd9              As String = "@9"
Const RumbaCmd10             As String = "@a"
Const RumbaCmd11             As String = "@b"
Const RumbaCmd12             As String = "@c"
Const RumbaCmd13             As String = "@d"
Const RumbaCmd14             As String = "@e"
Const RumbaCmd15             As String = "@f"
Const RumbaCmd16             As String = "@g"
Const RumbaCmd17             As String = "@h"
Const RumbaCmd18             As String = "@i"
Const RumbaCmd19             As String = "@j"
Const RumbaCmd20             As String = "@k"
Const RumbaCmd21             As String = "@1"
Const RumbaCmd22             As String = "@m"
Const RumbaCmd23             As String = "@n"
Const RumbaCmd24             As String = "@c"

Const TIME_OUT_SECS As Integer = 20

Private Type ROWCOL
  row As Integer
  col As Integer
End Type


Dim nRetValue  As Long
Dim endTime    As Date
Dim msgResult  As Variant
Dim errMessage As String
Dim lngPid As Long

Private Sub OIAWait()
  'This procedure waits for ICMS to finish processing what ever it's doing then _
   'it will continue execution
  Dim strOIA As String * 103
  Do
    Sleep 25
    nRetValue = WD_CopyOIA(lngPid, strOIA, 103)
  Loop Until InStr(strOIA, "II") = 0
End Sub


Sub activateICMSwindow()
    Dim hWndRUMBA As Integer
    Dim sGetWindowTitle As String
    sGetWindowTitle = String$(50, 0)
    hWndRUMBA = WD_GetSessionHWND(lngPid)
    nRetValue = GetWindowText(hWndRUMBA, sGetWindowTitle, 50)   'Windows API function
    AppActivate sGetWindowTitle, True
End Sub



Function ConnectPS() As Boolean
  Dim sNotifyData As String
  lngPid = GetCurrentProcessId()
  'MsgBox lngPid
  nRetValue = WD_ConnectPS(lngPid, "A")
  If nRetValue <> 0 Then
    ConnectPS = False
  Else
    ConnectPS = True
  
  End If
End Function

Private Sub Class_Terminate()
  'Purpose : Discconect connection to Rumba session
  WD_DisconnectPS (lngPid)
End Sub

Sub ClearAlerts()
  If scrRead(7, 27, 48) = "Select National Number" Then
    EnterKey
  End If
  If scrRead(11, 38, 47) = "Alert Code" Then
    EnterKey
  End If
  If scrRead(9, 27, 47) = "Line Level Alert Code" Then
    EnterKey
  End If
  If scrRead(12, 36, 52) = "Wholesale Service" Then
    EnterKey
  End If

End Sub

Sub cursorThere(ByVal cursorPos As Integer)
  'Loop until the cursor is in a certain position
  Do Until queryCursorPosition = cursorPos: Loop
End Sub

Sub clearFields()
  'Use only for service order entry macros to clear the line, account & customer number
  'from the initial service order entry screen.
  CopyStringToPS "           ", 5, 28
  CopyStringToPS "         ", 6, 28
  CopyStringToPS "          ", 7, 28
End Sub

Function queryCursorPosition() As Integer
  Dim nRetValue As Integer
  Dim nCursorPosition As Integer
  Dim nPosition As Integer
  endTime = DateAdd("s", TIME_OUT_SECS, Now)
  Do While Now < endTime
    nRetValue = WD_QueryCursorLocation(lngPid, nCursorPosition)
    If nRetValue = 0 Then
      queryCursorPosition = nCursorPosition
      Exit Function
    Else
      Sleep 0.1
    End If
  Loop
  killApp ("Could not reslove cursor position!")
End Function

Sub CopyStringToPS(ByVal varExpression, ByVal rowPos As Integer, ByVal colPos As Integer)
  Dim nRetValue As Integer
  Dim nExpressionPosition As Integer
  Dim sGetExpression As String
  Dim nLength As Integer
  nExpressionPosition = (rowPos - 1) * 80 + colPos
  nLength = Len(varExpression)
  endTime = DateAdd("s", TIME_OUT_SECS, Now)
  Do While Now < endTime
    nRetValue = WD_CopyStringToPS(lngPid, nExpressionPosition, varExpression, nLength)
    If nRetValue = 0 Then
      OIAWait
      Exit Sub
    End If
  Loop

  errMessage = "CopyStringToPS(" & varExpression & ", " & rowPos & ", " & colPos & ")" & vbNewLine
  errMessage = errMessage & "Has failed!"
  killApp errMessage
End Sub

Sub setCursor(ByVal row As Integer, ByVal col As Integer)
  Dim curPos As Integer
  curPos = (row - 1) * 80 + col
  endTime = DateAdd("s", TIME_OUT_SECS, Now)
  Do While Now < endTime
    nRetValue = WD_SetCursor(lngPid, curPos)
    If nRetValue = 0 Then Exit Sub
  Loop
  killApp ("Set cursor function failed for position " & curPos & " application will close.")
End Sub

Function scrRead(ByVal row As Integer, ByVal colStart As Integer, ByVal colEnd As Integer, Optional DontTrimString As Boolean) As String
  Dim S      As Integer
  Dim e      As Integer
  Dim ret    As Integer
  Dim strLen As Integer
  Dim strRet As String
  S = (row - 1) * 80 + colStart
  e = (row - 1) * 80 + colEnd
  strLen = (e - S) + 1
  strRet = String$(strLen, 0)
  ret = WD_CopyPSToString(lngPid, S, strRet, strLen)

  If DontTrimString <> True Then
    scrRead = Trim(strRet)
  Else
    scrRead = strRet
  End If
End Function

Sub scrChange(row As Integer, colStart As Integer, colEnd As Integer, screenID As String)
  endTime = DateAdd("s", TIME_OUT_SECS, Now)
  Do While Now < endTime
    DoEvents
    If scrRead(row, colStart, colEnd) = screenID Then
      Exit Sub
    End If
  Loop
  killApp "Screen wait for screen " & screenID & " has failed!"
End Sub

Sub callSleep(ByVal milliSeconds As Integer)
  Sleep milliSeconds
End Sub

Sub SendKeyString(ByVal sendString As String)
  nRetValue = WD_SendKey(lngPid, sendString)
  Sleep 150
End Sub

Sub EnterKey()
  endTime = DateAdd("s", TIME_OUT_SECS, Now)
  Do While Now < endTime
    nRetValue = WD_SendKey(lngPid, RumbaEnterKey)
    If nRetValue = 0 Then
      OIAWait
      Exit Sub
    End If
  Loop
  killApp "Send key {ENTER} failed!"
End Sub

Sub PageUp()
  endTime = DateAdd("s", TIME_OUT_SECS, Now)
  Do While Now < endTime
    nRetValue = WD_SendKey(lngPid, RumbaPageUp)
    If nRetValue = 0 Then
      OIAWait
      Exit Sub
    End If
  Loop
  killApp "Send key {PgUp} failed!"
End Sub
Sub PageDown()
  endTime = DateAdd("s", TIME_OUT_SECS, Now)
  Do While Now < endTime
    nRetValue = WD_SendKey(lngPid, RumbaPageDown)
    If nRetValue = 0 Then
      OIAWait
      Exit Sub
    End If
  Loop
  killApp "Send key {PgDn} failed!"
End Sub

Sub FunctionKey_01()
  endTime = DateAdd("s", TIME_OUT_SECS, Now)
  Do While Now < endTime
    nRetValue = WD_SendKey(lngPid, RumbaCmd1)
    If nRetValue = 0 Then
      OIAWait
      Exit Sub
    End If
  Loop
  killApp "Send key {F1} failed!"
End Sub
Sub FunctionKey_02()
  endTime = DateAdd("s", TIME_OUT_SECS, Now)
  Do While Now < endTime
    nRetValue = WD_SendKey(lngPid, RumbaCmd2)
    If nRetValue = 0 Then
      OIAWait
      Exit Sub
    End If
  Loop
  killApp "Send key {F2} failed!"
End Sub
Sub FunctionKey_03()
  endTime = DateAdd("s", TIME_OUT_SECS, Now)
  Do While Now < endTime
    nRetValue = WD_SendKey(lngPid, RumbaCmd3)
    If nRetValue = 0 Then
      OIAWait
      Exit Sub
    End If
  Loop
  killApp "Send key {F3} failed!"
End Sub
Sub FunctionKey_04()
  endTime = DateAdd("s", TIME_OUT_SECS, Now)
  Do While Now < endTime
    nRetValue = WD_SendKey(lngPid, RumbaCmd4)
    If nRetValue = 0 Then
      OIAWait
      Exit Sub
    End If
  Loop
  killApp "Send key {F4} failed!"
End Sub
Sub FunctionKey_05()
  endTime = DateAdd("s", TIME_OUT_SECS, Now)
  Do While Now < endTime
    nRetValue = WD_SendKey(lngPid, RumbaCmd5)
    If nRetValue = 0 Then
      OIAWait
      Exit Sub
    End If
  Loop
  killApp "Send key {ENTER} failed!"
End Sub
Sub FunctionKey_06()
  endTime = DateAdd("s", TIME_OUT_SECS, Now)
  Do While Now < endTime
    nRetValue = WD_SendKey(lngPid, RumbaCmd6)
    If nRetValue = 0 Then
      OIAWait
      Exit Sub
    End If
  Loop
  killApp "Send key {F6} failed!"
End Sub
Sub FunctionKey_07()
  endTime = DateAdd("s", TIME_OUT_SECS, Now)
  Do While Now < endTime
    nRetValue = WD_SendKey(lngPid, RumbaCmd7)
    If nRetValue = 0 Then
      OIAWait
      Exit Sub
    End If
  Loop
  killApp "Send key {F7} failed!"
End Sub
Sub FunctionKey_08()
  endTime = DateAdd("s", TIME_OUT_SECS, Now)
  Do While Now < endTime
    nRetValue = WD_SendKey(lngPid, RumbaCmd8)
    If nRetValue = 0 Then
      OIAWait
      Exit Sub
    End If
  Loop
  killApp "Send key {F8} failed!"
End Sub
Sub FunctionKey_09()
  endTime = DateAdd("s", TIME_OUT_SECS, Now)
  Do While Now < endTime
    nRetValue = WD_SendKey(lngPid, RumbaCmd9)
    If nRetValue = 0 Then
      OIAWait
      Exit Sub
    End If
  Loop
  killApp "Send key {F9} failed!"
End Sub
Sub FunctionKey_10()
  endTime = DateAdd("s", TIME_OUT_SECS, Now)
  Do While Now < endTime
    nRetValue = WD_SendKey(lngPid, RumbaCmd10)
    If nRetValue = 0 Then
      OIAWait
      Exit Sub
    End If
  Loop
  killApp "Send key {F10} failed!"
End Sub
Sub FunctionKey_11()
  endTime = DateAdd("s", TIME_OUT_SECS, Now)
  Do While Now < endTime
    nRetValue = WD_SendKey(lngPid, RumbaCmd11)
    If nRetValue = 0 Then
      OIAWait
      Exit Sub
    End If
  Loop
  killApp "Send key {F11} failed!"
End Sub
Sub FunctionKey_12()
  endTime = DateAdd("s", TIME_OUT_SECS, Now)
  Do While Now < endTime
    nRetValue = WD_SendKey(lngPid, RumbaCmd12)
    If nRetValue = 0 Then
      OIAWait
      Exit Sub
    End If
  Loop
  killApp "Send key {F12} failed!"
End Sub
Sub FunctionKey_13()
  endTime = DateAdd("s", TIME_OUT_SECS, Now)
  Do While Now < endTime
    nRetValue = WD_SendKey(lngPid, RumbaCmd13)
    If nRetValue = 0 Then
      OIAWait
      Exit Sub
    End If
  Loop
  killApp "Send key {F13} failed!"
End Sub
Sub FunctionKey_14()
  endTime = DateAdd("s", TIME_OUT_SECS, Now)
  Do While Now < endTime
    nRetValue = WD_SendKey(lngPid, RumbaCmd14)
    If nRetValue = 0 Then
      OIAWait
      Exit Sub
    End If
  Loop
  killApp "Send key {F14} failed!"
End Sub
Sub FunctionKey_15()
  endTime = DateAdd("s", TIME_OUT_SECS, Now)
  Do While Now < endTime
    nRetValue = WD_SendKey(lngPid, RumbaCmd15)
    If nRetValue = 0 Then
      OIAWait
      Exit Sub
    End If
  Loop
  killApp "Send key {F15} failed!"
End Sub
Sub FunctionKey_16()
  endTime = DateAdd("s", TIME_OUT_SECS, Now)
  Do While Now < endTime
    nRetValue = WD_SendKey(lngPid, RumbaCmd16)
    If nRetValue = 0 Then
      OIAWait
      Exit Sub
    End If
  Loop
  killApp "Send key {F16} failed!"
End Sub

Sub FieldExitKey()
endTime = DateAdd("s", TIME_OUT_SECS, Now)
Do While Now < endTime
nRetValue = WD_SendKey(lngPid, RumbaFieldExit)
 If nRetValue = 0 Then
      OIAWait
      Exit Sub
    End If
  Loop
  killApp "Send key FieldExit failed!"
End Sub

Private Sub killApp(ByVal killMessage As String)
  MsgBox killMessage, vbCritical, "Function/Method has timed out!"
  Class_Terminate
  End
End Sub

Function onChorusMainMenu() As Boolean
Dim ScreenScraped As String
Dim Check As String
Check = "*** CHORUS CUSTOMER SERVICE ***"
ScreenScraped = scrRead(6, 24, 54, True)
If ScreenScraped = Check Then
onChorusMainMenu = True
Else
onChorusMainMenu = False
End If
End Function

Function onWWCMenu() As Boolean
Dim ScreenScraped As String
Dim Check As String
Check = "WORK WITH CUSTOMER"
ScreenScraped = scrRead(2, 32, 49, True)
If ScreenScraped = Check Then
onWWCMenu = True
Else
onWWCMenu = False
End If
End Function

Function onServiceOrderEntryScreen() As Boolean
Dim ScreenScraped As String
Dim Check As String
Check = "SOE01"
ScreenScraped = scrRead(1, 2, 6, True)
If ScreenScraped = Check Then
onServiceOrderEntryScreen = True
Else
onServiceOrderEntryScreen = False
End If
End Function

Function onServiceOrderEntryScreenEnquiry() As Boolean
Dim ScreenScraped As String
Dim Check As String
Check = "SV0001FM/01"
ScreenScraped = scrRead(1, 2, 12, True)
If ScreenScraped = Check Then
onServiceOrderEntryScreenEnquiry = True
Else
onServiceOrderEntryScreenEnquiry = False
End If
End Function

Function OrderOpenPosted() As Boolean
Dim ScreenScraped As String
Dim Check As String
Check = "POSTED"
ScreenScraped = scrRead(7, 38, 43, True)
If ScreenScraped = Check Then
OrderOpenPosted = True
Else
OrderOpenPosted = False
End If
End Function






