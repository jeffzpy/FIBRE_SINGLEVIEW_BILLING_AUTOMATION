VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSingleViewMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private aSvHwnd As Long
Private aMDI As Long
Private aRespondMsg As String

Public Property Get svHwnd() As Long
    Dim winapi As New clsWinAPI
    With winapi
        aSvHwnd = .getObjectHwnd("TfrmMain", vbNullString)
        svHwnd = aSvHwnd
    End With
    Set winapi = Nothing
End Property

Public Property Get mdi() As Long
    Dim winapi As New clsWinAPI
    With winapi
        If aSvHwnd = 0 Then
            Dim sv As New clsSVProcess
            Dim svMain As New clsSingleViewMain
            aSvHwnd = svHwnd
        End If
        aMDI = .getChildObjectHwnd(aSvHwnd, 0&, "MDIClient", vbNullString)
        mdi = aMDI
    End With
    Set winapi = Nothing
End Property

Public Property Get RespondMsg() As String
    RespondMsg = aRespondMsg
End Property

Public Function HandlingConfirmWin(ByVal YesNo As Boolean) As Long
    Dim winapi As New clsWinAPI
    Dim WinConfirm As Long, i As Integer, TXPanel As Long, TxpanelSubLevel As Long, TXPanelNextSibling As Long, _
        BtnYes As Long, BtnNo As Long
    With winapi
        For i = 0 To 80
            If WinConfirm = 0 Then
                .Sleeping 100
                WinConfirm = .getObjectHwnd(vbNullString, "Confirm")
                .Sleeping 100
            Else
                Exit For
            End If
        Next
        Do While TXPanel = 0
            .Sleeping 100
            TXPanel = .getChildObjectHwnd(WinConfirm, 0, "TXPanel", vbNullString)
            .Sleeping 100
        Loop
        For i = 0 To 1
            If i = 1 Then TXPanel = TXPanelNextSibling
            TXPanelNextSibling = .getChildObjectHwnd(WinConfirm, TXPanel, "TXPanel", vbNullString)
        Next
        BtnYes = .getChildObjectHwnd(TXPanelNextSibling, 0, vbNullString, "&Yes")
        BtnNo = .getChildObjectHwnd(TXPanelNextSibling, 0, vbNullString, "&No")
        If YesNo = True Then
            .PostmsgClick BtnYes
        Else
            .PostmsgClick BtnNo
        End If
        .Sleeping 100
    End With
End Function

Public Sub ClickBtnSave()
    Dim winapi As New clsWinAPI
    With winapi
        If aSvHwnd = 0 Then aSvHwnd = svHwnd
        Call .ActivateMenuObject(aSvHwnd, "SaveCtrl+S", 1, "postmsg")
        .Sleeping 250
    End With
End Sub

Public Sub ClickBtnUpdate()
    Dim winapi As New clsWinAPI
    Dim sv As New clsSingleViewMain
    Dim svHD As Long
    With sv
        svHD = .svHwnd
    End With
    With winapi
        .ActivateMenuObject svHD, "UpdateF9", 10, "postmsg"
        .Sleeping 50
    End With
    Set winapi = Nothing
    Set sv = Nothing
End Sub

Public Sub ClickBtnInsert()
    Dim winapi As New clsWinAPI
    Dim sv As New clsSingleViewMain
    Dim svHD As Long
    With sv
        svHD = .svHwnd
    End With
    With winapi
        .ActivateMenuObject svHD, "InsertCtrl+Ins", 10, "sendmsg"
        .Sleeping 250
    End With
    Set winapi = Nothing
    Set sv = Nothing
End Sub

Public Sub ClickBtnDelete()
    Dim winapi As New clsWinAPI
    Dim sv As New clsSingleViewMain
    Dim svHD As Long
    With sv
        svHD = .svHwnd
    End With
    With winapi
        .ActivateMenuObject svHD, "DeleteCtrl+Del", 10, "postmsg"
        .Sleeping 50
    End With
    Set winapi = Nothing
    Set sv = Nothing
End Sub

Public Sub ClickBtnCopy()
    Dim winapi As New clsWinAPI
    Dim sv As New clsSingleViewMain
    Dim svHD As Long
    With sv
        svHD = .svHwnd
    End With
    With winapi
        .ActivateMenuObject svHD, "Copy", 10, "postmsg"
        .Sleeping 500
    End With
    Set winapi = Nothing
    Set sv = Nothing
End Sub

Public Sub CloseAllChildWindows()
    Dim winapi As New clsWinAPI
    With winapi
        If aSvHwnd = 0 Then aSvHwnd = svHwnd
        Call .ActivateMenuObject(aSvHwnd, "Close All", 5, "postmsg")
        .Sleeping 500
    End With
End Sub

Public Function WaitTConfirmForm() As Long
    Dim winapi As New clsWinAPI, tempHwnd As Long, BtnYes As Long, BtnNo As Long
    With winapi
        .Sleeping 250
        tempHwnd = .getObjectHwnd("Tform", "Confirm")
        If tempHwnd <> 0 Then
            WaitTConfirmForm = tempHwnd
        Else
            WaitTConfirmForm = 0
        End If
    End With
    Set winapi = Nothing
End Function

Public Sub ProcessTimeConfirmForm(ByVal YesorNo As Boolean, Optional aDate As String)
    Dim winapi As New clsWinAPI, tempHwnd As Long, BtnYes As Long, BtnNo As Long
    Dim TADBDateTime As Long, DateEditor As Long
    With winapi
        .Sleeping 250
        tempHwnd = .getObjectHwnd("Tform", "Confirm")
        If tempHwnd <> 0 Then
            BtnYes = .getChildObjectHwnd(tempHwnd, 0, "TXBitBtn", "&Yes")
            BtnNo = .getChildObjectHwnd(tempHwnd, 0, "TXBitBtn", "&No")
            If aDate <> "" Then
                TADBDateTime = .getChildObjectHwnd(tempHwnd, 0, "TADBDateTime", vbNullString)
                DateEditor = .getChildObjectHwnd(TADBDateTime, 0, "TMyAdrockCalendarEdit", vbNullString)
                .ObjSetFocus DateEditor
                .SendText DateEditor, aDate
                .ObjKillFocus DateEditor
            End If
            If YesorNo = True Then
                .PostmsgClick BtnYes
            Else
                .PostmsgClick BtnNo
            End If
        End If
    End With
    Set winapi = Nothing
End Sub

Public Function MsgFormChk() As Long
    Dim winapi As New clsWinAPI, tempHwnd As Long
    With winapi
        .Sleeping 250
        tempHwnd = .getObjectHwnd("TMessageForm", vbNullString)
        If tempHwnd = 0 Then tempHwnd = .getObjectHwnd("TXMessageForm", vbNullString)
        If tempHwnd <> 0 Then
            MsgFormChk = tempHwnd
        Else
            MsgFormChk = 0
        End If
    End With
    Set winapi = Nothing
End Function

Public Sub MsgFormClick()
    Dim winapi As New clsWinAPI, tempHwnd As Long
    Dim InforTXPanel1 As Long, InforTXPanel2 As Long, InforTXPanel3 As Long, i As Integer, _
        btnok As Long
    With winapi
        .Sleeping 250
        tempHwnd = .getObjectHwnd("TMessageForm", vbNullString)
        If tempHwnd = 0 Then tempHwnd = .getObjectHwnd("TXMessageForm", vbNullString)
        If tempHwnd <> 0 Then
            .CloseObjWindow tempHwnd
        End If
    End With
    Set winapi = Nothing
End Sub

Public Function GetSVStatusBarMsg() As String
    Dim winapi As New clsWinAPI
    Dim TXStatusBar As Long
    With winapi
        aSvHwnd = svHwnd
        TXStatusBar = .getChildObjectHwnd(aSvHwnd, 0, "TXStatusBar", vbNullString)
        GetSVStatusBarMsg = .getNameFromHwnd(TXStatusBar)
    End With
End Function

Public Sub InitialSizeOfSV()
    Dim winapi As New clsWinAPI
    With winapi
        aSvHwnd = svHwnd
        .SetaWindowPos aSvHwnd, 0, 0, 1000, 600, &H40
    End With
End Sub
