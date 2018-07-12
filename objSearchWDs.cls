VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "objSearchWDs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private aAcctId As Long
Private aASID As Long
Private aCircuitId As String
Private bCircuitId As String
Private aErrCode As String
Private aSearchText As String
Private aGetSearchText As Boolean

Property Let AcctId(Arg As Long)
    aAcctId = Arg
End Property

Property Let Asid(Arg As Long)
    aASID = Arg
End Property

Property Let circuitID(Arg As String)
    aCircuitId = Arg
End Property

Public Property Get circuitID() As String
    circuitID = aCircuitId
End Property

Public Property Get circuitID4ChangeOffer() As String
    circuitID4ChangeOffer = bCircuitId
End Property

Public Property Let circuitID4ChangeOffer(Arg As String)
    bCircuitId = Arg
End Property

Public Property Let GetSearchText(GetSearchText As Boolean)
    aGetSearchText = GetSearchText
End Property

Public Property Get ErrCode() As String
    ErrCode = aErrCode
End Property

Public Property Get SearchText() As String
    SearchText = aSearchText
End Property

Public Function SearchAcct(ByVal AcctId As Long) As Long
    Dim sv As New clsSingleViewMain
    Dim winapi As New clsWinAPI, i As Integer, i2 As Integer
    Dim WinRectA As Variant, WinRectB As Variant
    Dim svHwnd As Long, aMdi As Long
    Dim TfrmCustomerSearch As Long, TXPanel As Long, TxpanelSubLevel As Long, TXPanelNextSibling As Long, _
        TXPageControl As Long, TTabSheet As Long, TXDataEdit As Long, TXCheckBox_CF As Long, TXBitBtn As Long, _
        TXCheckBox_LH As Long, TXCheckBox_DH As Long, SearchBtn As Long
    aErrCode = ""
    aAcctId = AcctId
    With sv
        svHwnd = .svHwnd
        aMdi = .mdi
    End With
    With winapi
        .ActivateMenuObject svHwnd, "Manage Customers", 30, "sendmsg"
        Do While TfrmCustomerSearch = 0
            TfrmCustomerSearch = .getChildObjectHwnd(aMdi, 0, "TfrmCustomerSearch", vbNullString)
        Loop
        For i = 0 To 2
            Select Case i
                Case 0
                    TXPanel = .getChildObjectHwnd(TfrmCustomerSearch, 0, "TXPanel", vbNullString)
                    TXPageControl = .getChildObjectHwnd(TfrmCustomerSearch, 0, "TXPageControl", vbNullString)
                    TTabSheet = .getChildObjectHwnd(TXPageControl, 0, "TTabSheet", "&Details")
                Case 1
                    TXPanel = .getChildObjectHwnd(TfrmCustomerSearch, TXPanel, "TXPanel", vbNullString)
                    TXPanel = .getChildObjectHwnd(TXPanel, 0, "TXPanel", vbNullString)
                    TXCheckBox_CF = .getChildObjectHwnd(TXPanel, 0, "TXCheckBox", "Reuse Customer Form?")
                    TXCheckBox_LH = .getChildObjectHwnd(TXPanel, 0, "TXCheckBox", "Reuse Last Hierarchy?")
                    TXCheckBox_DH = .getChildObjectHwnd(TXPanel, 0, "TXCheckBox", "Always Display Hierarchy?")
                    SearchBtn = .getChildObjectHwnd(TXPanel, 0, "TXBitBtn", "&Search...")
                    If .ChkboxStatus(TXCheckBox_CF) = True Then .ChkboxClick (TXCheckBox_CF)
                    If .ChkboxStatus(TXCheckBox_LH) = False Then .ChkboxClick (TXCheckBox_LH)
                    If .ChkboxStatus(TXCheckBox_DH) = False Then .ChkboxClick (TXCheckBox_DH)
                Case 2
                    TXPanel = .getChildObjectHwnd(TTabSheet, 0, "TXPanel", vbNullString)
                    TXDataEdit = .getChildObjectHwnd(TXPanel, 0, "TXDataEdit", vbNullString)
                    WinRectA = .GetWindowsRect(TXPanel)
                    WinRectB = Array(0, 0)
                    Do While WinRectB(1) - WinRectA(1) <> 104
                         TXDataEdit = .getChildObjectHwnd(TXPanel, TXDataEdit, "TXDataEdit", vbNullString)
                        WinRectB = .GetWindowsRect(TXDataEdit)
                    Loop
                    .SendText TXDataEdit, CStr(aAcctId)
            End Select
        Next
        .PostmsgClick SearchBtn
        .Sleeping 1000
        'aRespondMsg = MsgFormChk
    End With
    Set winapi = Nothing
End Function

Public Function SearchCircuitByAsid() As Long
    Dim winapi As New clsWinAPI, i As Integer, i2 As Integer
    Dim sv As New clsSingleViewMain
    Dim sd As New objServiceDetailWD
    Dim svHwnd As Long, aMdi As Long
    Dim TfrmDynamicFormEV As Long, TXPanel As Long, TxpanelSubLevel As Long, _
        TAInplaceGrid As Long, TXPanelNextSibling As Long, _
        TAEntityValidationPanel As Long, TXDataEdit As Long, btnSearch As Long, _
        TfrmDynamicFormDA As Long
    Dim WinRect As Variant
    aErrCode = ""
    aSearchText = vbNullString
    With sv
        svHwnd = .svHwnd
        aMdi = .mdi
    End With
    With winapi
        .ActivateMenuObject svHwnd, "Locate Customer/Product", 1, "sendmsg"
        Do While TfrmDynamicFormEV = 0
            TfrmDynamicFormEV = .getChildObjectHwnd(aMdi, 0, "TfrmDynamicFormEV", vbNullString)
        Loop
        For i = 0 To 1
            Select Case i
                Case 0
                    Do While TxpanelSubLevel = 0
                        TXPanel = .getChildObjectHwnd(TfrmDynamicFormEV, 0, "TXPanel", vbNullString)
                        For i2 = 0 To 2
                            TxpanelSubLevel = .getChildObjectHwnd(TXPanel, 0, "TXPanel", vbNullString)
                        Next
                        TXPanelNextSibling = .getChildObjectHwnd(TXPanel, 0, "TXPanel", vbNullString)
                        TxpanelSubLevel = .getChildObjectHwnd(TXPanelNextSibling, 0, "TXPanel", vbNullString)
                        btnSearch = .getChildObjectHwnd(TxpanelSubLevel, 0, "TXBitBtn", vbNullString)
                    Loop
                Case 1
                    TXPanel = .getChildObjectHwnd(TXPanel, TXPanelNextSibling, "TXPanel", vbNullString)
                    TAEntityValidationPanel = .getChildObjectHwnd(TXPanel, 0, "TAEntityValidationPanel", vbNullString)
                    TXDataEdit = .getChildObjectHwnd(TAEntityValidationPanel, 0, "TXDataEdit", vbNullString)
            End Select
        Next
        .SendText TXDataEdit, CStr(aASID)
        .PostmsgClick btnSearch
        .Sleeping 900
        If sv.MsgFormChk > 0 Then
            If InStr(winapi.GetObjectText(sv.MsgFormChk), "Service Not Found") > 0 Then
                sv.MsgFormClick
                aErrCode = "Service Not Found"
                sv.CloseAllChildWindows
                Exit Function
            Else
                sv.MsgFormChk
            End If
        End If
        TXPanel = 0: TxpanelSubLevel = 0: TXPanelNextSibling = 0: TAInplaceGrid = 0
        Do While TfrmDynamicFormDA = 0
            TfrmDynamicFormDA = .getChildObjectHwnd(aMdi, 0, "TfrmDynamicFormDA", vbNullString)
        Loop
        .Sleeping 25
        Do While TXPanel = 0: TXPanel = .getChildObjectHwnd(TfrmDynamicFormDA, 0, "TXPanel", vbNullString): Loop
        .Sleeping 25
        Do While TxpanelSubLevel = 0: TxpanelSubLevel = .getChildObjectHwnd(TXPanel, 0, "TXPanel", vbNullString): Loop
        .Sleeping 25
        Do While TXPanelNextSibling = 0: TXPanelNextSibling = .getChildObjectHwnd(TXPanel, TxpanelSubLevel, "TXPanel", vbNullString): Loop
        .Sleeping 25
        Do While TAInplaceGrid = 0: TAInplaceGrid = .getChildObjectHwnd(TXPanelNextSibling, 0, "TAInplaceGrid", vbNullString): Loop
        .Sleeping 25
        .BkgrndSingleButtonClick TAInplaceGrid, 2, 37, 9
        For i = 0 To 7
            .VKeyDown TAInplaceGrid
            .Sleeping 25
        Next
        .SetClipboard vbNullString
        .Sleeping 50
        If aGetSearchText = True Then
            Do While InStr(aSearchText, "Account Number") < 1
                    If sv.MsgFormChk <> 0 Then
                        sv.MsgFormClick
                        .SetClipboard vbNullString
                        .Sleeping 2000
                    End If
                    sv.ClickBtnCopy
                    .Sleeping 250
                    aSearchText = .ClipBoard_GetData
            Loop
        End If
        TxpanelSubLevel = .getChildObjectHwnd(TxpanelSubLevel, 0, "TXPanel", vbNullString)
        .Sleeping 25
        btnSearch = .getChildObjectHwnd(TxpanelSubLevel, 0, "TXBitBtn", "&Details...")
        .SendmsgClick btnSearch
        sd.Init = aMdi
        aCircuitId = sd.GetPSNValue
    End With
    Set winapi = Nothing
    Set sv = Nothing
    Set sd = Nothing
End Function

Public Function SearchCircuitByCircuitID() As Long
    Dim winapi As New clsWinAPI, i As Integer, i2 As Integer
    Dim sv As New clsSingleViewMain
    Dim sd As New objServiceDetailWD
    Dim svHwnd As Long, aMdi As Long
    Dim Tfrm_service_search As Long, TXPageC_SS As Long, TTab_details As Long, TXPanel_details As Long, _
        TXDateEdit_details As Long, Sub_TXPanel_SS As Long, btn_search_SS As Long, TfrmDynamicFormDA As Long, TfrmNormalServiceList As Long, _
        ServiceTmessageForm As Long, TXPanel As Long, TXPanel_Sub1 As Long, TXPanel_sub2 As Long, TAInplaceGrid As Long, TXPanel_SS As Long
    Dim aHndCboxReuseDF As Long, aHndCboxAlwaysDL As Long
    Dim WinRect As Variant
    aErrCode = ""
    If InStr(aCircuitId, "CFS") > 0 Then aCircuitId = Replace(aCircuitId, "CFS", "")
    With sv
        svHwnd = .svHwnd
        aMdi = .mdi
    End With
    With winapi
        .ActivateMenuObject svHwnd, "Manage Services", 1, "sendmsg"
        Do While Tfrm_service_search = 0
            Tfrm_service_search = .getChildObjectHwnd(aMdi, 0, "TfrmNormalServiceSearch", vbNullString)
        Loop
        TXPageC_SS = .getChildObjectHwnd(Tfrm_service_search, 0, "TXPageControl", vbNullString)
        TTab_details = .getChildObjectHwnd(TXPageC_SS, 0, "TTabSheet", vbNullString)
        TTab_details = .getChildObjectHwnd(TXPageC_SS, TTab_details, "TTabSheet", vbNullString)
        TXPanel_details = .getChildObjectHwnd(TTab_details, 0, "TXPanel", vbNullString)
        For i = 1 To 3
            If i = 1 Then
                TXDateEdit_details = .getChildObjectHwnd(TXPanel_details, 0, "TXDataEdit", vbNullString)
            Else
                TXDateEdit_details = .getChildObjectHwnd(TXPanel_details, TXDateEdit_details, "TXDataEdit", vbNullString)
            End If
        Next i
        .Sleeping 500
        .SendText TXDateEdit_details, CStr(aCircuitId)
        TXPanel_SS = .getChildObjectHwnd(Tfrm_service_search, TXPageC_SS, "TXPanel", vbNullString)
        Sub_TXPanel_SS = .getChildObjectHwnd(TXPanel_SS, 0, "TXPanel", vbNullString)
        btn_search_SS = .getChildObjectHwnd(Sub_TXPanel_SS, 0, "TXBitBtn", "&Search...")
        aHndCboxReuseDF = .getChildObjectHwnd(Sub_TXPanel_SS, 0, "TXCheckBox", "Reuse Detail Form?")
        aHndCboxAlwaysDL = .getChildObjectHwnd(Sub_TXPanel_SS, 0, "TXCheckBox", "Always Display List?")
        .TickCheckBox aHndCboxReuseDF, True
        .TickCheckBox aHndCboxAlwaysDL, False
        .PostmsgClick btn_search_SS
        .Sleeping 1000
        If sv.MsgFormChk <> 0 Then
            sv.MsgFormClick
            aErrCode = "Service Not Found"
            sv.CloseAllChildWindows
            Set winapi = Nothing
            Set sv = Nothing
            Set sd = Nothing
            Exit Function
        Else
            For i = 1 To 40
                TfrmDynamicFormDA = .getChildObjectHwnd(aMdi, 0, "TfrmServiceDetail", vbNullString)
                If TfrmDynamicFormDA <> 0 Then Exit For
                Sleep 100
            Next i
        End If
    End With
    Set winapi = Nothing
    Set sv = Nothing
    Set sd = Nothing
End Function

