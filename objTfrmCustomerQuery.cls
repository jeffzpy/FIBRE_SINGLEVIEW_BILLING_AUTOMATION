VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "objTfrmCustomerQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit

Private aHwndCQ As Long
Private aHndTAQryLookupCombo As Long
Private aHndTXPageControl As Long
Private aHndTabService As Long
Private aHndTabGeneral As Long
Private aHndCalendarEdit As Long
Private aHndReasonLookupCombo As Long
Private aHndTXAmountEdit As Long
Private aHndTXInviceTxtEdit As Long
Private aHndBtnAssign As Long

Private aPortalID As Long
Private aChargeDate As String
Private aOneOffChargeType As String
Private aInvoiceText As String
Private aReasonCode As String
Private aAmount As String

Public Property Let Init(MdiHwnd As Long)
    Dim HwndCQ As Long, TXPanel_D As Long, TXPanel As Long, TXPanelNextSibling As Long, TxpanelSubLevel As Long, _
        TXScrollBox As Long, TXDataEdit As Long, TTabSheet As Long, TADBLookupCombo As Long, TXGroupBox As Long
    Dim winapi As New clsWinAPI, i As Integer

    With winapi
        Do While HwndCQ = 0
            If aHwndCQ <> 0 Then
                HwndCQ = aHwndCQ
            Else
                HwndCQ = .getChildObjectHwnd(MdiHwnd, 0, "TfrmCustomerQuery", vbNullString)
                aHwndCQ = HwndCQ
            End If
        Loop
        Do While TxpanelSubLevel = 0
            TXPanel = .getChildObjectHwnd(HwndCQ, 0, "TXPanel", vbNullString)
            TxpanelSubLevel = .getChildObjectHwnd(TXPanel, 0, "TXPanel", vbNullString)
            TxpanelSubLevel = .getChildObjectHwnd(TXPanel, TxpanelSubLevel, "TXPanel", vbNullString)
        Loop
        TXGroupBox = .getChildObjectHwnd(TxpanelSubLevel, 0, "TXGroupBox", vbNullString)
        aHndTAQryLookupCombo = .getChildObjectHwnd(TXGroupBox, 0, "TADBLookupCombo", vbNullString)
        
        aHndTXPageControl = .getChildObjectHwnd(TXPanel, 0, "TXPageControl", vbNullString)
        aHndTabService = .getChildObjectHwnd(aHndTXPageControl, 0, "TTabSheet", "&Service")
        TXGroupBox = .getChildObjectHwnd(aHndTabService, 0, "TXGroupBox", "Service Details")
        aHndBtnAssign = .getChildObjectHwnd(TXGroupBox, 0, "TXBitBtn", "Assign")
        
        aHndTabGeneral = .getChildObjectHwnd(aHndTXPageControl, 0, "TTabSheet", "&General")
        TXScrollBox = .getChildObjectHwnd(aHndTabGeneral, 0, "TXScrollBox", vbNullString)
        aHndCalendarEdit = .getChildObjectHwnd(TXScrollBox, 0, "TAAdrockCalendarEdit", vbNullString)
        TADBLookupCombo = .getChildObjectHwnd(TXScrollBox, 0, "TADBLookupCombo", vbNullString)
        If aOneOffChargeType = "3" Then
           TADBLookupCombo = .getChildObjectHwnd(TXScrollBox, TADBLookupCombo, "TADBLookupCombo", vbNullString)
        End If
        aHndReasonLookupCombo = .getChildObjectHwnd(TXScrollBox, TADBLookupCombo, "TADBLookupCombo", vbNullString)
        aHndTXInviceTxtEdit = .getChildObjectHwnd(TXScrollBox, 0, "TXDataEdit", vbNullString)
        If aOneOffChargeType = "3" Then
            TXDataEdit = .getChildObjectHwnd(TXScrollBox, aHndTXInviceTxtEdit, "TXDataEdit", vbNullString)
            aHndTXAmountEdit = .getChildObjectHwnd(TXScrollBox, TXDataEdit, "TXDataEdit", vbNullString)
        Else
            aHndTXAmountEdit = .getChildObjectHwnd(TXScrollBox, aHndTXInviceTxtEdit, "TXDataEdit", vbNullString)
        End If
    End With
End Property

Public Property Let PortalID(bPortalID As Long)
    aPortalID = bPortalID
End Property

Public Property Let OneOffChargeType(Arg As String)
    aOneOffChargeType = Arg
End Property

Public Property Let InvoiceText(Arg As String)
    aInvoiceText = Arg
End Property

Public Property Let ReasonCode(Arg As String)
    aReasonCode = Arg
End Property

Public Property Let Amount(Arg As String)
    aAmount = Arg
End Property

Public Property Let ChargeDate(Arg As String)
    aChargeDate = Arg
End Property

Public Sub SelectQryType(ByVal QryTypeDropIndex As Integer)
    Dim winapi As New clsWinAPI
    Dim sv As New clsSingleViewMain
    Dim svHD As Long, mdi As Long
    Dim qryTxt As String
    svHD = sv.svHwnd
    mdi = sv.mdi
    Init = mdi
    If winapi.IsWinVisible(aHndTAQryLookupCombo) = False Then Exit Sub
    Select Case QryTypeDropIndex
        Case 1
            qryTxt = "Dispute"
        Case 2
            qryTxt = "One Off Charges"
        Case 3
            qryTxt = "One Off Charges Variable"
    End Select
    With winapi
        .ObjSetFocus aHndTAQryLookupCombo
        .SendText aHndTAQryLookupCombo, qryTxt
        .ObjKillFocus aHndTAQryLookupCombo
    End With
End Sub

Public Sub TypeinChargeDate()
    Dim winapi As New clsWinAPI
    Dim sv As New clsSingleViewMain
    Dim svHD As Long, mdi As Long
    If aChargeDate = "" Then Exit Sub
    svHD = sv.svHwnd
    mdi = sv.mdi
    Init = mdi
    If winapi.IsWinVisible(aHndTabGeneral) = False Then Exit Sub
    With winapi
        .ObjSetFocus aHndCalendarEdit
        .SendText aHndCalendarEdit, aChargeDate
        .ObjKillFocus aHndCalendarEdit
    End With
End Sub

Public Sub TypeinChargeCode()
    Dim winapi As New clsWinAPI
    Dim sv As New clsSingleViewMain
    Dim svHD As Long, mdi As Long
    If aReasonCode = "" Then Exit Sub
    svHD = sv.svHwnd
    mdi = sv.mdi
    Init = mdi
    If winapi.IsWinVisible(aHndTabGeneral) = False Then Exit Sub
    With winapi
        .ObjSetFocus aHndReasonLookupCombo
        .SendText aHndReasonLookupCombo, Trim(aReasonCode)
        .ObjKillFocus aHndReasonLookupCombo
        .VKeyTab aHndReasonLookupCombo
    End With
End Sub

Public Sub TypeinChargeAmount()
    Dim winapi As New clsWinAPI
    Dim sv As New clsSingleViewMain
    Dim svHD As Long, mdi As Long, slicedChar As String
    If aAmount = "" Or aAmount = 0 Then Exit Sub
    If aOneOffChargeType <> "3" Then Exit Sub
    winapi.Sleeping 500
    svHD = sv.svHwnd
    mdi = sv.mdi
    Init = mdi
    If winapi.IsWinVisible(aHndTabGeneral) = False Then Exit Sub
    If InStr(aAmount, ".") < 1 Then aAmount = aAmount & ".00"
    With winapi
        .Sleeping 15
        .ObjSetFocus aHndTXAmountEdit
        .Sleeping 5
        .SendText aHndTXAmountEdit, aAmount
'        For i = 1 To Len(Trim(aAmount))
'            slicedChar = Mid(Trim(str(aAmount)), i, 1)
'            .KeyTypeIn aHndTXAmountEdit, slicedChar, 0&, 0&
'        Next
        .Sleeping 5
        .ObjKillFocus aHndTXAmountEdit
        .Sleeping 15
    End With
End Sub

Public Sub TypeinInvoiceText()
    Dim winapi As New clsWinAPI
    Dim sv As New clsSingleViewMain
    Dim svHD As Long, mdi As Long
    Dim i As Integer, slicedChar As String
    'If aAmount = "" Or aAmount = 0 Then Exit Sub
    'If aOneOffChargeType <> "One Off Charges Variable" Then Exit Sub
    svHD = sv.svHwnd
    mdi = sv.mdi
    Init = mdi
    If winapi.IsWinVisible(aHndTabGeneral) = False Then Exit Sub
    With winapi
        .ObjSetFocus aHndTXInviceTxtEdit
        '.BkgrndDoubleButtonClick aHndTXInviceTxtEdit, 240, 10
        .BkgrndSingleButtonClick aHndTXInviceTxtEdit, 2, 221, 10
        .ObjKillFocus aHndTXInviceTxtEdit
        .Sleeping 50
        For i = 0 To 20
            .VKeyRit aHndTXInviceTxtEdit
        Next
        .KeyTypeIn aHndTXInviceTxtEdit, " ", 0&, 0&
        .Sleeping 15
        For i = 1 To Len(Trim(CStr(aPortalID)))
            slicedChar = Mid(Trim(str(aPortalID)), i, 1)
            .KeyTypeIn aHndTXInviceTxtEdit, slicedChar, 0&, 0&
        Next
    End With
End Sub

Public Sub EnableServiceTab()
    Dim winapi As New clsWinAPI
    With winapi
        .ActivateATab aHndTXPageControl, 1
    End With
    winapi.Sleeping 200
    Set winapi = Nothing
End Sub

Public Sub EnableGeneralTab()
    Dim winapi As New clsWinAPI
    With winapi
        .ActivateATab aHndTXPageControl, 3
    End With
    Set winapi = Nothing
End Sub

Public Sub ClickBtnAssign()
    Dim winapi As New clsWinAPI
    Dim sv As New clsSingleViewMain, svHD As Long, mdi As Long
    svHD = sv.svHwnd
    mdi = sv.mdi
    Init = mdi
    winapi.Sleeping 500
    If aHndBtnAssign = 0 Then Exit Sub
    If winapi.IsWinVisible(aHndTabService) = False Then Exit Sub
    With winapi
        .PostmsgClick aHndBtnAssign
        .Sleeping 500
        Do While sv.MsgFormChk = 0
            .Sleeping 500
        Loop
    End With
    sv.HandlingConfirmWin True
    Set winapi = Nothing
End Sub
