VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "objProductInstanceWD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private aHwndPI As Long
Private aHndTXPageControl As Long
Private aHndClndarStartDate As Long
Private aHndTabDetail As Long
Private aHndTabComponents As Long
Private aHndTabIndustry As Long
Private aHndTabGeneral As Long
Private aHndLstProductName As Long
Private aHndLstProductSubTypeHnd As Long
Private aHndLstAcctId As Long
Private aHndTxtPSN As Long
Private aHndBtnCreate As Long
Private aHndBtnDetail As Long
Private aHndStatus As Long

Private aStartDate As String
Private aProductName As String
Private aProductSubType As String
Private aCFS As String
Private aCFS4ChangeOffer As String
Private aAcctId As String

Public Property Let Init(Args As Long)
    Dim HwndPI As Long, TXPanel_D As Long, TXPanel As Long, TXPanelNextSibling As Long, TxpanelSubLevel As Long, TXScrollBox As Long, _
        TXPanelNextSibling_sub As Long, longTXScrollBox As Long, TXDataEdit As Long, TXDataEdit_sibling As Long, TTabSheet As Long
    Dim winapi As New clsWinAPI, i As Integer
    aProductName = "NGA"
    aProductSubType = "Standard"
    With winapi
        .Sleeping 2000
        Do While TXPanel = 0
            HwndPI = .getChildObjectHwnd(Args, 0, "TfrmProductInstanceDetail", vbNullString)
            aHwndPI = HwndPI
            TXPanel = .getChildObjectHwnd(HwndPI, 0, "TXPanel", vbNullString)
            TXPanel = .getChildObjectHwnd(HwndPI, TXPanel, "TXPanel", vbNullString)
        Loop
        Do While aHndTXPageControl = 0
            aHndTXPageControl = .getChildObjectHwnd(TXPanel, 0, "TXPageControl", vbNullString)
        Loop
        aHndTabDetail = .getChildObjectHwnd(aHndTXPageControl, 0, "TTabSheet", "&Detail")
        aHndTabGeneral = .getChildObjectHwnd(aHndTXPageControl, 0, "TTabSheet", "General")
        TXScrollBox = .getChildObjectHwnd(aHndTabGeneral, 0, "TXScrollBox", vbNullString)
        For i = 0 To 1
            Select Case i
                Case 0
                    TXDataEdit = .getChildObjectHwnd(TXScrollBox, 0, "TXDataEdit", vbNullString)
                    aHndTxtPSN = TXDataEdit
                    aHndLstAcctId = .getChildObjectHwnd(TXScrollBox, 0, "TADBLookupCombo", vbNullString)
                Case 1
                    aHndLstProductSubTypeHnd = .getChildObjectHwnd(TXScrollBox, aHndLstAcctId, "TADBLookupCombo", vbNullString)
                    aHndLstProductName = .getChildObjectHwnd(TXScrollBox, aHndLstProductSubTypeHnd, "TADBLookupCombo", vbNullString)
            End Select
        Next i
        TxpanelSubLevel = .getChildObjectHwnd(TXPanel, 0, "TXPanel", vbNullString)
        TxpanelSubLevel = .getChildObjectHwnd(TXPanel, TxpanelSubLevel, "TXPanel", vbNullString)
        aHndStatus = .getChildObjectHwnd(TxpanelSubLevel, 0, "TADBLookupCombo", vbNullString)
        TXDataEdit = .getChildObjectHwnd(TxpanelSubLevel, 0, "TADBDateTime", vbNullString)
        TXDataEdit_sibling = .getChildObjectHwnd(TxpanelSubLevel, TXDataEdit, "TADBDateTime", vbNullString)
        TXDataEdit_sibling = .getChildObjectHwnd(TxpanelSubLevel, TXDataEdit_sibling, "TADBDateTime", vbNullString)
        aHndClndarStartDate = .getChildObjectHwnd(TXDataEdit_sibling, 0, "TMyAdrockCalendarEdit", vbNullString)
        '====================
        TXPanel_D = .getChildObjectHwnd(aHndTabDetail, 0, "TXPanel", vbNullString)
        TXPanelNextSibling = .getChildObjectHwnd(aHndTabDetail, TXPanel_D, "TXPanel", vbNullString)
        TXPanelNextSibling_sub = .getChildObjectHwnd(TXPanelNextSibling, 0, "TXPanel", vbNullString)
        aHndBtnCreate = .getChildObjectHwnd(TXPanelNextSibling, 0, "TXBitBtn", "Create...")
        aHndBtnDetail = .getChildObjectHwnd(TXPanelNextSibling_sub, 0, "TXBitBtn", "D&etails...")
    End With
End Property

Public Property Let StartDate(DateValue As String)
    aStartDate = DateValue
End Property

Public Property Let CFS(cfsValue As String)
    cfsValue = Replace(cfsValue, "CFS", "")
    aCFS = Trim(cfsValue)
End Property
'aCFS4ChangeOffer
Public Property Let CFS4ChangeOffer(bCFS4ChangeOffer As String)
    aCFS4ChangeOffer = bCFS4ChangeOffer
End Property

Public Property Let AcctId(AcctValue As String)
    aAcctId = AcctValue
End Property

Public Function EnableProductInstanceWindow() As Long
    Dim winapi As New clsWinAPI, i As Integer, i2 As Integer
    Dim sv As New clsSingleViewMain
    Dim aSvHwnd As Long, aMdi As Long
    Dim TfrmCustomerHierarchy As Long, TXPanel As Long, TxpanelSubLevel As Long, TXPanelNextSibling As Long, _
        TXTreeView As Long, TXGroupBox As Long, TXCheckBox As Long, TXBitBtn As Long, WhMenuTrue As Boolean
    Dim WinRect As Variant
    winapi.Sleeping 1000
    With sv
        aSvHwnd = .svHwnd
        aMdi = .mdi
    End With
    With winapi
        Do While TfrmCustomerHierarchy = 0
            TfrmCustomerHierarchy = .getChildObjectHwnd(aMdi, 0, "TfrmCustomerHierarchy", vbNullString)
        Loop
        For i = 0 To 4
            Select Case i
                Case 0
                    TXPanel = .getChildObjectHwnd(TfrmCustomerHierarchy, 0, "TXPanel", vbNullString)
                    TXPanelNextSibling = .getChildObjectHwnd(TfrmCustomerHierarchy, TXPanel, "TXPanel", vbNullString)
                Case 1
                    TXPanelNextSibling = .getChildObjectHwnd(TfrmCustomerHierarchy, TXPanelNextSibling, "TXPanel", vbNullString)
                Case 2
                    TXPanel = .getChildObjectHwnd(TXPanel, TXPanel, "TXPanel", vbNullString)
                    TxpanelSubLevel = .getChildObjectHwnd(TXPanel, 0, "TXPanel", vbNullString)
                    TXGroupBox = .getChildObjectHwnd(TxpanelSubLevel, 0, "TXGroupBox", vbNullString)
                    TXCheckBox = .getChildObjectHwnd(TXGroupBox, 0, "TXCheckBox", vbNullString)
                    Do Until TXGroupBox = 0
                        If .ChkboxStatus(TXCheckBox) = True Then .ChkboxClick (TXCheckBox)
                        TXCheckBox = .getChildObjectHwnd(TXGroupBox, TXCheckBox, "TXCheckBox", vbNullString)
                    Loop
                Case 3
                    TXPanel = .getChildObjectHwnd(TXPanelNextSibling, 0, "TXPanel", vbNullString)
                    TXPanelNextSibling = .getChildObjectHwnd(TfrmCustomerHierarchy, TXPanelNextSibling, "TXPanel", vbNullString)
                    TXTreeView = .getChildObjectHwnd(TXPanelNextSibling, 0, "TXTreeView", vbNullString)
                    TXCheckBox = .getChildObjectHwnd(TXPanel, 0, "TXCheckBox", "Reuse Form?")
                    If .ChkboxStatus(TXCheckBox) = True Then .ChkboxClick (TXCheckBox)
                Case 4
                    .BkgrndSingleButtonClick TXTreeView, 1, 4, 16
                    .ActivateMenuObject aSvHwnd, "Wholesale", 299, "sendmsg"
            End Select
        Next
        Do While aHwndPI = 0
            aHwndPI = .getChildObjectHwnd(aMdi, 0, "TfrmProductInstanceDetail", vbNullString)
        Loop
    End With
End Function

Public Sub EnableDetailTab()
    Dim winapi As New clsWinAPI
    With winapi
        .Sleeping 15
        .ActivateATab aHndTXPageControl, 0
        .Sleeping 15
    End With
    Set winapi = Nothing
End Sub

Public Sub EnableGeneralTab()
    Dim winapi As New clsWinAPI
    With winapi
        .Sleeping 15
        .ActivateATab aHndTXPageControl, 3
        .Sleeping 15
    End With
    Set winapi = Nothing
End Sub

Public Sub ClickBtnCreate()
    Dim winapi As New clsWinAPI
    If winapi.IsWinVisible(aHndTabDetail) = True And aHndBtnCreate <> 0 Then
        winapi.PostmsgClick (aHndBtnCreate)
    End If
    Set winapi = Nothing
End Sub

Public Sub ClickBtnDetail()
    Dim winapi As New clsWinAPI
    If winapi.IsWinVisible(aHndTabDetail) = True And aHndBtnDetail <> 0 Then
        winapi.SendmsgClick (aHndBtnDetail)
    End If
    Set winapi = Nothing
End Sub

Public Sub UpdateDate4ChangeOffer()
    Dim winapi As New clsWinAPI, ret As Variant, i As Integer, i2 As Integer, objHwnd As Long
    If aHndClndarStartDate = 0 Or aStartDate = "" Then Exit Sub
    winapi.Sleeping 500
    winapi.ObjSetFocus aHndClndarStartDate
    For i = 0 To 12
        winapi.Sleeping 5
        winapi.VKeyBack aHndClndarStartDate
        winapi.Sleeping 5
    Next
    For i = 1 To Len(aStartDate)
        winapi.KeyTypeIn aHndClndarStartDate, Mid(aStartDate, i, 1), 0, 0
        winapi.Sleeping 15
    Next
    winapi.ObjKillFocus objHwnd
    winapi.Sleeping 500
End Sub
'aCFS4ChangeOffer
Public Sub UpdateCFS4ChangeOffer()
    Dim winapi As New clsWinAPI, ret As Variant, i As Integer, i2 As Integer, objHwnd As Long
    If aHndTxtPSN = 0 Or aCFS4ChangeOffer = "" Then Exit Sub
    winapi.Sleeping 10
    winapi.ObjSetFocus aHndTxtPSN
    winapi.Sleeping 15
    winapi.SendText aHndTxtPSN, aCFS4ChangeOffer
    winapi.Sleeping 15
    winapi.VKeyReturn aHndTxtPSN
    winapi.Sleeping 15
    winapi.ObjKillFocus aHndTxtPSN
    winapi.Sleeping 10
End Sub

Public Sub PutInProductDtails()
    'Product Name // Subtype // Acct ID
    Dim winapi As New clsWinAPI, ret As Variant, i As Integer, i2 As Integer, objHwnd As Long
    Dim SendText As String, DateArr() As String
    If aCFS <> "" Then
        If Left(aCFS, 2) <> "00" Then
            aProductName = "NGA Tail Extension"
        End If
    ElseIf aCFS4ChangeOffer <> "" Then
        If Left(aCFS4ChangeOffer, 2) <> "00" Then
            aProductName = "NGA Tail Extension"
        End If
    End If
    
    With winapi
        .Sleeping 500
        For i = 0 To 4
            If aCFS4ChangeOffer <> "" And i = 4 Then Exit For
            Select Case i
                Case 0
                    objHwnd = aHndClndarStartDate
                    SendText = aStartDate
                    .Sleeping 25
                Case 1
                    objHwnd = aHndLstProductName
                    SendText = aProductName
                    .Sleeping 25
                Case 2
                    objHwnd = aHndLstProductSubTypeHnd
                    SendText = aProductSubType
                    .Sleeping 25
                Case 3
                    objHwnd = aHndTxtPSN
                    If aCFS4ChangeOffer = "" Then
                        SendText = aCFS
                    Else
                        SendText = aCFS4ChangeOffer
                    End If
                    .Sleeping 25
                Case 4
                    objHwnd = aHndLstAcctId
                    .Sleeping 25
            End Select
            If i = 0 Then
                .Sleeping 500
                .ObjSetFocus objHwnd
                For i2 = 0 To 12
                    .Sleeping 5
                    .VKeyBack objHwnd
                    .Sleeping 5
                Next
                For i2 = 1 To Len(SendText)
                    .KeyTypeIn objHwnd, Mid(SendText, i2, 1), 0, 0
                    .Sleeping 15
                Next
                .ObjKillFocus objHwnd
                .Sleeping 500
            End If
            If i > 0 And i < 4 Then
                .Sleeping 10
                .ObjSetFocus objHwnd
                .Sleeping 15
                .SendText objHwnd, SendText
                .Sleeping 15
                .VKeyReturn objHwnd
                .Sleeping 15
                .ObjKillFocus objHwnd
                .Sleeping 10
            Else
                For i2 = 1 To 3
                    If i2 < 3 Then
                        .VKeyDown objHwnd
                    Else
                        .VKeyTab objHwnd
                        .ObjKillFocus objHwnd
                    End If
                Next
            End If
        Next
    End With
    Set winapi = Nothing
End Sub

Public Function GetBillingStatus() As String
    Dim sv As New clsSingleViewMain, tempVar As Variant
    Dim winapi As New clsWinAPI

    With winapi
        Init = sv.mdi
        GetBillingStatus = .GetObjectText(aHndStatus)
    End With
End Function
