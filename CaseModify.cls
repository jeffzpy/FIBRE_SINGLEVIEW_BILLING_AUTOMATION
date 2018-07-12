VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CaseModify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private aPortalID As Long
Private aOrderType As String
Private aAcctNo As Long
Private aASID As Long
Private aStartDate As String
Private aCFS As String
Private aB1 As String
Private aB2 As String
Private aB3 As String
Private aB4 As String
Private aV1 As String
Private aV2 As String
Private aHighCir As String
Private aAmount As String
Private ChargeCodeArr As String
Private aErrCode As String
Private aEnableOffCharge As Boolean
Private aEnableBusinessPremium As Boolean
Private aEnableTesCharge As Boolean
Private aBillingStatus As String

Public Property Let PortalID(bPortalID As Long)
    aPortalID = bPortalID
End Property

Public Property Let AcctNo(bAcctNo As Long)
    aAcctNo = bAcctNo
End Property

Public Property Let StartDate(bStartDate As String)
    aStartDate = bStartDate
End Property

Public Property Let CFS(bCFS As String)
    aCFS = bCFS
End Property

Public Property Let B1CFS(bCFS As String)
    aB1 = bCFS
End Property

Public Property Let B2CFS(bCFS As String)
    aB2 = bCFS
End Property

Public Property Let B3CFS(bCFS As String)
    aB3 = bCFS
End Property

Public Property Let B4CFS(bCFS As String)
    aB4 = bCFS
End Property

Public Property Let V1CFS(bCFS As String)
    aV1 = bCFS
End Property

Public Property Let V2CFS(bCFS As String)
    aV2 = bCFS
End Property

Public Property Let Asid(bASID As Long)
    aASID = bASID
End Property

Public Property Let HighCIR(bHighCir As String)
    aHighCir = bHighCir
End Property

Public Property Let EnableOffCharge(bEnableOffCharge As Boolean)
    aEnableOffCharge = bEnableOffCharge
End Property

Public Property Let EnableBusinessPremium(bEnableBusinessPremium As Boolean)
    aEnableBusinessPremium = bEnableBusinessPremium
End Property

Public Property Let Amount(bAmount As String)
    If bAmount = "0" Then bAmount = ""
    aAmount = bAmount
End Property

Public Property Let EnableTesCharge(bEnableTesCharge As Boolean)
    aEnableTesCharge = bEnableTesCharge
End Property

Public Property Let ChargeCodeValue(aChargeCodeArr As String)
    ChargeCodeArr = aChargeCodeArr
End Property

Public Property Get BillingStatus() As String
    BillingStatus = aBillingStatus
End Property

Public Property Get ErrCode() As String
    ErrCode = aErrCode
End Property

Public Sub Process()
    Dim sv As New clsSingleViewMain
    Dim winapi As New clsWinAPI
    Dim srch As objSearchWDs
    Dim svHwnd As Long, mdi As Long
    Dim pi As objProductInstanceWD
    Dim sd As objServiceDetailWD
    Dim case_ooc As caseOneOffCharge
    Dim i As Integer, i2 As Integer, PCSName As String, DigitsCFS As Long
    Dim tempLng As Long, ChargeArr() As String, circuitID As String, tempValue As String
    Dim OldCfs As String, OldB1 As String, OldB2 As String, OldB3 As String, OldB4 As String, _
        OldV1 As String, OldV2 As String, ClipBdData As String
    Dim CFS4Loop As String, bFormatCFS As String, aDate As String
    Dim TesCollect As New Collection, TesIdx As Integer
    Dim ErrOccured As Boolean, TesError As String
    
 '=============TES VALIDATION FOR MODIFY ATTRIBUTE=============
    aErrCode = vbNullString
    Set srch = New objSearchWDs
    TesError = vbNullString
    With srch
        .Asid = aASID
        .GetSearchText = True
        .SearchCircuitByAsid
        winapi.Sleeping 1500
        If .ErrCode <> "" Then
            aErrCode = .ErrCode & " ASID:" & aASID
            Set srch = Nothing
            Set sv = Nothing
            Set winapi = Nothing
            Set pi = Nothing
            Set sd = Nothing
            Exit Sub
        End If
        tempValue = .SearchText
        If aB1 <> "" Then TesError = Analysis_Searched_Text(aB1, tempValue)
        If aB2 <> "" Then TesError = Analysis_Searched_Text(aB1, tempValue)
        If aB3 <> "" Then TesError = Analysis_Searched_Text(aB1, tempValue)
        If aB4 <> "" Then TesError = Analysis_Searched_Text(aB1, tempValue)
        If aV1 <> "" Then TesError = Analysis_Searched_Text(aB1, tempValue)
        If aV2 <> "" Then TesError = Analysis_Searched_Text(aB1, tempValue)
        If TesError = "" Then
            If InStr(tempValue, "NGA Tail Extension") > 0 Then
                TesError = "TES_MISSING_PORTAL"
            End If
        End If
    End With
    If TesError = "TES_MISSING_SV" Or TesError = "TES_MISSING_PORTAL" Then
        sv.CloseAllChildWindows
        ErrOccured = True
        aErrCode = TesError
        Set srch = Nothing
        Set winapi = Nothing
        Set sv = Nothing
        Exit Sub
    End If
    sv.CloseAllChildWindows
    Set srch = Nothing
'=================================
    
    TesCollect.Add ("00")
    If aB1 <> "" Then TesCollect.Add (aB1)
    If aB2 <> "" Then TesCollect.Add (aB2)
    If aB3 <> "" Then TesCollect.Add (aB3)
    If aB4 <> "" Then TesCollect.Add (aB4)
    If aV1 <> "" Then TesCollect.Add (aV1)
    If aV2 <> "" Then TesCollect.Add (aV2)
    
    If aEnableBusinessPremium = True Then
        For TesIdx = 1 To TesCollect.Count
            Set srch = New objSearchWDs
            Set pi = New objProductInstanceWD
            Set sd = New objServiceDetailWD
            winapi.Sleeping 1500

            With sv 'Within "clsSingleViewMain" Object code, call below methods, or properties
                svHwnd = .svHwnd
                If svHwnd = 0 Then Exit Sub
                mdi = .mdi
                If mdi = 0 Then Exit Sub
            End With 'sv
            
            With srch
                If IsNumeric(TesCollect.item(TesIdx)) = True Then
                        .Asid = aASID
                        .SearchCircuitByAsid
                        winapi.Sleeping 1500
                        If .ErrCode <> "" Then
                            aErrCode = .ErrCode & " ASID:" & aASID
                            Set srch = Nothing
                            Set sv = Nothing
                            Set winapi = Nothing
                            Set pi = Nothing
                            Set sd = Nothing
                            Exit Sub
                        End If
                        aCFS = .circuitID
                        If Left(aCFS, 1) = "0" And Mid(aCFS, 2, 1) <> "0" Then
                            aCFS = Right(aCFS, Len(aCFS) - 2)
                        Else
                            tempLng = aCFS
                            aCFS = tempLng
                        End If
                Else
                    Select Case Left(TesCollect.item(TesIdx), 2)
                        Case "B1"
                            .circuitID = "B1" & aCFS
                            .circuitID4ChangeOffer = aB1
                        Case "B2"
                            .circuitID = "B2" & aCFS
                            .circuitID4ChangeOffer = aB2
                        Case "B3"
                            .circuitID = "B3" & aCFS
                            .circuitID4ChangeOffer = aB3
                        Case "B4"
                            .circuitID = "B4" & aCFS
                            .circuitID4ChangeOffer = aB4
                        Case "V1"
                            .circuitID = "V1" & aCFS
                            .circuitID4ChangeOffer = aV1
                        Case "V2"
                            .circuitID = "V2" & aCFS
                            .circuitID4ChangeOffer = aV2
                    End Select
                    .SearchCircuitByCircuitID
                    If .ErrCode <> "" Then
                        aErrCode = .ErrCode & " Circuit ID:" & .circuitID
                        ErrOccured = True
                        Set srch = Nothing
                        Set sv = Nothing
                        Set winapi = Nothing
                        Set pi = Nothing
                        Set sd = Nothing
                    End If
                End If
            End With 'srch
            
            With sd
                If ErrOccured = False Then
                    .PortalID = aPortalID
                    .CFS = TesCollect.item(TesIdx)
                    .CFS4ChangeOffer = TesCollect.item(TesIdx)
                    .HighCIR = aHighCir
                    .StartDate = aStartDate
                    
                    winapi.Sleeping 1000
                    .Init = sv.mdi
                    .EnableListTab
                    .Init = sv.mdi
                    Call .ClickPCV
                    If .ErrCode <> "" Then
                        aErrCode = .ErrCode
                        sv.CloseAllChildWindows
                        winapi.Sleeping 1000
                        Exit Sub
                    End If
                    winapi.Sleeping 1000
                    If .ExistingHighCir <> aHighCir Then
                        .SortbyPCS
                        winapi.Sleeping 500
                        sv.ClickBtnUpdate
                        winapi.Sleeping 500
                        aDate = Format(Now(), "dd/mm/yyyy")
                        If aDate <> aStartDate Then
                            Do While sv.MsgFormChk = 0
                                winapi.Sleeping 500
                            Loop
                        End If
                        If sv.MsgFormChk <> 0 Then
                            sv.HandlingConfirmWin False
                        End If
                        .Init = sv.mdi
                        .NavigateToHCIRValue
                        Call .VerifyPcvGridDataLoaded("CV", "HCS")
                        Call .InsertCVValue("HCS")
                        winapi.Sleeping 250
                        Call .TabtoNextCell
                        winapi.Sleeping 250
                        .Init = sv.mdi
                        .UpdateDate
                        sv.ClickBtnSave
                        winapi.Sleeping 3000
                    Else
                        sv.CloseAllChildWindows
                        winapi.Sleeping 500
                        Exit For
                    End If '.ExistingHighCir <> aHighCir
                End If 'ErrOccured = False
            End With 'sd
        Next TesIdx
        sv.CloseAllChildWindows
        winapi.Sleeping 500
    End If 'aEnableBusinessPremium = True
    
'Apply One-Off Charge Below:
    If aEnableOffCharge = True Then
        Set srch = New objSearchWDs
        Set sd = New objServiceDetailWD
        winapi.Sleeping 1500
        With srch
            .Asid = aASID
            .SearchCircuitByAsid
            winapi.Sleeping 1500
            If .ErrCode <> "" Then
                aErrCode = .ErrCode & " Circuit ID:" & aCFS
                Set srch = Nothing
                Set sv = Nothing
                Set winapi = Nothing
                Set pi = Nothing
                Set sd = Nothing
                Exit Sub
            End If
        End With
        sd.Init = sv.mdi
        aBillingStatus = sd.BillingStatus
        On Error Resume Next
        If aBillingStatus = "Active" Then
            ChargeArr = Split(ChargeCodeArr, Chr(10))
            For i = 0 To UBound(ChargeArr)
                Set case_ooc = New caseOneOffCharge
                With case_ooc
                    If Trim(ChargeArr(i)) = "NGA non standard install" Then
                        .Amount = aAmount
                        .OneOffChargeType = 3
                    Else
                        .OneOffChargeType = 2
                    End If
                    .PortalID = aPortalID
                    .Asid = aASID
                    .ReasonCode = Trim(ChargeArr(i))
                    .StartDate = aStartDate
                    .Process sv.mdi
                End With
                Set case_ooc = Nothing
            Next
        Else
            aErrCode = "INACTIVE BILLING STATUS"
        End If
        On Error GoTo 0
        Set case_ooc = Nothing
        Set srch = Nothing
        Set sd = Nothing
    End If
'Apply One-Off Charge Above
    sv.CloseAllChildWindows
    Set sv = Nothing
    Set winapi = Nothing
End Sub

Private Function Analysis_Searched_Text(ByVal aText As String, SearchedText As String) As String
    If InStr(SearchedText, aText) > 0 Then
        Analysis_Searched_Text = "TES_MATCHED"
    Else
        Analysis_Searched_Text = "TES_MISSING_SV"
    End If
End Function




