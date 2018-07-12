VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CaseChangeOffer"
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
Private aProductPlan As String
Private aAccessSpeed As String
Private aHighCir As String
Private aAmount As String
Private ChargeCodeArr As String
Private aErrCode As String
Private aInstallChargable As String
Private aEnableOffCharge As Boolean
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

Public Property Let ProductPlan(bProductPlan As String)
    aProductPlan = bProductPlan
End Property

Public Property Let AccessSpeed(bAccessSpeed As String)
    aAccessSpeed = bAccessSpeed
End Property

Public Property Let HighCIR(bHighCir As String)
    aHighCir = bHighCir
End Property

Public Property Let InstallChargable(bInstallChargable As String)
    aInstallChargable = bInstallChargable
End Property

Public Property Let EnableOffCharge(bEnableOffCharge As Boolean)
    aEnableOffCharge = bEnableOffCharge
End Property

Public Property Let EnableTesCharge(bEnableTesCharge As Boolean)
    aEnableTesCharge = bEnableTesCharge
End Property

Public Property Let Amount(bAmount As String)
    If bAmount = "0" Then bAmount = ""
    aAmount = bAmount
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
    Dim sv As clsSingleViewMain
    Dim winapi As New clsWinAPI
    Dim srch As objSearchWDs
    Dim svHwnd As Long, mdi As Long
    Dim pi As objProductInstanceWD
    Dim sd As objServiceDetailWD
    Dim case_ooc As caseOneOffCharge
    Dim i As Integer, i2 As Integer, PCSName As String, DigitsCFS As Long
    Dim tempLng As Long, ChargeArr() As String, circuitID As String, tempValue As String
    Dim OldCfs As String, OldB1 As String, OldB2 As String, OldB3 As String, OldB4 As String, _
        OldV1 As String, OldV2 As String
    Dim CFS4Loop As String, bFormatCFS As String, aDate As String
    Dim TesCollect As New Collection, TesIdx As Integer
    Dim ErrOccured As Boolean
    If aCFS <> "" Then
        DigitsCFS = Replace(aCFS, "CFS", "")
        tempLng = DigitsCFS
        CFS4Loop = "00" & tempLng
        TesCollect.Add (CFS4Loop)
    End If
    If aB1 <> "" Then
        aB1 = "B1" & DigitsCFS
        TesCollect.Add (aB1)
    End If
    If aB2 <> "" Then
        aB2 = "B2" & DigitsCFS
        TesCollect.Add (aB2)
    End If
    If aB3 <> "" Then
        aB3 = "B3" & DigitsCFS
        TesCollect.Add (aB3)
    End If
    If aB4 <> "" Then
        aB4 = "B4" & DigitsCFS
        TesCollect.Add (aB4)
    End If
    If aV1 <> "" Then
        aV1 = "V1" & DigitsCFS
        TesCollect.Add (aV1)
    End If
    If aV2 <> "" Then
        aV2 = "V2" & DigitsCFS
        TesCollect.Add (aV2)
    End If
    For TesIdx = 1 To TesCollect.Count
        ErrOccured = False
        Set sv = New clsSingleViewMain
        Set srch = New objSearchWDs
        Set pi = New objProductInstanceWD
        Set sd = New objServiceDetailWD
        With sv 'Within "clsSingleViewMain" Object code, call below methods, or properties
            svHwnd = .svHwnd
            If svHwnd = 0 Then Exit Sub
            mdi = .mdi
            If mdi = 0 Then Exit Sub
        End With 'End clsSingleViewMain
    
        With srch 'Within "objSearchWDs" Object code, call below methods, or properties
            If IsNumeric(TesCollect.item(TesIdx)) = True Then
                .Asid = aASID
                .SearchCircuitByAsid
                If .ErrCode <> "" Then
                    aErrCode = .ErrCode & " Circuit ID:" & TesCollect.item(TesIdx)
                    Set srch = Nothing
                    Set sv = Nothing
                    Set winapi = Nothing
                    Set pi = Nothing
                    Set sd = Nothing
                    Exit Sub
                End If
                OldCfs = .circuitID
                If Left(OldCfs, 1) = "0" And Mid(OldCfs, 2, 1) <> "0" Then
                    OldCfs = Right(OldCfs, Len(OldCfs) - 2)
                Else
                    tempLng = OldCfs
                    OldCfs = tempLng
                End If
            Else
                Select Case Left(TesCollect.item(TesIdx), 2)
                    Case "B1"
                        .circuitID = "B1" & OldCfs
                        .circuitID4ChangeOffer = aB1
                    Case "B2"
                        .circuitID = "B2" & OldCfs
                        .circuitID4ChangeOffer = aB2
                    Case "B3"
                        .circuitID = "B3" & OldCfs
                        .circuitID4ChangeOffer = aB3
                    Case "B4"
                        .circuitID = "B4" & OldCfs
                        .circuitID4ChangeOffer = aB4
                    Case "V1"
                        .circuitID = "V1" & OldCfs
                        .circuitID4ChangeOffer = aV1
                    Case "V2"
                        .circuitID = "V2" & OldCfs
                        .circuitID4ChangeOffer = aV2
                End Select
                .SearchCircuitByCircuitID
                winapi.Sleeping 1000
                If .ErrCode <> "" Then
                    aErrCode = .ErrCode & " Circuit ID:" & TesCollect.item(TesIdx)
                    ErrOccured = True
                    Set srch = Nothing
                    Set sv = Nothing
                    Set winapi = Nothing
                    Set pi = Nothing
                    Set sd = Nothing
                End If
            End If
        End With 'End objSearchWDs
        With sd
            If ErrOccured = False Then
                .Init = sv.mdi
                .ClickBtnProdInstance
                winapi.Sleeping 1000
            End If
        End With
        
        With pi
            If ErrOccured = False Then
                .Init = sv.mdi
                .EnableGeneralTab
                sv.ClickBtnUpdate
                winapi.Sleeping 250
                aDate = Format(Now(), "dd/mm/yyyy")
                If aStartDate <> aDate Then
                    Do While sv.MsgFormChk = 0
                        winapi.Sleeping 500
                    Loop
                End If
                If sv.MsgFormChk <> 0 Then
                    sv.HandlingConfirmWin False
                End If
                .StartDate = aStartDate
                .CFS4ChangeOffer = TesCollect.item(TesIdx)
                .Init = sv.mdi
                .UpdateDate4ChangeOffer
                'sv.ClickBtnSave
                winapi.Sleeping 500
                'sv.ClickBtnUpdate
                '.Init = sv.mdi
                .UpdateCFS4ChangeOffer
                sv.ClickBtnSave
                winapi.Sleeping 3000
                sv.CloseAllChildWindows
                winapi.Sleeping 2000
            End If
        End With
        
        With srch
            If ErrOccured = False Then
                .circuitID = TesCollect.item(TesIdx)
                .circuitID4ChangeOffer = TesCollect.item(TesIdx)
                .SearchCircuitByCircuitID
                If .ErrCode <> "" Then
                        aErrCode = .ErrCode & " Circuit ID:" & TesCollect.item(TesIdx)
                        Set srch = Nothing
                        Set sv = Nothing
                        Set winapi = Nothing
                        Set pi = Nothing
                        Set sd = Nothing
                        ErrOccured = True
                End If
                winapi.Sleeping 1000
            End If
        End With
        
        With sd
            If ErrOccured = False Then
                'do while winapi.getChildObjectHwnd(mdi,0, "TfrmServiceDetail",vbnullstring)
                .PortalID = aPortalID
                .CFS = TesCollect.item(TesIdx)
                .CFS4ChangeOffer = TesCollect.item(TesIdx)
                .ProductPlan = aProductPlan
                .AccessSpeed = aAccessSpeed
                .HighCIR = aHighCir
                .StartDate = aStartDate
                .InstallChargable = aInstallChargable
                
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
                winapi.Sleeping 1500
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
                .SortbyPCS
                .Init = sv.mdi
                For i = 1 To 4
                    Select Case i
                        Case 1
                            PCSName = "AS"
                        Case 2
                            PCSName = "HCS"
                        Case 3
                            PCSName = "IRI"
                        Case 4
                            PCSName = "PN"
                    End Select
                    If PCSName = "AS" Then
                        For i2 = 0 To 1
                            Call .TabtoNextCell
                        Next
                    End If
                    Call .VerifyPcvGridDataLoaded("CV", PCSName)
                    Call .InsertCVValue(PCSName)
                    winapi.Sleeping 500
                    Call .TabtoNextCell
                    winapi.Sleeping 500
                    If i > 0 And i < 4 Then
                        winapi.BlockInputs True
                        If IsNumeric(TesCollect.item(TesIdx)) = True And PCSName = "HCS" Then
                            For i2 = 0 To 1
                                winapi.Sleeping 100
                                winapi.VKeyDown .HndTAPCVInplaceGrid
                                winapi.Sleeping 100
                            Next
                            winapi.Sleeping 100
                            winapi.VKeyLft .HndTAPCVInplaceGrid
                            winapi.Sleeping 100
                        Else
                            winapi.Sleeping 100
                            winapi.VKeyDown .HndTAPCVInplaceGrid
                            winapi.Sleeping 100
                            winapi.VKeyLft .HndTAPCVInplaceGrid
                            winapi.Sleeping 100
                        End If
                        winapi.BlockInputs False
                    End If
                Next i
            End If
            winapi.Sleeping 1000
        End With
        With sv 'Within "clsSingleViewMain" Object code, call below methods, or properties
            If ErrOccured = False Then
                .ClickBtnSave
                Do While .RespondMsg = "NA"
                    DoEvents
                Loop
                sd.WaitSaveCompleted4ChangeOffer
                sd.Init = sv.mdi
                aBillingStatus = sd.BillingStatus
                .CloseAllChildWindows
                i = 0
                Do While .MsgFormChk <> 0
                    If i = 10 Then Exit Do
                    .MsgFormClick
                    i = i + 1
                    winapi.Sleeping 1000
                Loop
                winapi.Sleeping 1000
            End If
        End With 'End clsSingleViewMain
        Set sv = Nothing
        Set srch = Nothing
        Set pi = Nothing
        Set sd = Nothing
        winapi.Sleeping 1500
    Next TesIdx
    
'Apply One-Off Charge Below:
    If aEnableOffCharge = True Then
        Set sv = New clsSingleViewMain
        Set srch = New objSearchWDs
        Set sd = New objServiceDetailWD
        winapi.Sleeping 1500
        With srch
            .circuitID = "00" & DigitsCFS
            .SearchCircuitByCircuitID
            winapi.Sleeping 500
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
        sv.CloseAllChildWindows
        Set sv = Nothing
        Set case_ooc = Nothing
        Set srch = Nothing
        Set sd = Nothing
    End If
'Apply One-Off Charge Above
    If Not sv Is Nothing Then Set sv = Nothing
    If Not winapi Is Nothing Then Set winapi = Nothing
End Sub



