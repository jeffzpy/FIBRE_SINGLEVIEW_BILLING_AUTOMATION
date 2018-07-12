VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CaseConnect"
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
Private aInstallChargable As String
Private aEnableOffCharge As Boolean
Private aEnableTesCharge As Boolean
Private aBillingStatus As String
Private aErrCode As String

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

'ChargeCodeArr
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
    Dim i As Integer, i2 As Integer, i3 As Integer, PCSName As String, DigitsCFS As Long
    Dim tempLng As Long, ChargeArr() As String, circuitID As String
    Dim TesCollect As New Collection, TesIdx As Integer
    
    If aCFS <> "" Then
        DigitsCFS = Replace(aCFS, "CFS", "")
        TesCollect.Add (aCFS)
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
        Set sv = New clsSingleViewMain
        Set srch = New objSearchWDs
        Set pi = New objProductInstanceWD
        Set sd = New objServiceDetailWD
        circuitID = TesCollect.item(TesIdx)
        If InStr(circuitID, "CFS") > 0 Then
            circuitID = "00" & DigitsCFS
        End If
        With sv 'Within "clsSingleViewMain" Object code, call below methods, or properties
            svHwnd = .svHwnd
            If svHwnd = 0 Then Exit Sub
            mdi = .mdi
            If mdi = 0 Then Exit Sub
        End With 'End clsSingleViewMain
    
        With srch 'Within "objSearchWDs" Object code, call below methods, or properties
            .SearchAcct aAcctNo
        End With 'End objSearchWDs
    
        With pi 'Within "objProductInstanceWD" Object code, call below methods, or properties
            .CFS = circuitID
            .StartDate = aStartDate
            .AcctId = aAcctNo
            .EnableProductInstanceWindow
            .Init = sv.mdi 'Initializing PI WD object
            .EnableGeneralTab
            .Init = sv.mdi 'PI need to be initialized again duo changing Child Tab Object
            Call .PutInProductDtails
            Call .EnableDetailTab
            Call .ClickBtnCreate 'Need something to check if cfs has been in the sv db
        End With 'End of "objProductInstanceWD"
        winapi.Sleeping 500
        sv.HandlingConfirmWin True 'Click Yes button when Confirm dialog window pop out
        winapi.Sleeping 50
        With sd
            'do while winapi.getChildObjectHwnd(mdi,0, "TfrmServiceDetail",vbnullstring)
            .PortalID = aPortalID
            .CFS = circuitID
            .Asid = aASID
            .ProductPlan = aProductPlan
            .AccessSpeed = aAccessSpeed
            .HighCIR = aHighCir
            .InstallChargable = aInstallChargable
            winapi.Sleeping 1000
            .Init = sv.mdi
            Call .EnableCircuitDetailTab
            .Init = sv.mdi
            .FillCircuitDetails
            .EnableListTab
            .Init = sv.mdi
            Call .ClickPCV
            winapi.Sleeping 2000
            For i = 1 To 5
                Select Case i
                    Case 1
                        PCSName = "PN"
                    Case 2
                        PCSName = "AS"
                    Case 3
                        PCSName = "HCS"
                    Case 4
                        If Left(circuitID, 2) = "00" Then
                            PCSName = "IC"
                        Else
                            PCSName = "TE"
                        End If
                    Case 5
                        PCSName = "IRI"
                End Select
                .ProcessErr = "REPROCESSING"
                Do While .ProcessErr = "REPROCESSING"
                    Call .VerifyPcvGridDataLoaded("PSN")
retryPSN:
                        .Init = sv.mdi
                        .ErrCode = "CSV_CELL_ERROR"
                        Do While .ErrCode = "CSV_CELL_ERROR"
                            winapi.Sleeping 2000
                            Call .PickCFSFromPSNCell
                            If sv.MsgFormChk <> 0 Then
                                Do While sv.MsgFormChk <> 0
                                    sv.MsgFormClick
                                Loop
                                winapi.Sleeping 2000
                                sv.ClickBtnDelete
                                winapi.Sleeping 1000
                                sv.ClickBtnInsert
                                .Click1stCell
                                GoTo retryPSN
                            End If
                        Loop
                    Call .TabtoNextCell
retryPCS:
                    Call .VerifyPcvGridDataLoaded("PCS")
                    winapi.Sleeping 250
                    Call .InsertPCSName(PCSName)
                    If sv.MsgFormChk <> 0 Then
                        Do While sv.MsgFormChk <> 0
                            sv.MsgFormClick
                        Loop
                        winapi.Sleeping 2000
                        sv.ClickBtnDelete
                        winapi.Sleeping 1000
                        sv.ClickBtnInsert
                        .Click1stCell
                        GoTo retryPSN
                    End If
                    'need something to fix the error window
                    Call .TabtoNextCell
retryCV:
                    Call .VerifyPcvGridDataLoaded("CV")
                    Call .InsertCVValue(PCSName)
                    If sv.MsgFormChk <> 0 Then
                        Do While sv.MsgFormChk <> 0
                            sv.MsgFormClick
                        Loop
                        winapi.Sleeping 2000
                        sv.ClickBtnDelete
                        winapi.Sleeping 1000
                        sv.ClickBtnInsert
                        .Click1stCell
                        GoTo retryPSN
                    End If
                    'need something to fix the error window
                    If .ProcessErr = "" Then Call .TabtoNextCell
                    If .ProcessErr = "" Then
                        If PCSName <> "IRI" Then
                            winapi.Sleeping 500
                            Call .ClickBtnInsert
                            winapi.Sleeping 500
                        End If
                    End If
                Loop
            Next
        End With
        With sv 'Within "clsSingleViewMain" Object code, call below methods, or properties
            .ClickBtnSave
            winapi.Sleeping 1000
            Do While .RespondMsg = "NA"
                DoEvents
            Loop
            For i2 = 0 To 11
                If .MsgFormChk <> 0 Then Exit For
                winapi.Sleeping 800
            Next
            For i2 = 0 To 11
                If i2 = 11 Then
                    aErrCode = "CIRCUIT ID#" & TesCollect.item(TesIdx) & " ALREADY EXISTED"
'                    For i3 = 0 To 1
'                        .CloseAllChildWindows
'                        winapi.Sleeping 500
'                        If .MsgFormChk > 0 Then
'                            .HandlingConfirmWin False
'                            winapi.Sleeping 1000
'                        End If
'                    Next
                    Set sv = Nothing
                    Set pi = Nothing
                    Set sd = Nothing
                    Set srch = Nothing
                    Set winapi = Nothing
                    Exit Sub
                End If
                If .MsgFormChk <> 0 Then Exit For
                winapi.Sleeping 2000
            Next
            If InStr(aErrCode, "ALREADY EXISTED") > 0 Then
                For i3 = 0 To 1
                    .CloseAllChildWindows
                    winapi.Sleeping 500
                    If .MsgFormChk > 0 Then
                        .HandlingConfirmWin False
                        winapi.Sleeping 1000
                    End If
                Next
                Exit Sub
            End If
            Call .HandlingConfirmWin(True)
            Do While .MsgFormChk <> 0
                DoEvents
            Loop
            aBillingStatus = pi.GetBillingStatus
            .CloseAllChildWindows
            Do While .MsgFormChk <> 0
                .MsgFormClick
            Loop
            winapi.Sleeping 250
        End With 'End clsSingleViewMain
        circuitID = ""
        Set sv = Nothing
        Set srch = Nothing
        Set pi = Nothing
        Set sd = Nothing
    Next TesIdx
    
'Apply One-Off Charge Below:
    If aEnableOffCharge = True Then
        winapi.Sleeping 500
        Set sv = New clsSingleViewMain
        Set srch = New objSearchWDs
        Set pi = New objProductInstanceWD
        Set sd = New objServiceDetailWD
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
        End If
        sv.CloseAllChildWindows
        On Error GoTo 0
    End If
'Apply One-Off Charge Above
    If Not sv Is Nothing Then Set sv = Nothing
    Set pi = Nothing
    Set sd = Nothing
    Set srch = Nothing
End Sub


