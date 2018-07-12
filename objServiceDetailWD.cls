VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "objServiceDetailWD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit

Private aHwndSD As Long
Private aHndTXPageControl As Long
Private aHndTabDetail As Long
Private aHndTabComponents As Long
Private aHndTabIndustry As Long
Private aHndTabList As Long
Private aHndtxtSID As Long
Private aHndTxtASID As Long
Private aHndLstType As Long
Private aHndTXPCVGrid As Long
Private ahndTAPCVInplaceGrid As Long
Private ahndBillingStatus As Long
Private ahndBtnProdInstance As Long
Private ahndCalenderEdit As Long
Private aPSN As Long
Private aCV As Long
Private aPCS As Long
Private ahndPSNLookup As Long
Private ahndPCSLookup As Long
Private ahndCVLookup As Long
Private aHndCell As Long

Private aPortalID As Long
Private aCFS As String
Private aStartDate As String
Private aCFS4ChangeOffer As String
Private aASID As Long
Private TesEnabled As Boolean
Private aProductPlan As String
Private aExistingProductPlan As String
Private aAccessSpeed As String
Private aExistingAccessSpeed As String
Private aHighCir As String
Private aExistingHighCir As String
Private aInstallChargable As String
Private aBillingStatus As String
Private aErrCode As String
Private aProcessErr As String

Public Property Let Init(Args As Long)
    Dim HwndSD As Long, TXPanel_D As Long, TXPanel As Long, TXPanelNextSibling As Long, TxpanelSubLevel As Long, _
        TXScrollBox As Long, TXDataEdit As Long, TTabSheet As Long, TADBLookupCombo As Long, TxpanelSubLevelSibling As Long
    Dim winapi As New clsWinAPI, i As Integer, x As Integer
    ahndBillingStatus = 0
    With winapi
        Do While TxpanelSubLevel = 0
            If x = 60 Then Exit Do
            aHwndSD = .getChildObjectHwnd(Args, 0, "TfrmServiceDetail", vbNullString)
            TXPanel = .getChildObjectHwnd(aHwndSD, 0, "TXPanel", vbNullString)
            TxpanelSubLevel = .getChildObjectHwnd(TXPanel, 0, "TXPanel", vbNullString)
            TxpanelSubLevelSibling = .getChildObjectHwnd(TxpanelSubLevel, 0, "TXPanel", vbNullString)
            ahndCalenderEdit = .getChildObjectHwnd(TxpanelSubLevelSibling, 0, "TADBDateTime", vbNullString)
            ahndCalenderEdit = .getChildObjectHwnd(TxpanelSubLevelSibling, ahndCalenderEdit, "TADBDateTime", vbNullString)
            ahndCalenderEdit = .getChildObjectHwnd(ahndCalenderEdit, 0, "TMyAdrockCalendarEdit", vbNullString)
            If ahndBillingStatus = 0 Then
                ahndBillingStatus = .getChildObjectHwnd(TxpanelSubLevel, 0, "TADBLookupCombo", vbNullString)
            End If
            For i = 1 To 2
                TxpanelSubLevel = .getChildObjectHwnd(TXPanel, TxpanelSubLevel, "TXPanel", vbNullString)
            Next
            .Sleeping 500
            x = x + 1
        Loop
        If TxpanelSubLevel <> 0 Then
            TXPanel_D = .getChildObjectHwnd(TxpanelSubLevel, 0, "TXPanel", vbNullString)
            ahndBtnProdInstance = .getChildObjectHwnd(TXPanel_D, 0, "TXBitBtn", "&Product Instance...")
            
            aHndTXPageControl = .getChildObjectHwnd(TXPanel, 0, "TXPageControl", vbNullString)
            aHndTabDetail = .getChildObjectHwnd(aHndTXPageControl, 0, "TTabSheet", "Circuit Details")
            aHndTabList = .getChildObjectHwnd(aHndTXPageControl, 0, "TTabSheet", "&Lists")
            If .IsWinVisible(aHndTabDetail) = True Then
                TXScrollBox = .getChildObjectHwnd(aHndTabDetail, 0, "TXScrollBox", vbNullString)
                aHndLstType = .getChildObjectHwnd(TXScrollBox, 0, "TADBLookupCombo", vbNullString)
                aHndLstType = .getChildObjectHwnd(TXScrollBox, aHndLstType, "TADBLookupCombo", vbNullString)
                For i = 0 To 3
                    If i = 0 Then
                        TXDataEdit = .getChildObjectHwnd(TXScrollBox, 0, "TXDataEdit", vbNullString)
                    End If
                    TXDataEdit = .getChildObjectHwnd(TXScrollBox, TXDataEdit, "TXDataEdit", vbNullString)
                    If i = 1 Then
                        aHndtxtSID = TXDataEdit
                    ElseIf i = 3 Then
                        aHndTxtASID = TXDataEdit
                    End If
                Next i
            End If
            If .IsWinVisible(aHndTabList) = True Then
                TXPanel = .getChildObjectHwnd(aHndTabList, 0, "TXPanel", vbNullString)
                aHndTXPCVGrid = .getChildObjectHwnd(aHndTabList, 0, "TXDataGrid", vbNullString)
                ahndTAPCVInplaceGrid = .getChildObjectHwnd(TXPanel, 0, "TAInplaceGrid", vbNullString)
            End If
        Else
            aErrCode = "UNABLE_LOCATING_SV_OBJECTS"
        End If
    End With
End Property

Public Property Let PortalID(bPortalID As Long)
    aPortalID = bPortalID
End Property

Public Property Let StartDate(bStartDate As String)
    aStartDate = bStartDate
End Property

Public Property Let CFS(bCFS As String)
    aCFS = bCFS
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
'aCFS4ChangeOffer
Public Property Let CFS4ChangeOffer(bCFS4ChangeOffer As String)
    aCFS4ChangeOffer = bCFS4ChangeOffer
End Property

Public Property Get BillingStatus() As String
    Dim winapi As New clsWinAPI
    If ahndBillingStatus <> 0 Then
        aBillingStatus = winapi.GetObjectText(ahndBillingStatus)
    End If
    BillingStatus = aBillingStatus
End Property

Public Property Get HndTAPCVInplaceGrid() As Long
    HndTAPCVInplaceGrid = ahndTAPCVInplaceGrid
End Property

Public Property Get ExistingProductPlan() As String
    If aExistingProductPlan = "" Then Call GetPCVGridData
    ExistingProductPlan = aExistingProductPlan
End Property

Public Property Get ExistingAccessSpeed() As String
    If aExistingAccessSpeed = "" Then Call GetPCVGridData
    ExistingAccessSpeed = aExistingAccessSpeed
End Property

Public Property Get ExistingHighCir() As String
    If aExistingHighCir = "" Then Call GetPCVGridData
    ExistingHighCir = aExistingHighCir
End Property

Public Property Get ErrCode() As String
    ErrCode = aErrCode
End Property

Public Property Let ErrCode(ErrCode As String)
    aErrCode = ErrCode
End Property

Public Property Get ProcessErr() As String
    ProcessErr = aProcessErr
End Property

Public Property Let ProcessErr(bProcessErr As String)
    aProcessErr = bProcessErr
End Property

Public Sub ProcessRowPlanName(ByVal RowChar As String)
    Dim winapi As New clsWinAPI
    Dim sv As New clsSingleViewMain
    Dim svHD As Long, mdi As Long
    svHD = sv.svHwnd
    mdi = sv.mdi
    Init = mdi
    If winapi.IsWinVisible(ahndTAPCVInplaceGrid) = False Then Exit Sub
    If aCFS4ChangeOffer <> "" Then
        VerifyPcvGridDataLoaded "PSN"
        PickCFSFromPSNCell
        TabtoNextCell
        VerifyPcvGridDataLoaded "PCS"
        InsertPCSName RowChar
        TabtoNextCell
    End If
    VerifyPcvGridDataLoaded "CV", RowChar
    InsertCVValue RowChar
    TabtoNextCell
End Sub

Public Sub EnableCircuitDetailTab()
    Dim winapi As New clsWinAPI
    With winapi
        .ActivateATab aHndTXPageControl, 1
    End With
    Set winapi = Nothing
End Sub

Public Sub EnableListTab()
    Dim winapi As New clsWinAPI
    With winapi
        .ActivateATab aHndTXPageControl, 2
    End With
    Set winapi = Nothing
End Sub

Public Sub FillCircuitDetails()
    Dim winapi As New clsWinAPI
    With winapi
        If .IsWinVisible(aHndTabDetail) = False Then Exit Sub
        .SendText aHndTxtASID, CStr(aASID)
        .SendText aHndtxtSID, "SID" & CStr(aASID)
        .ObjSetFocus aHndLstType
        .SendText aHndLstType, "NGA"
        .ObjKillFocus aHndLstType
    End With
    Set winapi = Nothing
End Sub

Public Sub ClickPCV()
    Dim winapi As New clsWinAPI
    Dim sv As New clsSingleViewMain
    Dim i As Integer, i2 As Integer, UpIdx As Integer, ClipBdData As String, clipbddataArr() As String
    If aHndTXPCVGrid = 0 Then Exit Sub
    With winapi
        .BlockInputs True
reTry:
        If aCFS4ChangeOffer = "" Then
            .BkgrndSingleButtonClick aHndTXPCVGrid, 1, 30, 85
            .Sleeping 500
        Else
            Do While InStr(ClipBdData, "DISPLAY_TABLE_NAME") < 1
                .SetClipboard vbNullString
                .BkgrndSingleButtonClick aHndTXPCVGrid, 1, 30, 10
                .Sleeping 1000
                sv.ClickBtnCopy
                .Sleeping 500
                If sv.MsgFormChk > 0 Then
                    If InStr(.GetObjectText(sv.MsgFormChk), "Error") > 0 Then
                        sv.MsgFormClick
                        sv.CloseAllChildWindows
                        aErrCode = "CANNOT OPEN CLIPBOARD WHILE NAVIGATING TO PCV"
                        Set winapi = Nothing: Set sv = Nothing
                        Exit Sub
                    Else
                        sv.MsgFormClick
                    End If
                End If
                ClipBdData = .ClipBoard_GetData
            Loop
                .SetClipboard vbNullString
                .ClearClipboard
                clipbddataArr = Split(ClipBdData, Chr(10))
                For i = 0 To UBound(clipbddataArr)
                    .VKeyDown aHndTXPCVGrid
                    .Sleeping 20
                Next
                For i = 0 To UBound(clipbddataArr)
                    If InStr(clipbddataArr(i), "Product Characteristic Value") > 0 Then
                        UpIdx = UBound(clipbddataArr) - (i + 1)
                        If UpIdx = 0 Then
                            ClipBdData = vbNullString
                            GoTo reTry
                        End If
                        Exit For
                    End If
                Next
            
            ClipBdData = vbNullString
            .Sleeping 1000
            For i = 1 To UpIdx
                .VKeyUp aHndTXPCVGrid
            Next
            .Sleeping 1000
        End If
        .BlockInputs False
    End With
    Set winapi = Nothing: Set sv = Nothing
    i = 0: i2 = 0: UpIdx = 0: ClipBdData = vbNullString: Erase clipbddataArr
End Sub

Private Sub ClickPCVwithValidation()
    Dim winapi As New clsWinAPI
    Call ValidatePCV
    Set winapi = Nothing
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
        .ForegroundWindow ahndTAPCVInplaceGrid
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
        .ActivateMenuObject svHD, "DeleteCtrl+Del", 10, "sendmsg"
        .Sleeping 50
    End With
    Set winapi = Nothing
    Set sv = Nothing
End Sub

Public Sub ClickBtnProdInstance()
    Dim winapi As New clsWinAPI
    With winapi
        .SendmsgClick ahndBtnProdInstance
        .Sleeping 50
    End With
    Set winapi = Nothing
End Sub

Public Sub SortbyPCS()
    Dim winapi As New clsWinAPI
    With winapi
        .BlockInputs True
        If ahndTAPCVInplaceGrid <> 0 Then
            .BkgrndSingleButtonClick ahndTAPCVInplaceGrid, 1, 357, 9
            .Sleeping 250
        End If
        .BlockInputs False
    End With
End Sub

Public Sub UpdateDate()
    Dim winapi As New clsWinAPI
    If ahndCalenderEdit = 0 Then Exit Sub
    With winapi
        .BlockInputs True
        .ObjSetFocus ahndCalenderEdit
        .Sleeping 10
        .SendText ahndCalenderEdit, aStartDate
        .Sleeping 10
        .ObjKillFocus ahndCalenderEdit
        .BlockInputs False
    End With
End Sub

Public Sub InsertPCSName(ByVal Name As String)
    Dim winapi As New clsWinAPI
    Dim LookupCombo As Long, i As Integer
    Dim aStr As String, slicedChar As String
    With winapi
        .BlockInputs True
        Do While aHndCell = 0
            LookupCombo = .getChildObjectHwnd(ahndTAPCVInplaceGrid, 0, "TADBLookupComboInplace", vbNullString)
            aHndCell = LookupCombo
        Loop
        .ObjSetFocus aHndCell
        Select Case Name
            Case "PN"
                aStr = "plan name"
            Case "AS"
                aStr = "access speed"
                .Sleeping 300
            Case "HCS"
                aStr = "high cir speed"
                .Sleeping 400
            Case "IC"
                aStr = "install chargeable"
                .Sleeping 500
            Case "IRI"
                aStr = "internal"
                .Sleeping 600
            Case "TE"
                aStr = "tail"
                .Sleeping 400
        End Select
        .Sleeping 50
        For i = 1 To Len(aStr)
            .Sleeping 10
            slicedChar = Mid(aStr, i, 1)
            If i = 1 Then
                .Sleeping 50
            Else
                .Sleeping 10
            End If
            .KeyTypeIn aHndCell, slicedChar, 0&, 0&
            If i = 1 Then
                .Sleeping 50
            Else
                .Sleeping 10
            End If
        Next
        .ObjKillFocus aHndCell
        .BlockInputs False
    End With
    Set winapi = Nothing
End Sub

Public Sub InsertCVValue(ByVal Name As String)
    Dim winapi As New clsWinAPI
    Dim sv As New clsSingleViewMain
    Dim LookupCombo As Long, i As Integer, slicedChar As String
    Dim aStr As String, i_index As Integer, aRng As String, aOffSetCol As Integer, TesType As String
    With winapi
        .BlockInputs True
        aHndCell = 0
        Do While aHndCell = 0
            LookupCombo = .getChildObjectHwnd(ahndTAPCVInplaceGrid, 0, "TADBLookupComboInplace", vbNullString)
            aHndCell = LookupCombo
        Loop
        .ObjSetFocus aHndCell
        Select Case Name
            Case "PN"
                aStr = aProductPlan
                aRng = "A:A"
                aOffSetCol = 5
            Case "AS"
                aStr = aAccessSpeed
                aRng = "M:M"
                aOffSetCol = 1
                .Sleeping 200
            Case "HCS"
                aStr = aHighCir
                aRng = "P:P"
                aOffSetCol = 1
                .Sleeping 400
            Case "IC"
                aStr = aInstallChargable
                .Sleeping 500
            Case "IRI"
                aStr = aPortalID
                .Sleeping 500
            Case "TE"
                aStr = aCFS
                .Sleeping 300
        End Select
        If Name = "IRI" Then
            For i = 1 To Len(aStr)
                slicedChar = Mid(aStr, i, 1)
                .Sleeping 10
                .KeyTypeIn aHndCell, slicedChar, 0&, 0&
                .Sleeping 10
            Next
        End If
        If Name = "PN" Or Name = "AS" Or Name = "HCS" Then
            i_index = NGA_PARAMETERS.Range(aRng).Find(aStr).OffSet(0, aOffSetCol).Value
            DropPickYes = True
        ElseIf Name = "IC" Then
            If aStr = "True" Then
                i_index = 1
            Else
                i_index = 0
            End If
        End If
        If Name = "IC" Then
            .KeyTypeIn aHndCell, " ", 0&, 0&
            If i_index = 1 Then
                .Sleeping 10
                .VKeyDown aHndCell
                .Sleeping 10
            Else
                .Sleeping 5
                .VKeyDown aHndCell
                .Sleeping 10
                .VKeyUp aHndCell
                .Sleeping 5
            End If
        End If
        If Name = "PN" Or Name = "AS" Or Name = "HCS" Then
            If aCFS4ChangeOffer <> "" Then RestCellCV
            .Sleeping 500
            .KeyTypeIn aHndCell, " ", 0&, 0&
            .Sleeping 400
            If sv.MsgFormChk <> 0 Then
                .Sleeping 1000
                sv.MsgFormClick
                aProcessErr = "REPROCESSING"
                .Sleeping 1000
                sv.ClickBtnDelete
                .Sleeping 2000
                sv.ClickBtnInsert
                .Sleeping 2000
                Click1stCell
                .Sleeping 2000
                Set winapi = Nothing
                Set sv = Nothing
                Exit Sub
            End If
            If i_index = 0 Then
                .Sleeping 10
                .VKeyDown aHndCell
                .Sleeping 5
                .VKeyUp aHndCell
                .Sleeping 10
            End If
            For i = 1 To i_index
                .VKeyDown aHndCell
                .Sleeping 10
            Next
            .VKeyReturn aHndCell
            .Sleeping 25
        End If
        If Name = "TE" Then
            TesType = Left(aCFS, 2)
            .KeyTypeIn aHndCell, " ", 0&, 0&
            .Sleeping 5
            Select Case TesType
                Case "B1"
                    .VKeyDown aHndCell
                    .Sleeping 5
                    .VKeyUp aHndCell
                    .Sleeping 5
                Case "B2"
                    .VKeyDown aHndCell
                    .Sleeping 5
                Case "B3"
                    For i_index = 0 To 1
                        .VKeyDown aHndCell
                        .Sleeping 15
                    Next
                Case "B4"
                    For i_index = 0 To 2
                        .VKeyDown aHndCell
                        .Sleeping 10
                    Next
                Case "V1"
                    For i_index = 0 To 3
                        .VKeyDown aHndCell
                        .Sleeping 10
                    Next
            End Select
        End If
        .ObjKillFocus aHndCell
        .BlockInputs False
    End With
    aProcessErr = ""
    Set winapi = Nothing
End Sub
    
Public Sub PickCFSFromPSNCell()
    Dim winapi As New clsWinAPI
    Dim sv As New clsSingleViewMain
    Dim LookupCombo As Long
    aErrCode = ""
    With winapi
        aHndCell = 0
        Do While aHndCell = 0
            LookupCombo = .getChildObjectHwnd(ahndTAPCVInplaceGrid, 0, "TADBLookupComboInplace", vbNullString)
            aHndCell = LookupCombo
        Loop
        .Sleeping 500
        .ObjSetFocus aHndCell
        .KeyTypeIn aHndCell, " ", 0&, 0&
        .VKeyDownPostWay aHndCell
        If sv.MsgFormChk <> 0 Then
            sv.MsgFormClick
            aErrCode = "CSV_CELL_ERROR"
            'sv.ClickBtnDelete
            .Sleeping 3000
            'sv.ClickBtnInsert
            '.Sleeping 3000
            Click1stCell
            Exit Sub
        End If
        .VKeyReturn aHndCell
        .ObjKillFocus aHndCell
    End With
    Set winapi = Nothing
    Set sv = Nothing
End Sub

Public Sub TabtoNextCell()
    Dim winapi As New clsWinAPI
    Dim LookupCombo As Long
    With winapi
        .BlockInputs True
        Do While aHndCell = 0
            LookupCombo = .getChildObjectHwnd(ahndTAPCVInplaceGrid, 0, "TADBLookupComboInplace", vbNullString)
            aHndCell = LookupCombo
        Loop
        .Sleeping 50
        .VKeyTab aHndCell
        .Sleeping 50
        .BlockInputs False
    End With
    Set winapi = Nothing
End Sub

Public Sub Click1stCell()
    Dim winapi As New clsWinAPI
    Dim LookupCombo As Long
    With winapi
        .Sleeping 250
        .BkgrndSingleButtonClick ahndTAPCVInplaceGrid, 1, 128, 28
        .Sleeping 250
    End With
    Set winapi = Nothing
End Sub

Public Sub NavigateToHCIRValue()
    Dim winapi As New clsWinAPI
    Dim LookupCombo As Long, i As Integer
    With winapi
        .Sleeping 250
        .VKeyDown ahndTAPCVInplaceGrid
        .Sleeping 100
        For i = 0 To 1
            .VKeyTab ahndTAPCVInplaceGrid
            .Sleeping 100
        Next
        .Sleeping 250
    End With
    Set winapi = Nothing
End Sub

Private Sub GetPCVGridData()
    Dim winapi As New clsWinAPI
    Dim sv As New clsSingleViewMain
    Dim i As Integer, i2 As Integer
    Dim TempData As String, TempDataA As String, TempArr() As String
    
    For i = 0 To 10
        If InStr(TempData, "Product Serial Number") > 0 Then Exit For
        Call Click1stCell
        With winapi
            sv.ClickBtnCopy
            .Sleeping 500
            If sv.MsgFormChk > 0 Then
                If InStr(.GetObjectText(sv.MsgFormChk), "Error") > 0 Then
                    sv.MsgFormClick
                    aErrCode = "CANNOT OPEN CLIPBOARD WHILE NAVIGATING TO PCV"
                    Exit Sub
                Else
                    sv.MsgFormClick
                End If
            End If
            TempData = .ClipBoard_GetData
            If InStr(TempData, "Product Serial Number") > 0 Then
                TempArr = Split(TempData, Chr(10))
                For i2 = 0 To UBound(TempArr)
                    If InStr(TempArr(i2), "Plan Name") > 0 Then
                        TempDataA = Trim(Mid(TempArr(i2), InStr(TempArr(i2), "Plan Name") + 10))
                        aExistingProductPlan = Split(TempDataA, " ")(0)
                    End If
                    If InStr(TempArr(i2), "Access Speed") > 0 Then
                        TempDataA = Trim(Mid(TempArr(i2), InStr(TempArr(i2), "MBPS/") - 10))
                        aExistingAccessSpeed = Split(TempDataA, " ")(0)
                    End If
                    If InStr(TempArr(i2), "High CIR Speed") > 0 Then
                        TempDataA = Trim(Mid(TempArr(i2), InStr(TempArr(i2), "MBPS/") - 10))
                        aExistingHighCir = Split(TempDataA, " ")(0)
                    End If
                Next
            End If
        End With
    Next i
End Sub

Sub VerifyPcvGridDataLoaded(ByVal ColPar As String, Optional RowChar As String)
    Dim winapi As New clsWinAPI
    Dim sv As New clsSingleViewMain
    Dim Grid_Coor_Arr() As Variant, Grid_X As Integer, Grid_Y As Integer
    Dim FocusedCell_Arr() As Variant, FocusedCell_X As Integer, FocusedCell_Y As Integer
    Dim x As Integer, Y As Integer
    Dim LookupCombo As Long, LookupCombo1 As Long, LookupCombo2 As Long, LookupCombo3 As Long
    Dim CellPSN As Long, CellPCS As Long, CellCV As Long
    Dim svHD As Long, mdiHD As Long
    
    With winapi
        .MaxmizeWindow aHwndSD
        .Sleeping 150
        '.ForegroundWindow ahndTAPCVInplaceGrid
        If ColPar = "PSN" Then
            .BkgrndSingleButtonClick ahndTAPCVInplaceGrid, 1, 128, 28
        End If
        .Sleeping 200
        For i = 0 To 2
            If i = 0 Then
                LookupCombo = .getChildObjectHwnd(ahndTAPCVInplaceGrid, 0, "TADBLookupComboInplace", vbNullString)
            Else
                LookupCombo = .getChildObjectHwnd(ahndTAPCVInplaceGrid, LookupCombo, "TADBLookupComboInplace", vbNullString)
            End If
            If .IsWinVisible(LookupCombo) = True Then Exit For
        Next
        Select Case ColPar
            Case "PSN"
                WaitCellFocused ahndTAPCVInplaceGrid, LookupCombo, 62
            Case "PCS"
                WaitCellFocused ahndTAPCVInplaceGrid, LookupCombo, 303
            Case "CV"
                WaitCellFocused ahndTAPCVInplaceGrid, LookupCombo, 424
        End Select
        Grid_Coor_Arr = .GetWindowsRect(ahndTAPCVInplaceGrid)
        Grid_X = Grid_Coor_Arr(0)
        Grid_Y = Grid_Coor_Arr(1)
        .ForegroundWindow aHwndSD
        Call WaitListLoaded(Grid_X, Grid_Y, ColPar, RowChar)
    End With
    aHndCell = LookupCombo
    Set winapi = Nothing
End Sub

Function GetPSNValue() As String
    Dim winapi As New clsWinAPI
    Dim TXDataEdit As Long, TXPanel As Long
    With winapi
        TXPanel = .getChildObjectHwnd(aHwndSD, 0, "TXPanel", vbNullString)
        TXPanel = .getChildObjectHwnd(TXPanel, 0, "TXPanel", vbNullString)
        TXDataEdit = .getChildObjectHwnd(TXPanel, 0, "TXDataEdit", vbNullString)
        GetPSNValue = .GetObjectText(TXDataEdit)
    End With
    Set winapi = Nothing
End Function

Function GetServiceStatus() As String
    Dim winapi As New clsWinAPI
    Dim TADBLookupCombo As Long, TXPanel As Long
    With winapi
        TXPanel = .getChildObjectHwnd(aHwndSD, 0, "TXPanel", vbNullString)
        TXPanel = .getChildObjectHwnd(TXPanel, 0, "TXPanel", vbNullString)
        TADBLookupCombo = .getChildObjectHwnd(TXPanel, 0, "TADBLookupCombo", vbNullString)
        GetServiceStatus = .GetObjectText(TADBLookupCombo)
    End With
    Set winapi = Nothing
End Function

Sub WaitSaveCompleted4ChangeOffer()
    Dim winapi As New clsWinAPI
    Dim TADBLookupCombo As Long, TXPanel As Long
    Dim pLocation As Point
    Dim lDC As Variant
    Dim lColour As Long
    Dim px As Variant
    Dim x As Long, Y As Long, i As Long, xyArr() As Variant
    With winapi
        If aHwndSD = 0 Then
            .Sleeping 8000
            Exit Sub
        End If
        TXPanel = .getChildObjectHwnd(aHwndSD, 0, "TXPanel", vbNullString)
        TXPanel = .getChildObjectHwnd(TXPanel, 0, "TXPanel", vbNullString)
        TADBLookupCombo = .getChildObjectHwnd(TXPanel, 0, "TADBLookupCombo", vbNullString)
        xyArr = .GetWindowsRect(TADBLookupCombo)
        x = xyArr(0): Y = xyArr(1)
        px = -1
        i = 1
        Do While px <> 16777215
            .ForegroundWindow aHwndSD
            If i = 15 Then Exit Do
            lDC = .GetWinDC(0)
            px = .GetPix(lDC, x + 214, Y + 10)
            i = i + 1
            .Sleeping 1000
        Loop
        .Sleeping 1000
    End With
    Set winapi = Nothing
End Sub


Private Sub WaitCellFocused(ByVal ahndTAPCVInplaceGrid As Long, LookupCombo As Long, Condition As Integer)
    Dim winapi As New clsWinAPI
    Dim Grid_Coor_Arr() As Variant, Grid_X As Integer, Grid_Y As Integer
    Dim FocusedCell_Arr() As Variant, FocusedCell_X As Integer, FocusedCell_Y As Integer
    With winapi
        Do While FocusedCell_X - Grid_X <> Condition
            .Sleeping 150
            .ForegroundWindow ahndTAPCVInplaceGrid
            .Sleeping 150
            Grid_Coor_Arr = .GetWindowsRect(ahndTAPCVInplaceGrid)
            FocusedCell_Arr = .GetWindowsRect(LookupCombo)
            Grid_X = Grid_Coor_Arr(0)
            Grid_Y = Grid_Coor_Arr(1)
            FocusedCell_X = FocusedCell_Arr(0)
            FocusedCell_Y = FocusedCell_Arr(1)
            If aCFS4ChangeOffer <> "" Then
            FocusedCell_X = Condition
            Grid_X = 0
            End If
        Loop
        .Sleeping 150
    End With
    Set winapi = Nothing
End Sub

Private Sub WaitListLoaded(ByVal Grid_X As Integer, Grid_Y As Integer, ColPar As String, Optional RowChar As String)
    Dim winapi As New clsWinAPI
    Dim pLocation As Point
    Dim lDC As Variant
    Dim lColour As Long
    Dim px As Variant
    Dim x As Long, Y As Long, i As Long
    'TesEnabled
    Y = Grid_Y + 26
    Select Case ColPar
        Case "PSN"
            x = Grid_X + 293
        Case "PCS"
            x = Grid_X + 413
        Case "CV"
            x = Grid_X + 535
            Select Case RowChar
                Case "HCS"
                    Y = Grid_Y + 44
                Case "IC"
                    Y = Grid_Y + 62
                Case "IRI"
                    Y = Grid_Y + 80
                Case "PN"
                    Y = Grid_Y + 98
            End Select
    End Select
    px = -1
    i = 1
    Do While px <> 0
        If px = 0 Then Exit Do
        If px = 13160660 Then Exit Do
        If i = 30 Then Exit Do
        winapi.Sleeping 250
        lDC = winapi.GetWinDC(0)
        px = winapi.GetPix(lDC, x, Y)
        i = i + 1
    Loop
    winapi.Sleeping 250
    Set winapi = Nothing
End Sub

Private Sub FillPCVRow()
    Dim winapi As New clsWinAPI
    With winapi
        .GetWindowsRect ahndTAPCVInplaceGrid
    End With
    Set winapi = Nothing
End Sub

Private Sub ValidatePCV()
    Dim winapi As New clsWinAPI, GridValue As String
    Dim sv As New clsSingleViewMain
    Dim i As Integer, x As Integer, Y As Integer, OffSet As Integer
    If aHndTXPCVGrid = 0 Then Exit Sub
    OffSet = 17
    x = 33
    Y = 145
    With winapi
        .SetClipboard "a"
        .Sleeping 1000
        .VkeyCtrlC aHndTXPCVGrid
        .Sleeping 500
        If sv.MsgFormChk > 0 Then
            If InStr(.GetObjectText(sv.MsgFormChk), "Error") > 0 Then
                sv.MsgFormClick
                aErrCode = "CANNOT OPEN CLIPBOARD WHILE NAVIGATING TO PCV"
                Exit Sub
            Else
                sv.MsgFormClick
            End If
        End If
        Do While GridValue = "" Or GridValue = "ERR_CB_LOCKED"
            GridValue = .ClipBoard_GetData
        Loop
        For i = 0 To 8
            If i = 8 Then
                i = 0
                Y = 145
            End If
            If GridValue = "Product Characteristic Value" Then Exit Sub
            If i = 0 Then
                OffSet = 0
            Else
                OffSet = 17
                Y = Y - OffSet
            End If
            .BkgrndSingleButtonClick aHndTXPCVGrid, 1, x, Y
            .VkeyCtrlC aHndTXPCVGrid
            GridValue = .ClipBoard_GetData
        Next
    End With
    Set winapi = Nothing
End Sub

Private Sub RestCellCV()
    Dim winapi As New clsWinAPI
    Dim i As Integer
    If aHndCell = 0 Then Exit Sub
    winapi.ObjSetFocus aHndCell
    winapi.KeyTypeIn aHndCell, " ", 0&, 0&
    For i = 0 To 150
        winapi.VKeyUp aHndCell
        winapi.Sleeping 3
    Next
End Sub

