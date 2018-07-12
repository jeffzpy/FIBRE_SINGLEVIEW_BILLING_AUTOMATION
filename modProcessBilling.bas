Attribute VB_Name = "modProcessBilling"
Option Explicit

Sub ProcessBilling()
    Shell ("taskkill /f /im iexplore.exe /t")
    Sleep 1000
    Dim cc As CaseConnect
    Dim disc As CaseDisconnect
    Dim co As CaseChangeOffer
    Dim trsfr As CaseTransfer
    Dim cm As CaseModify
    Dim lstRow As Long, row As Long, i As Integer, i2 As Integer, tempValue As String
    Dim ie As New InternetExplorer
    Dim sv As New clsSingleViewMain
    Dim winapi As New clsWinAPI
    Dim PortalID As Long, AcctNo As Long, Asid As Long, INTERNALID As Long, ordetype As String
    Dim JobType As String, StartDate As String, CFS As String, ProductPlan As String, _
        AccessSpeed As String, HighCIR As String, InstChargeable As String, charge As String, _
        OrderStatus As String
    Dim TesB1 As String, TesB2 As String, TesB3 As String, TesB4 As String, TesV1 As String, TesV2 As String
    Dim tesarr() As String, ErrCode As String
    If sv.svHwnd = 0 Then
        MsgBox "SingleView is not detected; Please make sure both Macro file & SV sitting in the same Citrix Session"
        Exit Sub
    End If
    lstRow = DATASHEET.Range("A" & Rows.Count).End(xlUp).row
    
    ie.Toolbar = 0
    ie.Visible = True
    If FuncLogin(ie, DASHBOARD.Range("Q13"), DASHBOARD.Range("Q14")) = False Then
        Set cc = Nothing: Set disc = Nothing: Set co = Nothing: Set trsfr = Nothing: Set cm = Nothing
        ie.Quit
        Set ie = Nothing
        Set sv = Nothing
        MsgBox "Issue with login into Portal"
        Exit Sub
    End If
    tempValue = ""
    On Error Resume Next
    For i = 0 To 10
        If i = 10 Then Exit For
        tempValue = ie.Document.URL
        If tempValue = "https://portal.chorus.co.nz/chorus-ssp-web/pages/LandingPage" Then Exit For
        Sleep 1000
    Next
    If tempValue <> "https://portal.chorus.co.nz/chorus-ssp-web/pages/LandingPage" Then
        Set cc = Nothing: Set disc = Nothing: Set co = Nothing: Set trsfr = Nothing: Set cm = Nothing
        ie.Quit
        Set ie = Nothing
        Set sv = Nothing
        MsgBox "Issue with login into Portal"
        Exit Sub
    End If
    On Error GoTo 0
    If sv.MsgFormChk <> 0 Then
        sv.MsgFormClick
    End If
    Sleep 3000
    ChorusPortalController.BrowseCP ie, "Search"
    For row = 2 To lstRow
        If DATASHEET.Range("F" & row).Value <> "ASID not found in CFS report" Then
            TimeStart = Format(Now(), "hh:mm:ss")
            Set disc = New CaseDisconnect
            Set trsfr = New CaseTransfer
            Set cc = New CaseConnect
            Set co = New CaseChangeOffer
            Set cm = New CaseModify
            PortalID = DATASHEET.Range("A" & row).Value
            JobType = DATASHEET.Range("B" & row).Value
            AcctNo = DATASHEET.Range("C" & row).Value
            StartDate = DATASHEET.Range("D" & row).Value
            Asid = DATASHEET.Range("E" & row).Value
            CFS = DATASHEET.Range("F" & row).Value
            ProductPlan = DATASHEET.Range("G" & row).Value
            AccessSpeed = DATASHEET.Range("H" & row).Value
            HighCIR = DATASHEET.Range("I" & row).Value
            InstChargeable = DATASHEET.Range("J" & row).Value
            OrderStatus = DATASHEET.Range("M" & row).Value
            INTERNALID = PortalID
            sv.InitialSizeOfSV
            If ifPortalIDExsit(row) = True Then
                OrderStatus = ""
                DATASHEET.Range("N" & row).Value = "SOMEONE ELSE PROCESSED THIS ORDER ALREADY"
                DATASHEET.Range("A" & row & ":AB" & row).Interior.Color = 10498160
            End If
            If OrderStatus = "POSTED" Or OrderStatus = "BILLING STAGE" Then
                If Left(JobType, 7) = "Connect" Then ordetype = "Connect"
                If Left(JobType, 8) = "Transfer" Then ordetype = "Transfer"
                If Left(JobType, 10) = "Disconnect" Then ordetype = "Disconnect"
                If JobType = "Change Offer" Then ordetype = "Change"
                If JobType = "Modify Attribute" Then ordetype = "Modify"
                cc.EnableTesCharge = False
                co.EnableTesCharge = False
                cm.EnableTesCharge = False
                If DATASHEET.Range("AB" & row).Value <> "" Then
                    cc.EnableTesCharge = True
                    co.EnableTesCharge = True
                    cm.EnableTesCharge = True
                    tesarr = Split(DATASHEET.Range("AB" & row).Value, "/")
                    For i = LBound(tesarr) To UBound(tesarr)
                        tempValue = tesarr(i)
                        Select Case tempValue
                            Case "B1"
                                TesB1 = tempValue
                            Case "B2"
                                TesB2 = tempValue
                            Case "B3"
                                TesB3 = tempValue
                            Case "B4"
                                TesB4 = tempValue
                            Case "V1"
                                TesV1 = tempValue
                            Case "V2"
                                TesV2 = tempValue
                        End Select
                    Next
                Else
                    cc.EnableTesCharge = False
                End If
                If DATASHEET.Range("R" & row).Value = "" Then
                    cc.EnableOffCharge = False: co.EnableOffCharge = False: cm.EnableOffCharge = False
                Else
                    cc.EnableOffCharge = True: co.EnableOffCharge = True: cm.EnableOffCharge = True
                    cc.ChargeCodeValue = DATASHEET.Range("R" & row).Value
                    co.ChargeCodeValue = DATASHEET.Range("R" & row).Value
                    cm.ChargeCodeValue = DATASHEET.Range("R" & row).Value
                End If
                charge = DATASHEET.Range("Z" & row).Value
                'ie.Visible = False
                If (InStr(ProductPlan, "Chorus Better Broadband") > 0) Or (InStr(ProductPlan, "Chorus Internal") > 0) Then DATASHEET.Range("N" & row).Value = "Completed"
                If DATASHEET.Range("N" & row).Value <> "Completed" Then
                    Select Case ordetype
                        Case "Connect"
                            With cc
                                .AccessSpeed = AccessSpeed
                                .AcctNo = AcctNo
                                .Asid = Asid
                                .CFS = CFS
                                .B1CFS = TesB1
                                .B2CFS = TesB2
                                .B3CFS = TesB3
                                .B4CFS = TesB4
                                .V1CFS = TesV1
                                .V2CFS = TesV2
                                .HighCIR = HighCIR
                                .InstallChargable = InstChargeable
                                .PortalID = PortalID
                                .ProductPlan = ProductPlan
                                .StartDate = StartDate
                                .Amount = charge
                                Call .Process: ErrCode = .ErrCode
                                DATASHEET.Range("O" & row).Value = .BillingStatus
                            End With
                        Case "Disconnect"
                            With disc
                                .StartDate = StartDate
                                .Asid = Asid
                                .B1CFS = TesB1
                                .B2CFS = TesB2
                                .B3CFS = TesB3
                                .B4CFS = TesB4
                                .V1CFS = TesV1
                                .V2CFS = TesV2
                                .Process: ErrCode = .ErrCode
                                DATASHEET.Range("O" & row).Value = .BillingStatus
                            End With
                        Case "Change"
                            With co
                                .ProductPlan = ProductPlan
                                .AccessSpeed = AccessSpeed
                                .HighCIR = HighCIR
                                .PortalID = PortalID
                                .Asid = Asid
                                .CFS = CFS
                                .StartDate = StartDate
                                .Amount = charge
                                .B1CFS = TesB1
                                .B2CFS = TesB2
                                .B3CFS = TesB3
                                .B4CFS = TesB4
                                .V1CFS = TesV1
                                .V2CFS = TesV2
                                .Process: ErrCode = .ErrCode
                                DATASHEET.Range("O" & row).Value = .BillingStatus
                            End With
                        Case "Modify"
                            With cm
                                .PortalID = PortalID
                                .AcctNo = AcctNo
                                .Amount = charge
                                .Asid = Asid
                                .CFS = CFS
                                .B1CFS = TesB1
                                .B2CFS = TesB2
                                .B3CFS = TesB3
                                .B4CFS = TesB4
                                .V1CFS = TesV1
                                .V2CFS = TesV2
                                .HighCIR = HighCIR
                                .StartDate = StartDate
                                If DATASHEET.Range("AA" & row).Value = "Business Premium" Then
                                    .EnableBusinessPremium = True
                                End If
                                .Process: ErrCode = .ErrCode
                                DATASHEET.Range("O" & row).Value = .BillingStatus
                            End With
                        Case "Transfer"
                            With disc
                                .StartDate = StartDate
                                .Asid = DATASHEET.Range("W" & row).Value
                                .B1CFS = TesB1
                                .B2CFS = TesB2
                                .B3CFS = TesB3
                                .B4CFS = TesB4
                                .V1CFS = TesV1
                                .V2CFS = TesV2
                                .Process: ErrCode = .ErrCode
                            End With
                            Sleep 2000
                                                        'sv.InitialSizeOfSV
                            With cc
                                .AccessSpeed = AccessSpeed
                                .AcctNo = AcctNo
                                .Asid = Asid
                                .CFS = CFS
                                .B1CFS = TesB1
                                .B2CFS = TesB2
                                .B3CFS = TesB3
                                .B4CFS = TesB4
                                .V1CFS = TesV1
                                .V2CFS = TesV2
                                .HighCIR = HighCIR
                                .InstallChargable = InstChargeable
                                .PortalID = PortalID
                                .ProductPlan = ProductPlan
                                .StartDate = StartDate
                                .Amount = charge
                                Call .Process
                                If ErrCode = "" Then
                                    ErrCode = .ErrCode
                                Else
                                    ErrCode = ErrCode & Chr(10) & .ErrCode
                                End If
                                DATASHEET.Range("O" & row).Value = .BillingStatus
                            End With
                    End Select
                    If ErrCode <> "" Then
                        With DATASHEET.Range("N" & row)
                            .Value = ErrCode
                            .Interior.Color = 255
                        End With
                        If ErrCode = "TES_MISSING_SV" Then DATASHEET.Range("N" & row).Value = "TES Billing Record Missing in SV"
                        If ErrCode = "TES_MISSING_PORTAL" Then DATASHEET.Range("N" & row).Value = "TES Billing Mismatch between Portal & SV"
                        If ErrCode = "CANNOT OPEN CLIPBOARD WHILE NAVIGATING TO PCV" Then
                            MsgBox "It seems that there`s an issue with system Clipboard function; Better restarting your VDX & CITRIX & VDI before re-run this macro"
                            Exit Sub
                        End If
                    End If
                End If
                'ie.Visible = True
                winapi.Sleeping 1000
                Procedure_ChorusPortal.complete_perform_billing_task row, ie
                DATASHEET.Range("A" & row & ":" & "AB" & row).Select
                If ErrCode = "" Then
                    With Selection.Interior
                        .Color = 5287936
                    End With
                End If
                'ie.Visible = False
                If DATASHEET.Range("N" & row).Value = "" Then
                    DATASHEET.Range("N" & row).Value = "Completed"
                End If
                Module_Database.InsertData row
                On Error Resume Next
                ThisWorkbook.Save
                Err.Clear
                On Error GoTo 0
            End If
            Sleep 1000
            Set disc = Nothing
            Sleep 100
            Set trsfr = Nothing
            Sleep 100
            Set cc = Nothing
            Sleep 100
            Set co = Nothing
            Sleep 100
            Set cm = Nothing
            Sleep 100
            Erase tesarr()
            Sleep 500
            PortalID = 0: AcctNo = 0: Asid = 0: INTERNALID = 0: OrderStatus = vbNullString
            ordetype = vbNullString: JobType = vbNullString: StartDate = vbNullString: CFS = vbNullString: ProductPlan = vbNullString
            AccessSpeed = vbNullString: HighCIR = vbNullString: InstChargeable = vbNullString: charge = vbNullString: tempValue = vbNullString
            TesB1 = vbNullString: TesB2 = vbNullString: TesB3 = vbNullString: TesB4 = vbNullString: TesV1 = vbNullString: TesV2 = vbNullString
            TimeFinish = Format(Now(), "hh:mm:ss")
            ErrCode = "": TimeStart = "": TimeFinish = ""
            TimeFinish = Format(Now(), "hh:mm:ss")
            Application.wait (Now + TimeValue("00:00:01"))
        End If
    Next row
    ie.Quit
    On Error Resume Next
    Err.Clear
    ThisWorkbook.Save
    If Err.Number <> 0 Then
        For i2 = 0 To 20
            If i2 = 20 Then Exit For
            If Err.Number = 0 Then Exit For
            Err.Clear
            Application.wait (Now + TimeValue("00:00:01"))
            ThisWorkbook.Save
        Next
    End If
    On Error GoTo 0
    MsgBox "Process Completed"
End Sub

Sub OpenSV()
    Dim hWnd As Long, hwndlogin As Long, Explorer_SV As Long
    Dim winapi As New clsWinAPI
    Dim Explorer_Shell_SV As Long, Explorer_Shell_DUI_SV As Long, Explorer_Shell_DUI_UIHWND_SV As Long, _
        shelldll As Long, CNS As Long, tNo As String
    Dim i As Integer
    With winapi
        hWnd = 0
        hWnd = .getObjectHwnd("TfrmMain", vbNullString)
        If hWnd = 0 Then
            BlockInput True
            Shell "C:\Windows\explorer.exe /select, C:\Program Files (x86)\Singleview 9.00.15.01 Prod\Billing.bat", vbNormalFocus
            Explorer_SV = 0
            Do While Explorer_SV = 0
                Explorer_SV = .getObjectHwnd("CabinetWClass", "Singleview 9.00.15.01 Prod")
            Loop
            SetWindowPos Explorer_SV, HWND_TOP1, 0, 0, 100, 100, SWP_SHOWWINDOW
            Explorer_Shell_SV = FindWindowExA(Explorer_SV, 0, "ShellTabWindowClass", vbNullString)
            Explorer_Shell_DUI_SV = FindWindowExA(Explorer_Shell_SV, 0, "DUIViewWndClassName", vbNullString)
            Explorer_Shell_DUI_UIHWND_SV = FindWindowExA(Explorer_Shell_DUI_SV, 0, "DirectUIHWND", vbNullString)
            For i = 1 To 2
                If i = 1 Then CNS = FindWindowExA(Explorer_Shell_DUI_UIHWND_SV, 0, "CtrlNotifySink", vbNullString)
                CNS = FindWindowExA(Explorer_Shell_DUI_UIHWND_SV, CNS, "CtrlNotifySink", vbNullString)
            Next
            .Sleeping 150
            shelldll = FindWindowExA(CNS, 0, "SHELLDLL_DefView", vbNullString)
            .Sleeping 150
            SetForegroundWindow shelldll
            .Sleeping 150
            .HitKeyReturn
            .CloseObjWindow Explorer_SV
            BlockInput False
        End If
        .Sleeping 2000
        hWnd = .getObjectHwnd("TfrmMain", vbNullString)
        hwndlogin = .getObjectHwnd("TfrmLogin", vbNullString)
    End With
    If hWnd = 0 And hwndlogin = 0 Then
        DASHBOARD.Range("J19").Font.Color = 255: DASHBOARD.Range("J19").Value = "SingleView is not detected. Please manually Open it up"
    ElseIf Not hWnd = 0 Then
        Dim sv As Billing.Application: Set sv = New Billing.Application: tNo = sv.LoginName
        DASHBOARD.Range("J19").Font.Color = -11489280: DASHBOARD.Range("J19").Value = "Macro`s ready."
    End If
    'If Not hwndlogin = 0 Then DASHBOARD.Range("J19").Font.Color = 255: DASHBOARD.Range("H16").Value = "You have SV Login window open; Please type your ID&Password and click OK to login"
    Set winapi = Nothing
End Sub

Function AutoOpenLoginSV() As Boolean
    Dim hWnd As Long, hwndlogin As Long, Explorer_SV As Long
    Dim sv As New clsSingleViewMain
    Dim winapi As New clsWinAPI
    Dim Explorer_Shell_SV As Long, Explorer_Shell_DUI_SV As Long, Explorer_Shell_DUI_UIHWND_SV As Long, _
        shelldll As Long, CNS As Long, tNo As String
    Dim TXPanelLogin As Long, TXPanelLogin_sub As Long, TXGroupBox As Long, _
        TXComboBoxPlus As Long, TXEdit_Pass As Long, TXEdit_User As Long, TXBitBtn_OK As Long
    Dim i As Integer
    DASHBOARD.Activate
    If DASHBOARD.Range("Q17").Value = "" Or DASHBOARD.Range("Q18").Value = "" Then
        AutoOpenLoginSV = False
        Exit Function
    End If
    Call OpenSV
    With winapi
        hWnd = 0
        hWnd = .getObjectHwnd("TfrmMain", vbNullString)
        If hWnd = 0 Then
            AutoOpenLoginSV = False
            Exit Function
        End If
        hwndlogin = .getObjectHwnd("TfrmLogin", vbNullString)
        If hwndlogin <> 0 Then
            .Sleeping 500
            TXPanelLogin = .getChildObjectHwnd(hwndlogin, 0, "TXPanel", vbNullString)
            TXPanelLogin_sub = .getChildObjectHwnd(TXPanelLogin, 0, "TXPanel", vbNullString)
            TXGroupBox = .getChildObjectHwnd(TXPanelLogin_sub, 0, "TXGroupBox", vbNullString)
            TXComboBoxPlus = .getChildObjectHwnd(TXGroupBox, 0, "TXComboBoxPlus", vbNullString)
            TXEdit_Pass = .getChildObjectHwnd(TXGroupBox, 0, "TXEdit", vbNullString)
            TXEdit_User = .getChildObjectHwnd(TXGroupBox, TXEdit_Pass, "TXEdit", vbNullString)
            TXBitBtn_OK = .getChildObjectHwnd(TXPanelLogin_sub, 0, "TXBitBtn", "OK")
            .SendText TXComboBoxPlus, "PROD"
            .SendText TXEdit_User, DASHBOARD.Range("Q17").Value
            .SendText TXEdit_Pass, DASHBOARD.Range("Q18").Value
            .PostmsgClick TXBitBtn_OK
            .Sleeping 1500
            If sv.MsgFormChk <> 0 Then
                sv.MsgFormClick
                .Sleeping 500
                .CloseObjWindow hwndlogin
                AutoOpenLoginSV = False
                Set sv = Nothing
                Set winapi = Nothing
                Exit Function
            End If
        End If
        Do While hwndlogin <> 0
            hwndlogin = .getObjectHwnd("TfrmLogin", vbNullString)
            .Sleeping 1000
        Loop
    End With
    Set winapi = Nothing
    Set sv = Nothing
    BlockInput False
    AutoOpenLoginSV = True
End Function
