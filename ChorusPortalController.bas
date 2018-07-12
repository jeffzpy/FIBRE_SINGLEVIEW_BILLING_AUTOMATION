Attribute VB_Name = "ChorusPortalController"
Option Compare Text

Sub itinalIEObj(ByVal ie As InternetExplorer)
    ie.Toolbar = 0
    ie.Visible = True
End Sub

Sub BrowseCP(ByVal ie As InternetExplorer, address As String)
'Browse Navigator------------------------------------^
    Dim cptitle As String
    On Error Resume Next
    If address = "Search" Then
        If ie.Document.URL = "https://portal.chorus.co.nz/chorus-ssp-web/pages/OrderManagement/Search" Then
            Exit Sub
        End If
    End If
RetryNav:
    For i = 1 To 10
        cptitle = Trim(ie.Document.title)
        If cptitle <> "Chorus Self Service Portal" Then
            Sleep 500
        Else
            Exit For
        End If
    Next
    If cptitle = "Sign In" Then
        Login ie, DATASHEET1.Range("Q13").Value, DATASHEET1.Range("Q14").Value
        GoTo RetryNav
    End If

    Select Case address
        Case "Create"
            address = "https://portal.chorus.co.nz/chorus-ssp-web/pages/OrderManagement/InitiateOrder"
        Case "Search"
            address = "https://portal.chorus.co.nz/chorus-ssp-web/pages/OrderManagement/Search"
            Err.Clear
            If ie.Document.getElementById("button-quick-search").innerText = "SEARCH" Then
                If Err.Number = 0 Then Exit Sub
            End If
        Case "Work"
            address = "https://portal.chorus.co.nz/chorus-ssp-web/pages/WorkQueue/WorkQueue"
    End Select
    ie.Navigate address
    CPWait ie
End Sub

Sub BrowseCP2(ByVal ie As InternetExplorerMedium, address As String)
'Browse Navigator------------------------------------^
    Dim cptitle As String
    On Error Resume Next
    If address = "Search" Then
        If ie.Document.URL = "https://portal.chorus.co.nz/chorus-ssp-web/pages/OrderManagement/Search" Then
            Exit Sub
        End If
    End If
RetryNav:
    For i = 1 To 10
        cptitle = Trim(ie.Document.title)
        If cptitle <> "Chorus Self Service Portal" Then
            Sleep 500
        Else
            Exit For
        End If
    Next
    If cptitle = "Sign In" Then
        Login ie, DATASHEET1.Range("Q13").Value, DATASHEET1.Range("Q14").Value
        GoTo RetryNav
    End If

    Select Case address
        Case "Create"
            address = "https://portal.chorus.co.nz/chorus-ssp-web/pages/OrderManagement/InitiateOrder"
        Case "Search"
            address = "https://portal.chorus.co.nz/chorus-ssp-web/pages/OrderManagement/Search"
            Err.Clear
            If ie.Document.getElementById("button-quick-search").innerText = "SEARCH" Then
                If Err.Number = 0 Then Exit Sub
            End If
        Case "Work"
            address = "https://portal.chorus.co.nz/chorus-ssp-web/pages/WorkQueue/WorkQueue"
    End Select
    ie.Navigate address
    CPWait ie
End Sub

Function VerifyCPLogin(ByVal ie As InternetExplorer, User As String, Pass As String) As Boolean
    Dim CPUser, CPPass, btnLogin As Object, CPURL As String, cptitle As String
    Dim i As Integer
    With ie
retrylogin:
        .Navigate "https://portal.chorus.co.nz/"
        CPWait ie
        If .LocationURL Like "*SessionTimeout*" Then GoTo retrylogin
        For i = 1 To 200
            On Error Resume Next
            cptitle = Trim(.Document.title)
            If cptitle = "Sign In" Then
                Set CPUser = .Document.getElementById("ContentPlaceHolder1_UsernameTextBox")
                Set CPPass = .Document.getElementById("ContentPlaceHolder1_PasswordTextBox")
                Set btnLogin = .Document.getElementById("ContentPlaceHolder1_btnSubmitButton")
                CPUser.Value = User
                CPPass.Value = Pass
                btnLogin.Click
                CPWait ie
                If .Document.getElementById("ContentPlaceHolder1_ErrorTextLabel").innerText Like "*Please enter a valid user name*" Then
                    VerifyCPLogin = False
                    Exit For
                Else
                    VerifyCPLogin = True
                    Exit For
                End If
                Exit For
            ElseIf cptitle = "Chorus Self Service Portal" Then
                    CPWait ie
                    VerifyCPLogin = True
                    Exit For
            Else
                .Navigate "https://portal.chorus.co.nz/"
                Sleep 1000
                CPWait ie
                VerifyCPLogin = True
            End If
        Next
    End With
End Function

Sub Login(ByVal ie As InternetExplorer, User As String, Pass As String)
    Dim CPUser, CPPass, btnLogin As Object, CPURL As String, cptitle As String
    Dim i As Integer
    With ie
retrylogin:
        .Navigate "https://portal.chorus.co.nz/"
        CPWait ie
        If InStr(.Document.title, "Maintenance") > 0 Then
            Exit Sub
        End If
        If .LocationURL Like "*SessionTimeout*" Then GoTo retrylogin
        For i = 1 To 200
            On Error Resume Next
            cptitle = Trim(.Document.title)
            If cptitle = "Sign In" Then
                Set CPUser = .Document.getElementById("ContentPlaceHolder1_UsernameTextBox")
                Set CPPass = .Document.getElementById("ContentPlaceHolder1_PasswordTextBox")
                Set btnLogin = .Document.getElementById("ContentPlaceHolder1_btnSubmitButton")
                CPUser.Value = User
                CPPass.Value = Pass
                btnLogin.Click
                CPWait ie
                Exit For
            ElseIf cptitle = "Chorus Self Service Portal" Then
                    CPWait ie
                    Exit For
            Else
                .Navigate "https://portal.chorus.co.nz/"
                Sleep 1000
                CPWait ie
            End If
        Next
    End With
End Sub

Sub test()
    Dim ie As New InternetExplorer
    ie.Visible = True
    FuncLogin ie, "818507", "Sbdldgs98bo"
    Stop
End Sub

Function FuncLogin(ByVal ie As InternetExplorer, User As String, Pass As String) As Boolean
    Dim CPUser, CPPass, btnLogin As Object, CPURL As String, cptitle As String
    Dim i As Integer, i2 As Integer, retLogin As Boolean
    Dim WinApi32 As New clsWinAPI
    With ie
        i = 0
        i2 = 0
retrylogin:
        If i = 10 Then
            FuncLogin = False
            Exit Function
        End If
        .Navigate "https://portal.chorus.co.nz/"
        WinApi32.Sleeping 3500
        Do While ie.Busy
            If i2 = 19 Then Exit Do
            DoEvents
            WinApi32.Sleeping 500
            i2 = i2 + 1
        Loop
        i2 = 0
        Do Until ie.ReadyState = 4
            If i2 = 19 Then Exit Do
            DoEvents
            i2 = i2 + 1
        Loop
        If InStr(.Document.title, "Maintenance") > 0 Then
            FuncLogin = False
            Exit Function
        End If
        If InStr(.LocationURL, ".chorus.co.nz") < 0 Then
            retLogin = True
            WinApi32.Sleeping 500
            i = i + 1
            GoTo retrylogin
        End If
        For i = 1 To 200
            On Error Resume Next
            cptitle = Trim(.Document.title)
            If cptitle = "Sign In" Then
                Set CPUser = .Document.getElementById("ContentPlaceHolder1_UsernameTextBox")
                Set CPPass = .Document.getElementById("ContentPlaceHolder1_PasswordTextBox")
                Set btnLogin = .Document.getElementById("ContentPlaceHolder1_btnSubmitButton")
                CPUser.Value = User
                CPPass.Value = Pass
                btnLogin.Click
                CPWait ie
                Exit For
            ElseIf cptitle = "Chorus Self Service Portal" Then
                CPWait ie
                Exit For
            Else
                .Navigate "https://portal.chorus.co.nz/"
                Sleep 1000
                CPWait ie
            End If
        Next
        Sleep 1000
        If InStr(.Document.title, "Sign In") > 0 Then FuncLogin = False
        If InStr(.Document.title, "Chorus Self Service Portal") > 0 Then FuncLogin = True
    End With
    Sleep 1000
    Set WinApi32 = Nothing
End Function
    
    Sub CPLogout(ByVal ie As InternetExplorer)
        Dim link As HTMLAnchorElement, i As Integer
            For Each link In ie.Document.getElementsByTagName("a")
                If link.innerText = "Logout" Then
                    link.Click
                    CPWait ie
                    Exit For
                End If
            Next
    End Sub
    
    Sub Tab_CreateOrder(ByVal ie As InternetExplorer, var As Long)
        Dim TLCID, ProID, TLCSearch, ProSearech As Object
        Dim link As HTMLAnchorElement
        With ie
            Set TLCID = .Document.getElementById("locationSearchId")
            Set ProID = .Document.getElementById("productSearchId")
            Set TLCSearch = .Document.getElementById("locationSearchButton")
            Set Prosearch = .Document.getElementById("productSearchButton")
            
            If var Like "*1636*" Or var Like "*1621*" And Len(var) = 10 Then
            
            
                ProID.Value = var
                Prosearch.Click
                CPWait ie
            Else
                TLCID.Value = var
                TLCSearch.Click
                CPWait ie
            End If
            
            Set AddressSearchMsg = .Document.getElementsByTagName("Span")(0)
            If AddressSearchMsg.innerText Like "*could not be found*" Then
            Exit Sub
            End If
            Dim Z As Object
            Dim x As Object
            Set evt = .Document.createEvent("HTMLEvents")
            evt.initEvent "change", True, False
            Set Z = .Document.getElementById("customer")
            Z.selectedIndex = 7 'Select RSP; Pick 'CallPlus' for now, Select different index# will get script to select different RSP.
            Z.dispatchEvent evt 'trigger event happen
    
            
            Set x = .Document.getElementById("selectedProductInstanceAtLocationIndex1")
            x.Checked = True
    End With
End Sub

Sub Tab_SearchOrders(ByVal ie As InternetExplorer, SearchValue As Long)
'This Sub Procedure can only search for Portal ID & Product ID
'It does not support Advanced Search functionalities
    Dim findvalue As Object, btnSearch As Object
    Dim link As HTMLAnchorElement
    With ie
        If Left(SearchValue, 4) = 1636 Or Left(SearchValue, 4) = 1621 Then
            .Document.getElementsByTagName("option")(3).Selected = True ' Select Product ID
        Else
            .Document.getElementsByTagName("option")(0).Selected = True ' Select Order ID
        End If
                Set findvalue = .Document.getElementById("searchValue")
                CPWait ie
                Set findvalue = .Document.getElementById("searchValue")
                findvalue.Value = SearchValue
                Set btnSearch = .Document.getElementById("button-quick-search")
                btnSearch.Click
                CPWait ie
                'Click searched result
                For Each link In .Document.getElementsByTagName("a")
                    If link.innerText = CStr(SearchValue) Then
                        link.Click
                        Exit For
                    End If
                Next
                CPWait ie
    End With
End Sub

Sub WorkQueuePageManager(ByVal ie As InternetExplorer, QueueType As String, TaskType As String, Filter_ProductOffer As String, Filter_Classification As String, Filter_ServiceProvider As String, _
    Filter_ProductFamaily As String, Filter_OrderType As String, Filter_OrderStatus As String, Filter_Substatus As String, Filter_TaskType As String, Filter_Assignee As String)
    Dim link As HTMLAnchorElement
    
    With ie
        For Each link In .Document.getElementsByTagName("label")
            If link.innerText Like TaskType Then
                link.Click
                CPWait ie
                Sleep 500
                Exit For
            End If
        Next
        For Each link In .Document.getElementsByTagName("option")
            If link.innerText Like QueueType Then
                link.Selected = True
                link.Click
                CPWait ie
                Exit For
            End If
        Next
        Set evt = .Document.createEvent("HTMLEvents")
        evt.initEvent "change", True, False
        Set Z = .Document.getElementById("workQueueList")
        On Error Resume Next
        CPWait ie
        Z.dispatchEvent evt
        CPWait ie
        WQT_FilterSelect ie, Filter_ProductOffer
        WQT_FilterSelect ie, Filter_Classification
        WQT_FilterSelect ie, Filter_ServiceProvider
        WQT_FilterSelect ie, Filter_ProductFamaily
        WQT_FilterSelect ie, Filter_OrderType
        WQT_FilterSelect ie, Filter_OrderStatus
        WQT_FilterSelect ie, Filter_Substatus
        WQT_FilterSelect ie, Filter_TaskType
        WQT_FilterSelect ie, Filter_Assignee

        For Each link In .Document.getElementsByTagName("button")
            If link.innerText = "Apply Filters" Then
                link.Click
                CPWait ie
                Exit For
            End If
        Next
    End With
End Sub

'Function TES_Charges_CP(ByVal IE As InternetExplorer, Description As String, rownumber As Integer) As Collection
'    Dim link As HTMLAnchorElement, i As Integer, tempVar() As String
'    With IE
'        CPWait IE
'        i = 1
'        For Each link In .Document.getElementsByTagName("td")
'            tempVar(i) = CStr(link.innerText)
'            If InStr(tempVar(i), "Tail Extn V1") > 0 Then
'                DATASHEET.Range("AA" & rownumber).value = DATASHEET.Range("AA" & rownumber).value & "TESV1"
'                TES_Charges_CP.Add (tempVar(i))
'            End If
'            If InStr(tempVar(i), "Tail Extn B1") > 0 Then
'                DATASHEET.Range("AA" & rownumber).value = DATASHEET.Range("AA" & rownumber).value & "TESB1"
'                TES_Charges_CP.Add (tempVar(i))
'            End If
'            If InStr(tempVar(i), "Tail Extn B2") > 0 Then
'                DATASHEET.Range("AA" & rownumber).value = DATASHEET.Range("AA" & rownumber).value & "TESB2"
'                TES_Charges_CP.Add (tempVar(i))
'            End If
'            If InStr(tempVar(i), "Tail Extn B3") > 0 Then
'                DATASHEET.Range("AA" & rownumber).value = DATASHEET.Range("AA" & rownumber).value & "TESB3"
'                TES_Charges_CP.Add (tempVar(i))
'            End If
'            If InStr(tempVar(i), "Tail Extn B4") > 0 Then
'                DATASHEET.Range("AA" & rownumber).value = DATASHEET.Range("AA" & rownumber).value & "TESB4"
'                TES_Charges_CP.Add (tempVar(i))
'            End If
'            i_summary = i_summary + 1
'        Next
'        cse = planlist(19)
'    End With
'End Function

Sub btnSELECT_ALL_WQ(ByVal ie As InternetExplorer)
    Dim link As HTMLAnchorElement
    For Each link In ie.Document.getElementsByTagName("button")
        If link.innerText Like "*Select All*" Then
            link.Click
            CPWait ie
            Exit For
        End If
    Next
End Sub

Sub btnUNASSIGN_WQ(ByVal ie As InternetExplorer)
    Dim link As HTMLAnchorElement
    For Each link In ie.Document.getElementsByTagName("button")
        If link.innerText Like "*Un-Assign*" Then
            link.Click
            CPWait ie
            Exit For
        End If
    Next
End Sub

Sub btnASSIGNTOME_WQ(ByVal ie As InternetExplorer)
    Dim link As HTMLAnchorElement
    For Each link In ie.Document.getElementsByTagName("button")
        If link.innerText Like "*Assign to Me*" Then
            link.Click
            CPWait ie
            Exit For
        End If
    Next
End Sub

Private Sub btnASSIGNTOOTHER_WQ(ByVal ie As InternetExplorer)
    Dim link As HTMLAnchorElement
    For Each link In ie.Document.getElementsByTagName("button")
        If link.innerText Like "*Assign To Other*" Then
            link.Click
            CPWait ie
            Exit For
        End If
    Next
End Sub

Sub AssignJobstoOther(ByVal ie As InternetExplorer, AssigneeName As String)
    Dim link As HTMLAnchorElement
    btnASSIGNTOOTHER_WQ ie
    Sleep 800
    For Each link In ie.Document.getElementsByTagName("Option")
        If link.innerText Like AssigneeName Then
            link.Selected = True
            link.Click
            CPWait ie
            Exit For
        End If
    Next
    'ie.Document.getElementById("AssignJobstoOther").Click
    For Each link In ie.Document.getElementsByClassName("button continue float-right ssp-action-btn")
        If link.ID Like "*Assign*" Then
            link.Click
            CPWait ie
            Exit For
        End If
    Next
End Sub
Function getAlerMsg(ByVal ie As InternetExplorer) As String
    Dim link As HTMLAnchorElement, Value(100) As String, i As Integer
    i = 1
    For Each link In ie.Document.getElementsByClassName("alert alert-success")
        Value(i) = link.innerText
        'Debug.Print Value(i)
    Next
End Function

Private Sub WQT_FilterSelect(ByVal ie As InternetExplorer, FilterValue As String)
    Dim link As HTMLAnchorElement
    If FilterValue = "" Or FilterValue = vbNullString Or FilterValue = NullString Then Exit Sub
    CPWait ie
    For Each link In ie.Document.getElementsByTagName("option")
        If link.innerText Like FilterValue Then
        link.Selected = True
        CPWait ie
        Exit For
        End If
    Next
    CPWait ie
End Sub


Sub order_tab_click(ByVal ie_summary As InternetExplorer, ie_navigate As InternetExplorer, tab_name As String)
    Dim link As HTMLAnchorElement
    With ie_summary
    For Each link In .Document.getElementsByTagName("a")
       If InStr(link.innerText, tab_name) > 0 Then
           ie_navigate.Navigate link.getAttribute("href")
           CPWait ie_summary
           Exit Sub
       End If
    Next
    End With
End Sub

Function summary_tab_object(ByVal ie As InternetExplorer, Details As String) As Object
    Dim link As HTMLAnchorElement, p_tag As Object
    With ie
        Set p_tag = .Document.getElementsByTagName("p")
        For Each link In p_tag
            If InStr(link.innerText, Details) > 0 Then
                Set summary_tab_object = link
                Set link = Nothing
                Set p_tag = Nothing
                Exit Function
            End If
        Next
    End With
End Function

Function get_business_premium(ByVal ie As InternetExplorer, OrderType As String) As String
    If InStr(OrderType, "Modify") > 0 Then
        get_business_premium = Trim(summary_tab_object(ie, "Classification").innerText)
        If InStr(get_business_premium, "Business Premium") > 0 Then
            get_business_premium = "Business Premium"
        Else
            get_business_premium = "NA"
        End If
    End If
End Function

Sub get_tes_plans(charge_tab, order_type, rownumber)
    Dim link As HTMLAnchorElement
    On Error Resume Next
    Err.Clear
    If InStr(order_type, "Disconnect") > 0 Then aOrderType = "disc"
    For Each link In charge_tab.Document.getElementsByTagName("td")
        'If Err.Number <> 0 Then Exit Sub
        If InStr(link.innerText, "Tail Extn V1") > 0 Then
'            If aOrderType <> "disc" Then
'                TESV1_PLAN = link.innerText
'                DATASHEET.Range("G" & rownumber).Value = DATASHEET.Range("G" & rownumber).Value & vbNewLine & Trim(TESV1_PLAN)
'                DATASHEET.Range("H" & rownumber).Value = DATASHEET.Range("H" & rownumber).Value & vbNewLine & _
'                                                Application.WorksheetFunction.VLookup(Trim(TESV1_PLAN), Sheet2.Range("S:U"), 2, False)
'                DATASHEET.Range("I" & rownumber).Value = DATASHEET.Range("I" & rownumber).Value & vbNewLine & _
'                                                Application.WorksheetFunction.VLookup(Trim(TESV1_PLAN), Sheet2.Range("S:U"), 3, False)
'            End If
            If DATASHEET.Range("AB" & rownumber).Value = "" Then
                DATASHEET.Range("AB" & rownumber).Value = "V1"
            Else
                DATASHEET.Range("AB" & rownumber).Value = DATASHEET.Range("AB" & rownumber).Value & "/" & "V1"
            End If
        End If
        If InStr(link.innerText, "Tail Extn V2") > 0 Then
'            If aOrderType <> "disc" Then
'                TESV2_PLAN = link.innerText
'                DATASHEET.Range("G" & rownumber).Value = DATASHEET.Range("G" & rownumber).Value & vbNewLine & Trim(TESV2_PLAN)
'                DATASHEET.Range("H" & rownumber).Value = DATASHEET.Range("H" & rownumber).Value & vbNewLine & _
'                                                    Application.WorksheetFunction.VLookup(Trim(TESV2_PLAN), Sheet2.Range("S:U"), 2, False)
'                DATASHEET.Range("I" & rownumber).Value = DATASHEET.Range("I" & rownumber).Value & vbNewLine & _
'                                                    Application.WorksheetFunction.VLookup(Trim(TESV2_PLAN), Sheet2.Range("S:U"), 3, False)
'            End If
            If DATASHEET.Range("AB" & rownumber).Value = "" Then
                DATASHEET.Range("AB" & rownumber).Value = "V2"
            Else
                DATASHEET.Range("AB" & rownumber).Value = DATASHEET.Range("AB" & rownumber).Value & "/" & "V2"
            End If
        End If
        If InStr(link.innerText, "Tail Extn B1") > 0 Then
'            If aOrderType <> "disc" Then
'                TESB1_PLAN = link.innerText
'                DATASHEET.Range("G" & rownumber).Value = DATASHEET.Range("G" & rownumber).Value & vbNewLine & Trim(TESB1_PLAN)
'                DATASHEET.Range("H" & rownumber).Value = DATASHEET.Range("H" & rownumber).Value & vbNewLine & _
'                                                Application.WorksheetFunction.VLookup(Trim(TESB1_PLAN), Sheet2.Range("S:U"), 2, False)
'                DATASHEET.Range("I" & rownumber).Value = DATASHEET.Range("I" & rownumber).Value & vbNewLine & _
'                                                Application.WorksheetFunction.VLookup(Trim(TESB1_PLAN), Sheet2.Range("S:U"), 3, False)
'            End If
            If DATASHEET.Range("AB" & rownumber).Value = "" Then
                DATASHEET.Range("AB" & rownumber).Value = "B1"
            Else
                DATASHEET.Range("AB" & rownumber).Value = DATASHEET.Range("AB" & rownumber).Value & "/" & "B1"
            End If
        End If
        If InStr(link.innerText, "Tail Extn B2") > 0 Then
'            If aOrderType <> "disc" Then
'                TESB2_PLAN = link.innerText
'                DATASHEET.Range("G" & rownumber).Value = DATASHEET.Range("G" & rownumber).Value & vbNewLine & Trim(TESB2_PLAN)
'                DATASHEET.Range("H" & rownumber).Value = DATASHEET.Range("H" & rownumber).Value & vbNewLine & _
'                                                Application.WorksheetFunction.VLookup(Trim(TESB2_PLAN), Sheet2.Range("S:U"), 2, False)
'                DATASHEET.Range("I" & rownumber).Value = DATASHEET.Range("I" & rownumber).Value & vbNewLine & _
'                                                Application.WorksheetFunction.VLookup(Trim(TESB2_PLAN), Sheet2.Range("S:U"), 3, False)
'            End If
            If DATASHEET.Range("AB" & rownumber).Value = "" Then
                DATASHEET.Range("AB" & rownumber).Value = "B2"
            Else
                DATASHEET.Range("AB" & rownumber).Value = DATASHEET.Range("AB" & rownumber).Value & "/" & "B2"
            End If
        End If
        If InStr(link.innerText, "Tail Extn B3") > 0 Then
'            If aOrderType <> "disc" Then
'                TESB3_PLAN = link.innerText
'                DATASHEET.Range("G" & rownumber).Value = DATASHEET.Range("G" & rownumber).Value & vbNewLine & Trim(TESB3_PLAN)
'                DATASHEET.Range("H" & rownumber).Value = DATASHEET.Range("H" & rownumber).Value & vbNewLine & _
'                                                Application.WorksheetFunction.VLookup(Trim(TESB3_PLAN), Sheet2.Range("S:U"), 2, False)
'                DATASHEET.Range("I" & rownumber).Value = DATASHEET.Range("I" & rownumber).Value & vbNewLine & _
'                                                Application.WorksheetFunction.VLookup(Trim(TESB3_PLAN), Sheet2.Range("S:U"), 3, False)
'            End If
            If DATASHEET.Range("AB" & rownumber).Value = "" Then
                DATASHEET.Range("AB" & rownumber).Value = "B3"
            Else
                DATASHEET.Range("AB" & rownumber).Value = DATASHEET.Range("AB" & rownumber).Value & "/" & "B3"
            End If
        End If
        If InStr(link.innerText, "Tail Extn B4") > 0 Then
'            If aOrderType <> "disc" Then
'                TESB4_PLAN = link.innerText
'                DATASHEET.Range("G" & rownumber).Value = DATASHEET.Range("G" & rownumber).Value & vbNewLine & Trim(TESB4_PLAN)
'                DATASHEET.Range("H" & rownumber).Value = DATASHEET.Range("H" & rownumber).Value & vbNewLine & _
'                                                Application.WorksheetFunction.VLookup(Trim(TESB4_PLAN), Sheet2.Range("S:U"), 2, False)
'                DATASHEET.Range("I" & rownumber).Value = DATASHEET.Range("I" & rownumber).Value & vbNewLine & _
'                                                Application.WorksheetFunction.VLookup(Trim(TESB4_PLAN), Sheet2.Range("S:U"), 3, False)
'            End If
            If DATASHEET.Range("AB" & rownumber).Value = "" Then
                DATASHEET.Range("AB" & rownumber).Value = "B4"
            Else
                DATASHEET.Range("AB" & rownumber).Value = DATASHEET.Range("AB" & rownumber).Value & "/" & "B4"
            End If
        End If
    Next

End Sub

Function non_standard_install(ByVal charge_tab As InternetExplorer) As String
    Dim quote_status_obj As Object, QuoteValue As String, QuoteValues As String
    On Error Resume Next: Err.Clear
    Set quote_status_obj = charge_tab.Document.getElementById("quoteStatus")
    If quote_status_obj.Value = "STATUS_APPROVED" Then
        QuoteValue = charge_tab.Document.getElementById("quoteValue").Value
        QuoteValues = CStr(QuoteValue)
        QuoteValues = WorksheetFunction.Substitute(CStr(QuoteValues), ",", "")
        non_standard_install = QuoteValues
    End If
End Function
Function get_acct(ByVal ie_OrderCharges As InternetExplorer) As Long
    Dim billing_panel As Object
    Set billing_panel = ie_OrderCharges.Document.getElementsByClassName("row billing-panel")
    For Each link In billing_panel
        If InStr(link.innerText, "Billing Account:") > 0 Then
            acct_str = link.innerText
            acct_str = Trim(Replace(acct_str, "Billing Account:", ""))
            acct_str = Trim(Left(acct_str, InStr(acct_str, "Billing Start") - 1))
            get_acct = acct_str
            Set billing_panel = Nothing
            Set billing_panel = Nothing
            Set link = Nothing
            acct_str = Null
            Exit For
        End If
    Next
End Function

Property Get get_oneoff_charge(ByVal ie_ProductCha As InternetExplorer, order_type As String, voda_plan As String, RorB, charge_array As Collection) As Collection
    Dim productchar_div As HTMLDivElement, productchar_section As Object, link As HTMLAnchorElement
    Dim cpe_type As String, wyah_value As String, jp_value As Integer, job_instructions As String, cpe_value As String, _
        ups_value As String, enclosure_value As String, CPETruckRoll As String
        
    If order_type = "Change Offer" Or order_type = "Modify Attribute" Then
        Set productchar_div = ie_ProductCha.Document.getElementById("productInstanceCharacteristicsDiv")
        CPETruckRoll = Trim(productchar_div.getElementsByTagName("h2")(0).innerText)
        If CPETruckRoll = "Internal Relocation - CPE Only" Then charge_array.Add "Relocate Simple CPE"
        If CPETruckRoll = "Internal Relocation - ONT Only" Then charge_array.Add "Relocate ONT" & RorB
        
        For Each link In productchar_div.getElementsByTagName("selection")
            If InStr(link.innerText, "While You Are Here:") > 0 Then
                Set productchar_section = link
                Exit For
            End If
        Next
        If productchar_section Is Not Empty Then
            Set wyah_obj = ie_ProductCha.Document.getElementById("orderItem1.productOrder0.characteristic0.value")
            Set cpe_obj = ie_ProductCha.Document.getElementById("orderItem1.productOrder0.characteristic8.value")
            Set jp_obj = ie_ProductCha.Document.getElementById("orderItem1.productOrder0.characteristic1.value")
            Set job_instructions_obj = ie_ProductCha.Document.getElementById("orderItem1.productOrder0.characteristic2.value")
            Set ups_obj = ie_ProductCha.Document.getElementById("orderItem1.productOrder0.characteristic11.value")
            Set enclosure_obj = ie_ProductCha.Document.getElementById("orderItem1.productOrder0.characteristic12.value")
                            
            job_instructions = job_instructions_obj.Value
            jp_value = jp_obj.item(jp_obj.selectedIndex).text
            cpe_value = cpe_obj.item(cpe_obj.selectedIndex).Value
            ups_value = ups_obj.item(ups_obj.selectedIndex).Value
            enclosure_value = enclosure_obj.item(enclosure_obj.selectedIndex).Value
            
            If jp_value > 0 Then
                Do While jp_value
                    charge_array.Add "Install Jack Point"
                Loop
            End If
            If cpe_value = "Simple CPE" Then
                charge_array.Add "Install Simple CPE Standard"
            ElseIf cpe_value = "Complex CPE" Then
                charge_array.Add "Install Complex CPE Standard"
            End If
            If ups_value = "Yes" Then charge_array.Add "Install UPS Standard"
            If enclosure_value = "Yes" Then charge_array.Add "Install Enclosure"
        End If
    End If
    If Left(order_type, 7) = "Connect" Or Left(order_type, 8) = "Transfer" Then
        If voda_plan = "Double Play" Then charge_array.Add "Vodafone Double Play"
        If voda_plan = "Triple Play" Then charge_array.Add "Vodafone Triple Play"
        If voda_plan = "Triple Play X2" Then charge_array.Add "Vodafone Triple Play X2"
    End If
    If Left(order_type, 8) = "Transfer" Then charge_array.Add "NGA transfer"
    Set get_oneoff_charge = charge_array
End Property


Function get_wyah(ByVal ie_ProductCha As InternetExplorer) As Boolean
    Dim productchar_div As Object, link As HTMLAnchorElement, wyah_value As String
    Set productchar_div = ie_ProductCha.Document.getElementById("productInstanceCharacteristicsDiv")
    Set wyah_obj = ie_ProductCha.Document.getElementById("orderItem1.productOrder0.characteristic0.value")
    Set aOptions = wyah_obj.getElementsByTagName("option")
    For Each link In aOptions
      If link.hasAttribute("selected") = True Then
          wyah_value = Trim(link.innerText)
          Exit For
      End If
    Next
    If wyah_value = "Yes" Then
        get_wyah = True
    ElseIf wyah_value = "No" Then
        get_wyah = False
    End If
End Function

Function get_hcir_value(ByVal ie_ProductCha As InternetExplorer, product_name As String) As String
    Dim hcir_value As String, product_char As HTMLDivElement, link As HTMLAnchorElement, temp_obj As Object
    Dim down_hp As String, up_hp As String, bandwidth_profile As String
    On Error Resume Next
    If InStr(product_name, "Chorus Better Broadband") < 1 Then
        hcir_value = NGA_PARAMETERS.Range("A:A").Find(product_name, MatchCase:=False).OffSet(0, 3).Value
    Else
        hcir_value = "Chorus Better Boradband"
    End If
    If hcir_value = "As per Portal order" Then
        Set product_char = ie_ProductCha.Document.getElementById("productInstanceCharacteristicsDiv")
        For Each link In product_char
            If InStr(link.innerText, "Downstream HP:") > 0 Then
                Set temp_obj = link.NextSibling
                down_hp = Trim(temp_obj.item(temp_obj.selectedIndex).text)
            End If
            If InStr(link.innerText, "Upstream HP:") > 0 Then
                Set temp_obj = link.NextSibling
                up_hp = Trim(temp_obj.item(temp_obj.selectedIndex).text)
            End If
            If InStr(link.innerText, "Bandwidth Profile:") > 0 Then
                Set temp_obj = link.NextSibling
                bandwidth_profile = Trim(temp_obj.item(temp_obj.selectedIndex).text)
            End If
        Next
        If IsEmpty(bandwidth_profile) = True Then
            get_hcir_value = down_up & "MBPS/" & up_hp & "MBPS"
        Else
            get_hcir_value = NGA_PARAMETERS.Range("X:X").Find(bandwidth_profile, MatchCase:=False).OffSet(0, 1).Value
        End If
    Else
        get_hcir_value = hcir_value
    End If
    If get_hcir_value = "" Then
        If InStr(product_name, "SFP") > 0 Or InStr(product_char.innerText, "Bandwidth Profile:") > 0 Then
            SFP_HighCIR = Trim(ie_ProductCha.Document.getElementById("orderItem0.productOrder0.characteristic0.value").Value)
            If Err.Number <> 0 Then
                SFP_HighCIR = Trim(ie_ProductCha.Document.getElementById("orderItem1.productOrder0.characteristic0.value.value").Value)
                Err.Clear
            End If
            sfp_hcir = NGA_PARAMETERS.Range("X:X").Find(SFP_HighCIR, MatchCase:=False).row
            sfp_hcir = NGA_PARAMETERS.Range("Y" & sfp_hcir).Value
            If InStr(sfp_hcir, "MBPS") > 0 Then
                get_hcir_value = sfp_hcir
            End If
        End If
    End If
End Function


Sub get_allocated_jobs(ByVal ie As InternetExplorer)
    If DASHBOARD.chkRESET_DATASHEET = False Then
        'Call ResetValues
        BrowseCP ie, "Work"
        WorkQueuePageManager ie, "NGA Provisioning", "My Tasks", vbNullString, vbNullString, _
                            vbNullString, vbNullString, vbNullString, _
                            vbNullString, vbNullString, "Perform Billing", vbNullString
        NextPage = True
        With ie
            Do While NextPage = True
                Set html_tag_name = .Document.getElementsByTagName("p")
                For Each link In html_tag_name
                    If (Left(Trim(link.innerText), 2) = 10 Or Left(Trim(link.innerText), 2) = 11) And Len(Trim(link.innerText)) = 9 Then
                        'portalID_CollectID(i) = link.innerText
                        NextRow = DATASHEET.Range("A" & Rows.Count).End(xlUp).row + 1
                        DATASHEET.Range("A" & NextRow).Value = Trim(link.innerText)
                        CPWait ie
                    End If
                Next link
                Set workQueueTaskListForm = .Document.getElementById("workQueueTaskListForm")
                Set html_tag_name = workQueueTaskListForm.getElementsByTagName("a")
                For Each link In html_tag_name
                    If InStr(link.innerText, "Next") > 0 Then
                        link.Click
                        CPWait ie
                        NextPage = True
                    Else
                        NextPage = False
                    End If
                Next link
            Loop
        End With
    Else
        BrowseCP ie, "Search"
        CPWait ie
    End If
End Sub

Function get_service_date(ByVal date_value As String) As String
    Dim startdate_Year As String, startdate_Month As String, startdate_Day As String, StartDate() As String
    StartDate = Split(date_value, ",")
    startdate_Year = Left(Trim(StartDate(2)), 4)
    StartDate = Split(StartDate(1), " ")
    startdate_Day = Format(onlyDigits(StartDate(1)), "00")
    startdate_Month = Format(Month(startdate_Day & "-" & StartDate(2) & "-" & startdate_Year), "00")
    get_service_date = CStr(startdate_Day & "/" & startdate_Month & "/" & startdate_Year)
    Erase StartDate
    startdate_Year = vbNullString
    startdate_Month = vbNullString
    startdate_Day = vbNullString
End Function



Function CPWait(ByVal ie As InternetExplorer)
    Do While ie.Busy
      DoEvents
    Loop
    Do Until ie.ReadyState = 4
    Loop
End Function


Function GetCSEOrderNo(ByVal ie As InternetExplorer) As Long
    Dim link As HTMLAnchorElement, CPEIndex As Integer
    On Error Resume Next: Err.Clear
    order_msg = Trim(ie.Document.getElementsByClassName("alert alert-error")(0).innerText)
    If Err.Number = 0 Then
        If order_msg = "Unable to find Work Order" Then
            GetCSEOrderNo = 0
            Exit Function
        End If
    End If
    For Each link In ie.Document.getElementById("woForm").getElementsByTagName("td")
        If link.innerText Like "*CPE*" Or link.innerText Like "*Install*" Or link.innerText Like "*CSE*" Then
            Set SiblingText = link.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling.NextSibling
            CPEOrder = SiblingText.textContent
            Exit For
        End If
    Next
    If Not link Is Nothing Then
        GetCSEOrderNo = CPEOrder
    Else
        GetCSEOrderNo = 0
    End If
End Function

Function AncillaryChargeCO(ByVal ie As InternetExplorer) As Boolean
'===========================================
'This function goes live on 24th of July.Need remove this day
'AncillaryChargeCO = False
'Exit Function
'===========================================
    With ie
        For Each link In .Document.getElementsByTagName("option")
            'Debug.Print link.innerText
            If link.innerText = "Total" Then Exit For
            If InStr(link.innerText, "NGA Change") > 0 Then
                If link.Selected = False Then
                    AncillaryChargeCO = False
                    Exit For
                Else
                    AncillaryChargeCO = True
                    Exit For
                End If
            End If
        Next
    End With
End Function

Function TransferChargeCO(ByVal ie As InternetExplorer) As Boolean
'===========================================
'This function goes live on 24th of July.Need remove this day
'TransferChargeCO = False
'Exit Function
'===========================================
    With ie
        For Each link In .Document.getElementsByTagName("option")
            'Debug.Print link.innerText
            If link.innerText = "Total" Then Exit For
            If InStr(link.innerText, "NGA Transfer") > 0 Then
                If link.Selected = False Then
                    TransferChargeCO = False
                    Exit For
                Else
                    TransferChargeCO = True
                    Exit For
                End If
            End If
        Next
    End With
End Function

Sub kill_exisiting_ie()
    Dim ie_chk As Long
    For i = 1 To 50
        ie_chk = FindWindow(vbNullString, "Internet Explorer")
        If ie_chk = 0 Then Exit For
        If ie_chk <> 0 Then SendMessage ie_chk, &H10, 0&, 0&
        Sleep 50
        ie_chk = 0
    Next
    For i = 1 To 50
        ie_chk = FindWindow(vbNullString, "https://portal.chorus.co.nz/ - Chorus Self Service Portal - Internet Explorer")
        If ie_chk = 0 Then Exit For
        If ie_chk <> 0 Then SendMessage ie_chk, &H10, 0&, 0&
        Sleep 50
        ie_chk = 0
    Next
    For i = 1 To 50
        ie_chk = FindWindow(vbNullString, "Chorus Self Service Portal - Internet Explorer")
        If ie_chk = 0 Then Exit For
        If ie_chk <> 0 Then PostMessage ie_chk, &H10, 0&, 0&
        Sleep 50
        ie_chk = 0
    Next
End Sub
