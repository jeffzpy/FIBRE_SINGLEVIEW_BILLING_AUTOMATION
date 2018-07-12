VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CaseDisconnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private aPortalID As Long
Private aOrderType As String
Private aASID As Long
Private aStartDate As String
Private aCFS As String
Private aBillingtatus As String
Private aB1 As String
Private aB2 As String
Private aB3 As String
Private aB4 As String
Private aV1 As String
Private aV2 As String
Private aErrCode As String

Public Property Let PortalID(bPortalID As Long)
    aPortalID = bPortalID
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

Public Property Get BillingStatus() As String
    BillingStatus = aBillingtatus
End Property
'aErrCode
Public Property Get ErrCode() As String
    ErrCode = aErrCode
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

Public Sub Process()
    Dim sv As New clsSingleViewMain
    Dim winapi As New clsWinAPI
    Dim srch As New objSearchWDs
    Dim svHwnd As Long, mdi As Long
    Dim pi As New objProductInstanceWD
    Dim sd As New objServiceDetailWD
    Dim i As Integer, PCSName As String
    
    With sv 'Within "clsSingleViewMain" Object code, call below methods, or properties
        svHwnd = .svHwnd
        If svHwnd = 0 Then Exit Sub
        mdi = .mdi
        If mdi = 0 Then Exit Sub
    End With 'End clsSingleViewMain
    
    With srch 'Within "objSearchWDs" Object code, call below methods, or properties
        .Asid = aASID
        .SearchCircuitByAsid
        If .ErrCode <> "" Then
            aErrCode = .ErrCode & " Circuit ID:" & aCFS
            Set srch = Nothing
            Set sv = Nothing
            Set pi = Nothing
            Set sd = Nothing
            Exit Sub
        End If
        If aCFS = "" Then
            aCFS = srch.circuitID
        End If
    End With 'End objSearchWDs
    
    With sd
        .Init = sv.mdi
        aBillingtatus = .GetServiceStatus
        If aCFS = "" Then
            aCFS = .GetPSNValue
        End If
    End With
    If aBillingtatus = "Active" Then
        winapi.Sleeping 500
        sd.ClickBtnProdInstance

        With pi 'Within "objProductInstanceWD" Object code, call below methods, or properties
            .Init = sv.mdi 'PI need to be initialized again duo changing Child Tab Object
        End With 'End of "objProductInstanceWD"
    
        '.HandlingConfirmWin True 'Click Yes button when Confirm dialog window pop out

        With sv 'Within "clsSingleViewMain" Object code, call below methods, or properties
            .ClickBtnDelete
            Do While .WaitTConfirmForm = 0
                winapi.Sleeping 100
            Loop
            Call .ProcessTimeConfirmForm(True, aStartDate)
            winapi.Sleeping 500
            Do While .MsgFormChk = 0
                winapi.Sleeping 100
            Loop
            Call .MsgFormClick
        End With 'End clsSingleViewMain
    Else
        aErrCode = "Service already Deleted"
    End If
    Call sv.CloseAllChildWindows
    aBillingtatus = "Cancelled"
    Call Process_Tes_Bill
    Set sv = Nothing
    Set pi = Nothing
    Set sd = Nothing
    Set srch = Nothing
End Sub

Private Sub Process_Tes_Bill()
    Dim sv As New clsSingleViewMain
    Dim winapi As New clsWinAPI
    Dim srch As objSearchWDs
    Dim svHwnd As Long, mdi As Long
    Dim pi As objProductInstanceWD
    Dim sd As objServiceDetailWD
    Dim i As Integer, PCSName As String
    Dim tempValue As String, tempLng As Long
    Dim aCollect As Collection: Set aCollect = New Collection
    Dim ErrOccured As Boolean
    aCFS = Trim(aCFS)
    If Left(aCFS, 2) = "00" Then
        tempLng = CLng(aCFS): tempValue = tempLng
    End If
    If Left(aCFS, 1) = "0" And Mid(aCFS, 2, 1) <> "0" Then
        tempValue = Trim(Right(aCFS, Len(aCFS) - 2))
    End If
    If aB1 <> "" Then
        aB1 = "B1" & tempValue
        aCollect.Add (aB1)
    End If
    If aB2 <> "" Then
        aB2 = "B2" & tempValue
        aCollect.Add (aB2)
    End If
    If aB3 <> "" Then
        aB3 = "B3" & tempValue
        aCollect.Add (aB3)
    End If
    If aB4 <> "" Then
        aB4 = "B4" & tempValue
        aCollect.Add (aB4)
    End If
    If aV1 <> "" Then
        aV1 = "V1" & tempValue
        aCollect.Add (aV1)
    End If
    If aV2 <> "" Then
        aV2 = "V2" & tempValue
        aCollect.Add (aV2)
    End If
    For i = 1 To aCollect.Count
        ErrOccured = False
        Set srch = New objSearchWDs
        Set pi = New objProductInstanceWD
        Set sd = New objServiceDetailWD
        With sv 'Within "clsSingleViewMain" Object code, call below methods, or properties
            svHwnd = .svHwnd
            If svHwnd = 0 Then Exit Sub
            mdi = .mdi
            If mdi = 0 Then Exit Sub
        End With 'End clsSingleViewMain
        With srch
            .circuitID = aCollect.item(i)
            .SearchCircuitByCircuitID
            If .ErrCode <> "" Then
                aErrCode = .ErrCode & " Circuit ID:" & aCollect.item(i)
                Set winapi = Nothing
                Set srch = Nothing
                Set sv = Nothing
                Set pi = Nothing
                Set sd = Nothing
                ErrOccured = True
            End If
        End With
        With sd
            If ErrOccured = False Then
                .Init = mdi
                aBillingtatus = .GetServiceStatus
                'aCFS = .GetPSNValue
            End If
        End With
        
        If ErrOccured = False Then
            If aBillingtatus = "Active" Then
                winapi.Sleeping 500
                sd.ClickBtnProdInstance
                With pi 'Within "objProductInstanceWD" Object code, call below methods, or properties
                    .Init = mdi 'PI need to be initialized again duo changing Child Tab Object
                End With 'End of "objProductInstanceWD"
            
            '.HandlingConfirmWin True 'Click Yes button when Confirm dialog window pop out
        
                With sv 'Within "clsSingleViewMain" Object code, call below methods, or properties
                    .ClickBtnDelete
                    Do While .WaitTConfirmForm = 0
                        winapi.Sleeping 100
                    Loop
                    Call .ProcessTimeConfirmForm(True, aStartDate)
                    winapi.Sleeping 500
                    Do While .MsgFormChk = 0
                        winapi.Sleeping 100
                    Loop
                    Call .MsgFormClick
                End With 'End clsSingleViewMain
            Else
                aErrCode = "INACTIVE BILLING STATUS BEFORE CLICK DELETE BUTTON"
            End If
        End If
        Call sv.CloseAllChildWindows
        Set srch = Nothing
        Set pi = Nothing
        Set sd = Nothing
    Next i
    Set sv = Nothing
    Set winapi = Nothing
End Sub



