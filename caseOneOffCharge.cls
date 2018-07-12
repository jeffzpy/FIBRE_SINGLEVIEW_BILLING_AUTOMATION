VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "caseOneOffCharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private aPortalID As Long
Private aASID As Long
Private aStartDate As String
Private aOneOffChargeType As String
Private aReasonCode As String
Private aAmount As String

Property Let PortalID(Arg As Long)
    aPortalID = Arg
End Property

Property Let Asid(Arg As Long)
    aASID = Arg
End Property

Property Let StartDate(Arg As String)
    aStartDate = Arg
End Property

Property Let OneOffChargeType(Arg As String)
    aOneOffChargeType = Arg
End Property

Property Let ReasonCode(Arg As String)
    aReasonCode = Arg
End Property

Property Let Amount(Arg As String)
    aAmount = Arg
End Property

Sub Process(MdiHwnd As Long)
    Dim cq As New objTfrmCustomerQuery
    Dim sv As New clsSingleViewMain
    Dim winapi As New clsWinAPI
    Dim TempArr() As String, i As Integer
    If aReasonCode = "" Then Exit Sub
    If MdiHwnd = 0 Then Exit Sub
    Call winapi.ActivateMenuObject(sv.svHwnd, "New Query", 45, "sendmsg")
    winapi.Sleeping 500
    With cq
        .PortalID = aPortalID
        .ChargeDate = aStartDate
        .Amount = aAmount
        .InvoiceText = ""
        .OneOffChargeType = aOneOffChargeType
        .ReasonCode = aReasonCode
        .Init = MdiHwnd
        .SelectQryType aOneOffChargeType
        .EnableServiceTab
        .ClickBtnAssign
        .EnableGeneralTab
        .TypeinChargeCode
        winapi.Sleeping 50
        .TypeinChargeDate
        .TypeinInvoiceText
        winapi.Sleeping 50
        If aOneOffChargeType = 3 Then .TypeinChargeAmount
    End With
    With sv
        '.svHwnd
        .ClickBtnSave
        winapi.Sleeping 5000
        If .MsgFormChk <> 0 Then .MsgFormClick
    End With
    'Next
End Sub
