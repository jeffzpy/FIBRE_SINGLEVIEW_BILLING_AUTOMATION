VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEnum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Enum OrderType
    aConnectPrimary = 100
    aConnectSecondary = 100
    aConnectAndReplace = 100
    aTransferPrimary = 200
    aTransferSecondary = 200
    aDisconnectPrimary = 300
    aDisconnectSecondary = 300
    aChangeOffer = 400
    aModifyAttribute = 500
End Enum

Public Property Get OT_Connect() As OrderType
    OT_Connect = aConnectPrimary
    OT_Connect = aConnectSecondary
End Property
