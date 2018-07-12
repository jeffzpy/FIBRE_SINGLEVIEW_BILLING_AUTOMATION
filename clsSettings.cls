VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Const adb_location = "G"
Const aPlanIndexLocation = ""

Public Property Get db_location() As String
    db_location = adb_location
End Property
