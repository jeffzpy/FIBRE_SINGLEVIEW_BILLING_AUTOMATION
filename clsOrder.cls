VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private aJobNumber As Long
Private aJobType As String
Private aJobAsid As Long
Private aJobPrimaryCFS As Long
Private aJobTesCFS_B1 As String
Private aJobTesCFS_B2 As String
Private aJobTesCFS_B3 As String
Private aJobTesCFS_B4 As String
Private aJobTesCFS_V1 As String
Private aJobTesCFS_V2 As String

Public Property Get jobnumber() As Long
    jobnumber = aJobNumber
End Property

Public Property Let jobnumber(Value As Long)
    aJobNumber = Value
End Property

Public Property Get JobType() As String
    JobType = aJobType
End Property

Public Property Let Jobype(Value As Long)
    aJobType = Value
End Property
