VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBMsgSenderNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public number As String
Public state As Integer
Public representYN As Boolean

Public Sub fromDictionary(dic As Variant)
    number = dic.Item("number")
    state = dic.Item("state")
    representYN = dic.Item("representYN")
   
End Sub
