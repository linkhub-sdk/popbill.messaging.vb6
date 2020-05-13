VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBMessageBriefInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public rNum As String
Public sn As String
Public stat As String
Public rlt As String
Public sDT As String
Public rDT As String
Public net As String
Public srt As String

Public Sub fromDictionary(dic As Variant)
    rNum = dic.Item("rNum")
    sn = dic.Item("sn")
    stat = dic.Item("stat")
    rlt = dic.Item("rlt")
    sDT = dic.Item("sDT")
    rDT = dic.Item("rDT")
    net = dic.Item("net")
    srt = dic.Item("srt")
End Sub
