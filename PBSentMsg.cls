VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBSentMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public state           As Integer
Public result          As Integer
Public subject         As String
Public messageType     As String
Public content         As String
Public sendNum         As String
Public senderName      As String
Public receiveNum      As String
Public receiveName     As String
Public receiptDT       As String
Public reserveDT       As String
Public sendDT          As String
Public resultDT        As String
Public sendResult      As String
Public tranNet         As String
Public receiptNum      As String
Public requestNum      As String
Public interOPRefKey   As String

Public Sub fromDictionary(dic As Variant)

    state = dic.Item("state")
    result = dic.Item("result")
    subject = dic.Item("subject")
    messageType = dic.Item("type")
    content = dic.Item("content")
    sendNum = dic.Item("sendNum")
    senderName = dic.Item("senderName")
    receiveNum = dic.Item("receiveNum")
    receiveName = dic.Item("receiveName")
    receiptDT = dic.Item("receiptDT")
    reserveDT = dic.Item("reserveDT")
    sendDT = dic.Item("sendDT")
    resultDT = dic.Item("resultDT")
    sendResult = dic.Item("sendResult")
    tranNet = dic.Item("tranNet")
    receiptNum = dic.Item("receiptNum")
    requestNum = dic.Item("requestNum")
    interOPRefKey = dic.Item("interOPRefKey")
    
End Sub

