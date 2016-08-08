VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public sender As String
Public senderName As String
Public receiver As String
Public receiverName As String
Public content As String
Public subject As String


Public Function toDictionary() As Variant
    Set toDictionary = CreateObject("Scripting.Dictionary")
    
    toDictionary.Add "rcv", receiver
    If sender <> "" Then toDictionary.Add "snd", sender
    If senderName <> "" Then toDictionary.Add "sndnm", senderName
    If receiverName <> "" Then toDictionary.Add "rcvnm", receiverName
    If content <> "" Then toDictionary.Add "msg", content
    If subject <> "" Then toDictionary.Add "sjt", subject
    
End Function
