VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBMSGService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'=================================================================================
' Class Module for base module for Popbill API SDK. It include base functionality for
' RESTful web service request and parse json result. It uses Linkhub class module
' to accomplish authentication APIs.
'
' This module uses advapi32.dll and crypt32.dll. most of windows system has it
' on system32 by default.
'
' http://www.popbill.com
' Author : Kim Seongjun (pallet027@gmail.com)
' Written : 2014-04-22
' Contributor : Jeong Yohan (code@linkhub.co.kr)
' Updated : 2017-08-30
'
' Thanks for your interest.
'=================================================================================
Private Const ServiceID_REAL = "POPBILL"
Private Const ServiceID_TEST = "POPBILL_TEST"
Private Const ServiceURL_REAL = "https://popbill.linkhub.co.kr"
Private Const ServiceURL_TEST = "https://popbill-test.linkhub.co.kr"
Private Const APIVersion = "1.0"

Private m_IsTest As Boolean

Private m_token_Dic
Private m_Linkhub As Linkhub
Private m_scope As New Collection
Private m_LastErrCode As Long
Private m_LastErrMessage As String

Public Enum MsgType
    SMS
    LMS
    MMS
End Enum

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Sub Class_Initialize()
    Set m_token_Dic = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    m_token_Dic.RemoveAll
    Set m_token_Dic = Nothing
End Sub
Private Function UniStrToUTF8(UniString) As Byte()
   ' Convert a Unicode string to a byte stream of UTF-8
   Dim BArray() As Byte
   Dim TempB() As Byte
   Dim i As Long
   Dim k As Long
   Dim TLen As Long
   Dim b1 As Byte
   Dim b2 As Byte
   Dim UTF16 As Long
   Dim j
   TLen = Len(UniString)  ' Obtain length of Unicode input string
   If TLen = 0 Then Exit Function  ' get out if there's nothing to convert
   k = 0
   For i = 1 To TLen
      ' Work out the UTF16 value of the Unicode character
      CopyMemory b1, ByVal StrPtr(UniString) + ((i - 1) * 2), 1
      CopyMemory b2, ByVal StrPtr(UniString) + ((i - 1) * 2) + 1, 1
      ' Combine the 2 bytes into the Unicode UTF-16
      UTF16 = b2  ' assign b2 to UTF16 before multiplying by 256 to avoid overflow
      UTF16 = UTF16 * 256 + b1
      ' Convert UTF-16 to 2 or 3 bytes of UTF-8
      TempB = ToUTF8(UTF16)
      ' Copy the resultant bytes to BArray
      For j = 0 To UBound(TempB)
         ReDim Preserve BArray(k)
         BArray(k) = TempB(j): k = k + 1
      Next
      ReDim TempB(0)
   Next
   UniStrToUTF8 = BArray  ' Return the resultant UTF-8 byte array
End Function

Private Function ToUTF8(ByVal UTF16 As Long) As Byte()
   ' Convert a 16bit UTF-16BE to 2 or 3 UTF-8 bytes
   Dim BArray() As Byte
   If UTF16 < &H80 Then
      ReDim BArray(0)  ' one byte UTF-8
      BArray(0) = UTF16  ' Use number as is
   ElseIf UTF16 < &H800 Then
      ReDim BArray(1)  ' two byte UTF-8
      BArray(1) = &H80 + (UTF16 And &H3F)  ' Least Significant 6 bits
      UTF16 = UTF16 \ &H40  ' Shift UTF16 number right 6 bits
      BArray(0) = &HC0 + (UTF16 And &H1F)  ' Use 5 remaining bits
   Else
      ReDim BArray(2)  ' three byte UTF-8
      BArray(2) = &H80 + (UTF16 And &H3F)  ' Least Significant 6 bits
      UTF16 = UTF16 \ &H40  ' Shift UTF16 number right 6 bits
      BArray(1) = &H80 + (UTF16 And &H3F)  ' Use next 6 bits
      UTF16 = UTF16 \ &H40  ' Shift UTF16 number right 6 bits again
      BArray(0) = &HE0 + (UTF16 And &HF)  ' Use 4 remaining bits
   End If
   ToUTF8 = BArray  ' Return UTF-8 bytes in an array
End Function
'최근 오류코드 확인
Public Property Get LastErrCode() As Long
    LastErrCode = m_LastErrCode
End Property
'최근오류메시지 확인
Public Property Get LastErrMessage() As String
    LastErrMessage = m_LastErrMessage
End Property
'테스트 플래그
Public Property Let IsTest(ByVal value As Boolean)
    m_IsTest = value
End Property
'초기화
Public Sub Initialize(LinkID As String, SecretKey As String)
    Set m_Linkhub = New Linkhub
    
    m_Linkhub.LinkID = LinkID
    m_Linkhub.SercetKey = SecretKey
    
    m_scope.Add "member"
    ''문자메시지
    m_scope.Add "150"
    m_scope.Add "151"
    m_scope.Add "152"
    
    m_IsTest = False
    
End Sub
Private Function getSession_token(CorpNum As String) As String
    Dim refresh As Boolean
    
    refresh = False
    Dim m_Token As LinkhubToken
    
    Set m_Token = Nothing
    
    If m_token_Dic.Exists(CorpNum) Then
        Set m_Token = m_token_Dic.Item(CorpNum)
    End If
    
    If m_Token Is Nothing Then
        refresh = True
    Else
        Dim utcnow As String
        utcnow = m_Linkhub.GetTime()
        refresh = m_Token.expiration < utcnow
    End If
    
    If refresh Then
    
        Set m_Token = m_Linkhub.getToken(IIf(m_IsTest, ServiceID_TEST, ServiceID_REAL), CorpNum, m_scope)
        
        If m_Token Is Nothing Then
            Err.Raise m_Linkhub.LastErrCode, "POPBILL", m_Linkhub.LastErrMessage
            Exit Function
        End If
        
        If m_token_Dic.Exists(CorpNum) Then
            m_token_Dic.Remove (CorpNum)
        End If

        m_token_Dic.Add CorpNum, m_Token
        
    End If
    
    getSession_token = m_Token.session_token

End Function

Private Function httpGET(url As String, BearerToken As String, Optional UserID As String) As Variant
     
    Dim winhttp1
    
    Set winhttp1 = CreateObject("MSXML2.XMLHTTP.6.0")
    Call winhttp1.Open("GET", IIf(m_IsTest, ServiceURL_TEST, ServiceURL_REAL) + url, False)
    
    Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
    Call winhttp1.setRequestHeader("x-pb-version", APIVersion)
    Call winhttp1.setRequestHeader("Accept-Encoding", "gzip,deflate")
    
    If UserID <> "" Then
        Call winhttp1.setRequestHeader("x-pb-userid", UserID)
    End If
    
    
    winhttp1.Send
    
    Dim Response As String
    Response = winhttp1.responseText
       
    Dim parsedDic As Object
    
    If winhttp1.Status <> 200 Then
    
        Set parsedDic = m_Linkhub.parse(Response)
        
        Err.Raise parsedDic.Item("code"), "POPBILL", parsedDic.Item("message")
        
        Exit Function
        
    End If
    
    Set winhttp1 = Nothing
    
    Set httpGET = m_Linkhub.parse(Response)

End Function

Private Function httpPOST(url As String, BearerToken As String, postdata As String, Optional UserID As String) As Variant
    Set httpPOST = httpPOST_override(url, BearerToken, "", postdata, UserID)
    
End Function

Private Function httpPOST_override(url As String, BearerToken As String, override As String, postdata As String, Optional UserID As String) As Variant
     
    Dim winhttp1
    
    Set winhttp1 = CreateObject("MSXML2.XMLHTTP.6.0")
    Call winhttp1.Open("POST", IIf(m_IsTest, ServiceURL_TEST, ServiceURL_REAL) + url, False)
    Call winhttp1.setRequestHeader("x-pb-version", APIVersion)
    Call winhttp1.setRequestHeader("Accept-Encoding", "gzip,deflate")
    
    If BearerToken <> "" Then
        Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
    End If
    
    If override <> "" Then
        Call winhttp1.setRequestHeader("X-HTTP-Method-Override", override)
    End If
    
    If UserID <> "" Then
        Call winhttp1.setRequestHeader("x-pb-userid", UserID)
    End If
    
    Call winhttp1.setRequestHeader("Content-Type", "Application/json")
    
    winhttp1.Send (postdata)
        
    Dim Response As String
    Response = winhttp1.responseText
       
    Dim parsedDic As Object
    
    If winhttp1.Status <> 200 Then
    
        Set parsedDic = m_Linkhub.parse(Response)
        
        Err.Raise parsedDic.Item("code"), "POPBILL", IIf(IsNull(parsedDic.Item("message")), "Null Exception", parsedDic.Item("message"))
        
        Exit Function
        
    End If
    
    Set winhttp1 = Nothing
    
    Set httpPOST_override = m_Linkhub.parse(Response)

End Function

Private Function GetOnlyFileName(ByVal FilePath As String) As String

     Dim Temp() As String
     Temp = Split(FilePath, "\")
     GetOnlyFileName = Split(FilePath, "\")(UBound(Temp))

End Function

Private Function httpPOST_File(url As String, BearerToken As String, postdata As String, FilePaths As Collection, Optional UserID As String) As Variant
    Dim winhttp1
    Dim boundary As String
    boundary = "---------------------popbill"
    
    Set winhttp1 = CreateObject("MSXML2.XMLHTTP.6.0")
    Call winhttp1.Open("POST", IIf(m_IsTest, ServiceURL_TEST, ServiceURL_REAL) + url, False)
    Call winhttp1.setRequestHeader("x-pb-version", APIVersion)
    Call winhttp1.setRequestHeader("Accept-Encoding", "gzip,deflate")
    
    
    If BearerToken <> "" Then
        Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
    End If
    
    If UserID <> "" Then
        Call winhttp1.setRequestHeader("x-pb-userid", UserID)
    End If
    
    Call winhttp1.setRequestHeader("Content-Type", "multipart/form-data; boundary=" + boundary)
    
    Dim applicationform As String
    
    If postdata <> "" Then
        applicationform = "--" & boundary & vbCrLf & _
                          "Content-Disposition: form-data; name=""form""" & vbCrLf & _
                          "Content-Type: Application/json" & vbCrLf & vbCrLf & _
                          postdata
    
    End If
    
    Dim btFilePart As New Collection
    
    'Dim btFilePart() As Byte
    Dim filePartSize As Long
    
    filePartSize = 0
    
    Dim fileCnt As Integer
    
    For fileCnt = 1 To FilePaths.Count
        Dim FilePath As String
        
        FilePath = FilePaths.Item(fileCnt)
        
        If FilePath = "" Then Err.Raise -99999999, "POPBILL", "전송파일 경로가 잘못되었습니다."
        
        Dim FileHeader As String
        
        FileHeader = vbCrLf & "--" & boundary & vbCrLf & _
               "Content-Disposition: form-data; name=""file""; filename=""" & GetOnlyFileName(FilePath) & """" + vbCrLf & _
               "Content-Type: application/octet-stream" & vbCrLf & vbCrLf
               
        
        '### 이미지 바이너리를 얻는다.
        Dim binary() As Byte
        Open FilePath For Binary Access Read As #1
             ReDim binary(LOF(1))
             Get #1, , binary
        Close #1
              
        Dim btFHeader() As Byte
        Dim headerSize, fileSize As Long
        btFHeader = UniStrToUTF8(FileHeader)
        
        headerSize = UBound(btFHeader) + 1
        fileSize = UBound(binary) - LBound(binary)
        Dim btFile() As Byte
        
        ReDim btFile(headerSize + fileSize)
        
        CopyMemory btFile(0), btFHeader(0), headerSize
        CopyMemory btFile(headerSize), binary(0), fileSize
        
        filePartSize = filePartSize + UBound(btFile)
        
        btFilePart.Add btFile
        
        
    Next
    
    
    Dim tail As String
    
    tail = vbCrLf & "--" & boundary & "--" & vbCrLf
    
    Dim btForm() As Byte, BtTail() As Byte, Buffer() As Byte
    Dim formSize, tailSize As Long
    
    btForm = StrConv(applicationform, vbFromUnicode)
    BtTail = StrConv(tail, vbFromUnicode)
    
    formSize = UBound(btForm) + 1
    tailSize = UBound(BtTail) + 1
    
    ReDim Buffer(formSize + filePartSize + tailSize)
    CopyMemory Buffer(0), btForm(0), formSize
    
    Dim i As Integer
    Dim offsetPost As Long
    offsetPost = formSize
    For i = 1 To btFilePart.Count
        Dim tmp() As Byte
        tmp = btFilePart.Item(i)
        
        CopyMemory Buffer(offsetPost), tmp(0), UBound(tmp)
        offsetPost = offsetPost + UBound(tmp)
    Next
    
    CopyMemory Buffer(formSize + filePartSize), BtTail(0), tailSize
    
    winhttp1.Send (Buffer)
        
    Dim Response As String
    Response = winhttp1.responseText
       
    Dim parsedDic As Object
    
    If winhttp1.Status <> 200 Then
    
        Set parsedDic = m_Linkhub.parse(Response)
        
        Err.Raise parsedDic.Item("code"), "POPBILL", IIf(IsNull(parsedDic.Item("message")), "Null Exception", parsedDic.Item("message"))
        
        Exit Function
        
    End If
    
    Set winhttp1 = Nothing
    
    Set httpPOST_File = m_Linkhub.parse(Response)
End Function

'연동회원 아이디 중복 확인
Public Function CheckID(IDString As String) As PBResponse
On Error GoTo ErrHandler
    Dim result As Variant
    
    Set result = httpGET("/IDCheck?ID=" + IDString, "")
    
    Set CheckID = New PBResponse
    CheckID.code = result.Item("code")
    CheckID.message = result.Item("message")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    Set CheckID = Nothing
End Function
'연동회원 담당자 목록조회
Public Function ListContact(CorpNum As String, Optional UserID As String = "") As Collection
On Error GoTo ErrHandler
        
    Dim result As Variant
       
    Set result = httpGET("/IDs", getSession_token(CorpNum), UserID)
    
    Set ListContact = New Collection
    
    Dim T As Variant
    Dim info As PBContactInfo
    
    For Each T In result
        Set info = New PBContactInfo
        
        info.fromDictionary T
        
        ListContact.Add info
    Next
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    Set ListContact = Nothing
End Function
'담당자 추가
Public Function RegistContact(CorpNum As String, ContactInfo As PBContactInfo, Optional UserID As String = "") As PBResponse
On Error GoTo ErrHandler
    Set RegistContact = Nothing
    
    Dim postdata As String
    Dim tmp
    
    Set tmp = CreateObject("Scripting.Dictionary")
    
    tmp.Add "id", ContactInfo.id
    tmp.Add "pwd", ContactInfo.pwd
    tmp.Add "personName", ContactInfo.personName
    tmp.Add "tel", ContactInfo.tel
    tmp.Add "hp", ContactInfo.hp
    tmp.Add "fax", ContactInfo.fax
    tmp.Add "email", ContactInfo.email
    tmp.Add "searchAllAllowYN", ContactInfo.searchAllAllowYN
    tmp.Add "mgrYN", ContactInfo.mgrYN
            
    postdata = m_Linkhub.toString(tmp)
    
    Dim result As Variant
    
    Set result = httpPOST("/IDs/New", getSession_token(CorpNum), postdata)
    
    Set RegistContact = New PBResponse
    
    RegistContact.code = result.Item("code")
    RegistContact.message = result.Item("message")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
End Function
'담당자 정보 수정
Public Function UpdateContact(CorpNum As String, ContactInfo As PBContactInfo, UserID As String) As PBResponse
On Error GoTo ErrHandler
    Set UpdateContact = Nothing
    
    Dim postdata As String
    Dim tmp
    
    Set tmp = CreateObject("Scripting.Dictionary")
    
    tmp.Add "id", ContactInfo.id
    tmp.Add "personName", ContactInfo.personName
    tmp.Add "tel", ContactInfo.tel
    tmp.Add "hp", ContactInfo.hp
    tmp.Add "fax", ContactInfo.fax
    tmp.Add "email", ContactInfo.email
    tmp.Add "searchAllAllowYN", ContactInfo.searchAllAllowYN
    tmp.Add "mgrYN", ContactInfo.mgrYN
            
    postdata = m_Linkhub.toString(tmp)
    
    Dim result As Variant
    
    Set result = httpPOST("/IDs", getSession_token(CorpNum), postdata, UserID)
    
    Set UpdateContact = New PBResponse
    
    UpdateContact.code = result.Item("code")
    UpdateContact.message = result.Item("message")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
End Function
'회사정보 수정
Public Function UpdateCorpInfo(CorpNum As String, CorpInfo As PBCorpInfo, Optional UserID As String = "") As PBResponse
On Error GoTo ErrHandler
    Set UpdateCorpInfo = Nothing
    
    Dim postdata As String
    Dim tmp
    
    Set tmp = CreateObject("Scripting.Dictionary")
    
    tmp.Add "ceoname", CorpInfo.ceoname
    tmp.Add "corpName", CorpInfo.corpName
    tmp.Add "addr", CorpInfo.addr
    tmp.Add "bizType", CorpInfo.bizType
    tmp.Add "bizClass", CorpInfo.bizClass
            
    postdata = m_Linkhub.toString(tmp)
    
    Dim result As Variant
    
    Set result = httpPOST("/CorpInfo", getSession_token(CorpNum), postdata, UserID)
    
    Set UpdateCorpInfo = New PBResponse
    
    UpdateCorpInfo.code = result.Item("code")
    UpdateCorpInfo.message = result.Item("message")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
End Function
'회사정보 조회
Public Function GetCorpInfo(CorpNum As String, Optional UserID As String = "") As PBCorpInfo
On Error GoTo ErrHandler
    Dim result As Variant
       
    Set result = httpGET("/CorpInfo", getSession_token(CorpNum), UserID)
    
    Set GetCorpInfo = New PBCorpInfo
        
    GetCorpInfo.fromDictionary result
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    Set GetCorpInfo = Nothing
End Function
'회원잔액조회
Public Function GetBalance(CorpNum As String) As Double
    On Error GoTo ErrHandler
    
    GetBalance = m_Linkhub.GetBalance(getSession_token(CorpNum), IIf(m_IsTest, ServiceID_TEST, ServiceID_REAL))
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    GetBalance = m_LastErrCode
End Function
'파트너 잔액조회
Public Function GetPartnerBalance(CorpNum As String) As Double
    On Error GoTo ErrHandler
    
    GetPartnerBalance = m_Linkhub.GetPartnerBalance(getSession_token(CorpNum), IIf(m_IsTest, ServiceID_TEST, ServiceID_REAL))
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    GetPartnerBalance = m_LastErrCode
End Function

'파트너 포인트 충전 URL - 2017/08/30 추가
Public Function GetPartnerURL(CorpNum As String, TOGO As String) As String
    On Error GoTo ErrHandler
    
    GetPartnerURL = m_Linkhub.GetPartnerURL(getSession_token(CorpNum), IIf(m_IsTest, ServiceID_TEST, ServiceID_REAL), TOGO)
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    GetPartnerURL = m_LastErrCode
End Function

'팝빌 기본 URL
Public Function GetPopbillURL(CorpNum As String, UserID As String, TOGO As String) As String
On Error GoTo ErrHandler
    Dim result As Variant
    
    Set result = httpGET("/?TG=" + TOGO, getSession_token(CorpNum), UserID)
    
    GetPopbillURL = result.Item("url")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    GetPopbillURL = ""
End Function
'회원가입 여부
Public Function CheckIsMember(CorpNum As String, LinkID As String) As PBResponse
On Error GoTo ErrHandler
    Dim result As Variant
    
    Set result = httpGET("/Join?CorpNum=" + CorpNum + "&LID=" + LinkID, "")
    
    Set CheckIsMember = New PBResponse
    CheckIsMember.code = result.Item("code")
    CheckIsMember.message = result.Item("message")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    Set CheckIsMember = Nothing
End Function
'회원가입
Public Function JoinMember(JoinInfo As PBJoinForm) As PBResponse
On Error GoTo ErrHandler
    Set JoinMember = Nothing
    
    Dim postdata As String
    Dim tmp
    
    Set tmp = CreateObject("Scripting.Dictionary")
    
    tmp.Add "LinkID", JoinInfo.LinkID
    tmp.Add "CorpNum", JoinInfo.CorpNum
    tmp.Add "CorpName", JoinInfo.corpName
    tmp.Add "CEOName", JoinInfo.ceoname
    tmp.Add "Addr", JoinInfo.addr
    tmp.Add "ZipCode", JoinInfo.ZipCode
    tmp.Add "BizClass", JoinInfo.bizClass
    tmp.Add "BizType", JoinInfo.bizType
    tmp.Add "ContactName", JoinInfo.ContactName
    tmp.Add "ContactEmail", JoinInfo.ContactEmail
    tmp.Add "ContactFAX", JoinInfo.ContactFAX
    tmp.Add "ContactHP", JoinInfo.ContactHP
    tmp.Add "ContactTEL", JoinInfo.ContactTEL
    tmp.Add "ID", JoinInfo.id
    tmp.Add "PWD", JoinInfo.pwd
    
    postdata = m_Linkhub.toString(tmp)
    
    Dim result As Variant
    
    Set result = httpPOST("/Join", "", postdata)
    
    Set JoinMember = New PBResponse
    
    JoinMember.code = result.Item("code")
    JoinMember.message = result.Item("message")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
End Function

'과금정보 조회
Public Function GetChargeInfo(CorpNum As String, MType As MsgType, Optional UserID As String = "") As PBChargeInfo
On Error GoTo ErrHandler
    Dim result As Variant
    Dim m_Type As String
    
    Select Case MType
        Case SMS
            m_Type = "SMS"
        Case LMS
            m_Type = "LMS"
        Case MMS
            m_Type = "MMS"
    End Select
       
    Set result = httpGET("/Message/ChargeInfo?Type=" & m_Type, getSession_token(CorpNum), UserID)
    
    Set GetChargeInfo = New PBChargeInfo
        
    GetChargeInfo.fromDictionary result
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    Set GetChargeInfo = Nothing
End Function
'''''''''''''''''''''''''' End Of PopbillBase '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''단가확인
Public Function GetUnitCost(CorpNum As String, MType As MsgType) As Single
On Error GoTo ErrHandler
    Dim result As Variant
    Dim m_Type As String
    Select Case MType
        Case SMS
            m_Type = "SMS"
        Case LMS
            m_Type = "LMS"
        Case MMS
            m_Type = "MMS"
    End Select
    
    Set result = httpGET("/Message/UnitCost?Type=" & m_Type, getSession_token(CorpNum))
    
    GetUnitCost = result.Item("unitCost")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    GetUnitCost = m_LastErrCode
End Function
'SendMMS
Public Function SendMMS(CorpNum As String, sender As String, subject As String, Contents As String, Messages As Collection, FilePaths As Collection, reserveDT As String, adsYN As Boolean, Optional UserID As String = "", Optional requestNum As String = "") As String
On Error GoTo ErrHandler
    
    Dim result As Variant
    If Messages Is Nothing Or Messages.Count = 0 Then
        Err.Raise -99999999, "POPBILL", "전송할 메시지가 입력되지 않았습니다."
    End If
    If FilePaths Is Nothing Or FilePaths.Count = 0 Then Err.Raise -99999999, "POPBILL", "전송할 파일경로가 입력되지 않았습니다."
    
    Dim request As Variant
    
    Set request = CreateObject("Scripting.Dictionary")
    If adsYN Then request.Add "adsYN", True
    If sender <> "" Then request.Add "snd", sender
    If Contents <> "" Then request.Add "content", Contents
    If subject <> "" Then request.Add "subject", subject
    If reserveDT <> "" Then request.Add "sndDT", reserveDT
    If requestNum <> "" Then request.Add "requestNum", requestNum
    
    Dim msgs As New Collection
    
    Dim msg As PBMessage
    
    For Each msg In Messages
    
        msgs.Add msg.toDictionary
    Next
    
    request.Add "msgs", msgs
    
    Dim postdata As String
    
    postdata = m_Linkhub.toString(request)
    
    Set result = httpPOST_File("/MMS", getSession_token(CorpNum), postdata, FilePaths, UserID)
    
    SendMMS = result.Item("receiptNum")
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    SendMMS = ""
End Function
'단문 전송
Public Function SendSMS(CorpNum As String, sender As String, Contents As String, Messages As Collection, reserveDT As String, adsYN As Boolean, Optional UserID As String = "", Optional requestNum As String = "") As String
    SendSMS = SendMessage("SMS", CorpNum, sender, "", Contents, Messages, reserveDT, adsYN, UserID, requestNum)
End Function
'장문전송
Public Function SendLMS(CorpNum As String, sender As String, subject As String, Contents As String, Messages As Collection, reserveDT As String, adsYN As Boolean, Optional UserID As String = "", Optional requestNum As String = "") As String
    SendLMS = SendMessage("LMS", CorpNum, sender, subject, Contents, Messages, reserveDT, adsYN, UserID, requestNum)
End Function
'단/장문 길이 자동인식 전송
Public Function SendXMS(CorpNum As String, sender As String, subject As String, Contents As String, Messages As Collection, reserveDT As String, adsYN As Boolean, Optional UserID As String = "", Optional requestNum As String = "") As String
    SendXMS = SendMessage("XMS", CorpNum, sender, subject, Contents, Messages, reserveDT, adsYN, UserID, requestNum)
End Function

Private Function SendMessage(MType As String, CorpNum As String, sender As String, subject As String, Contents As String, Messages As Collection, reserveDT As String, adsYN As Boolean, Optional UserID As String = "", Optional requestNum As String = "") As String
On Error GoTo ErrHandler
    Dim result As Variant
    If Messages Is Nothing Or Messages.Count = 0 Then
        Err.Raise -99999999, "POPBILL", "전송할 메시지가 입력되지 않았습니다."
    End If
    
    Dim request As Variant
    
    Set request = CreateObject("Scripting.Dictionary")
    If adsYN Then request.Add "adsYN", True
    
    If sender <> "" Then request.Add "snd", sender
    If Contents <> "" Then request.Add "content", Contents
    If subject <> "" Then request.Add "subject", subject
    If reserveDT <> "" Then request.Add "sndDT", reserveDT
    If requestNum <> "" Then request.Add "requestNum", requestNum
    
    Dim msgs As New Collection
    
    Dim msg As PBMessage
    
    For Each msg In Messages
    
        msgs.Add msg.toDictionary
    Next
    
    request.Add "msgs", msgs
    
    Dim postdata As String
    
    postdata = m_Linkhub.toString(request)
    
    Set result = httpPOST("/" & MType, getSession_token(CorpNum), postdata, UserID)
    
    SendMessage = result.Item("receiptNum")
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    SendMessage = ""
End Function

'전송내역 목록 조회
Public Function Search(CorpNum As String, SDate As String, EDate As String, state As Collection, Item As Collection, ReserveYN As Boolean, SenderYN As Boolean, Page As Integer, PerPage As Integer, Order As String) As PBSearchList
On Error GoTo ErrHandler
    Dim result As Variant
    Dim uri As String
    Dim i As Integer
            
    uri = "/Message/Search"
    
    If Len(SDate) = 0 Or IsNull(SDate) Then
        Err.Raise -99999999, "POPBILL", "시작일자가 입력되지 않았습니다"
    End If
    
    If Len(EDate) = 0 Or IsNull(EDate) Then
        Err.Raise -99999999, "POPBILL", "종료일자가 입력되지 않았습니다"
    End If
    
    uri = uri + "?SDate=" + SDate
    uri = uri + "&EDate=" + EDate
    
    If state.Count > 0 Then
        uri = uri + "&State="
        For i = 1 To state.Count
            uri = uri + state(i) + ","
        Next
        
    End If
        
    If Item.Count > 0 Then
        uri = uri + "&Item="
        For i = 1 To Item.Count
            uri = uri + Item(i) + ","
        Next
        
    End If
    
    If ReserveYN Then
        uri = uri + "&ReserveYN=1"
    End If
    
    If SenderYN Then
        uri = uri + "&SenderYN=1"
    End If
            
    Page = IIf(Page < 1, 1, Page)
    PerPage = IIf(PerPage < 1, 500, PerPage)
    
    uri = uri + "&Page=" + CStr(Page)
    uri = uri + "&PerPage=" + CStr(PerPage)
    uri = uri + "&Order=" + Order
        
    Set result = httpGET(uri, getSession_token(CorpNum), "")
    
    Set Search = New PBSearchList
       
    Search.fromDictionary result
        
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    Set Search = Nothing
End Function

'전송내역 확인
Public Function GetMessages(CorpNum As String, ReceiptNum As String, Optional UserID As String = "") As Collection
On Error GoTo ErrHandler
    Dim result As Variant
    If ReceiptNum = "" Then
        Err.Raise -99999999, "POPBILL", "접수번호가 입력되지 않았습니다."
    End If
    
    Set result = httpGET("/Message/" & ReceiptNum, getSession_token(CorpNum), UserID)
    
    Set GetMessages = New Collection
    
    Dim T As Variant
    Dim sentMsg As PBSentMsg
    
    For Each T In result
        Set sentMsg = New PBSentMsg
        
        sentMsg.fromDictionary T
        
        GetMessages.Add sentMsg
    Next
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    Set GetMessages = Nothing

End Function

'전송내역 확인(요청번호 할당)
Public Function GetMessagesRN(CorpNum As String, requestNum As String, Optional UserID As String = "") As Collection
On Error GoTo ErrHandler
    Dim result As Variant
    If requestNum = "" Then
        Err.Raise -99999999, "POPBILL", "요청번호가 입력되지 않았습니다."
    End If
    
    Set result = httpGET("/Message/Get/" & requestNum, getSession_token(CorpNum), UserID)
    
    Set GetMessagesRN = New Collection
    
    Dim T As Variant
    Dim sentMsg As PBSentMsg
    
    For Each T In result
        Set sentMsg = New PBSentMsg
        
        sentMsg.fromDictionary T
        
        GetMessagesRN.Add sentMsg
    Next
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    Set GetMessagesRN = Nothing

End Function

'전송내역 요약정보 확인
Public Function GetStates(CorpNum As String, ReceiptNumList As Collection, Optional UserID As String = "") As Collection
On Error GoTo ErrHandler

    If ReceiptNumList Is Nothing Or ReceiptNumList.Count = 0 Then
        Err.Raise -99999999, "POPBILL", "접수번호가 입력되지 않았습니다."
    End If
    
    Dim result As Variant
    
    Dim postdata As String
    
    postdata = m_Linkhub.toString(ReceiptNumList)
    
    Set result = httpPOST("/Message/States/", getSession_token(CorpNum), postdata, UserID)
    
    Set GetStates = New Collection
    
    Dim T As Variant
    Dim briefMsg As PBMessageBriefInfo
    
    For Each T In result
        Set briefMsg = New PBMessageBriefInfo
        
        briefMsg.fromDictionary T
        
        GetStates.Add briefMsg
    Next
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    Set GetStates = Nothing
End Function


'예약전송 취소
Public Function CancelReserve(CorpNum As String, ReceiptNum As String, Optional UserID As String = "") As PBResponse
On Error GoTo ErrHandler
    Dim result As Variant
    If ReceiptNum = "" Then
        Err.Raise -99999999, "POPBILL", "접수번호가 입력되지 않았습니다."
    End If
    
    Set result = httpGET("/Message/" & ReceiptNum & "/Cancel", getSession_token(CorpNum), UserID)
    Set CancelReserve = New PBResponse
    
    CancelReserve.code = result.Item("code")
    CancelReserve.message = result.Item("message")
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    Set CancelReserve = Nothing

End Function

'예약전송 취소 (요청번호 할당)
Public Function CancelReserveRN(CorpNum As String, requestNum As String, Optional UserID As String = "") As PBResponse
On Error GoTo ErrHandler
    Dim result As Variant
    If requestNum = "" Then
        Err.Raise -99999999, "POPBILL", "요청번호가 입력되지 않았습니다."
    End If
    
    Set result = httpGET("/Message/Cancel/" & requestNum, getSession_token(CorpNum), UserID)
    Set CancelReserveRN = New PBResponse
    
    CancelReserveRN.code = result.Item("code")
    CancelReserveRN.message = result.Item("message")
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    Set CancelReserveRN = Nothing

End Function

'문자 관련 URL
Public Function GetURL(CorpNum As String, UserID As String, TOGO As String) As String
On Error GoTo ErrHandler
    Dim result As Variant
    
    Set result = httpGET("/Message/?TG=" + TOGO, getSession_token(CorpNum), UserID)
    
    GetURL = result.Item("url")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    GetURL = ""
End Function

'080 수신거부 목록 확인
Public Function GetAutoDenyList(CorpNum As String) As Collection
On Error GoTo ErrHandler

    Dim result As Variant
    
    Set result = httpGET("/Message/Denied", getSession_token(CorpNum), "")
    
    Set GetAutoDenyList = New Collection
    
    Dim T As Variant
    Dim info As PBAutoDenyInfo
    
    For Each T In result
        Set info = New PBAutoDenyInfo
        info.fromDictionary T
        GetAutoDenyList.Add info
    Next
    
    Exit Function
   
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    Set GetAutoDenyList = Nothing
End Function

'발신번호 목록 조회
Public Function GetSenderNumberList(CorpNum As String, Optional UserID As String = "") As Collection
On Error GoTo ErrHandler

    Dim result As Variant
    
    Set result = httpGET("/Message/SenderNumber", getSession_token(CorpNum), UserID)
    
    Set GetSenderNumberList = New Collection
    
    Dim T As Variant
    Dim info As PBMsgSenderNumber
    
    For Each T In result
        Set info = New PBMsgSenderNumber
        info.fromDictionary T
        GetSenderNumberList.Add info
    Next
    
    Exit Function
   
ErrHandler:
    m_LastErrCode = Err.number
    m_LastErrMessage = Err.Description
    Set GetSenderNumberList = Nothing
End Function


