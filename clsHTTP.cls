VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHTTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
''
Private sURL As String
Private sLastError As String
Private sContentType As String
Private lContentLength As Long
Private sLastModified As String
Private sVersion As String
Private sStatusCode As String
Private sStatusText As String
Private sRawHeaders As String
Private sResponseHeaders As String
Private sForwarded As String
Private sServer As String
Private sRequestMethod As String
Private sPragma As String
Public sRequestHeaders As String
Private sUserAgent As String
Private sRequestMethod2 As String

Private hInternetSession As Long
Private hInternetConnect As Long
Private hHttpOpenRequest As Long

Private lMajor As Long
Private lMinor As Long

Event Progress(ProcentDone As Byte, Speed As Currency, Downloaded As Long)

' Initializes an application's use of the Win32 Internet functions
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

' Use registry access clsSettings.
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0

' Opens a HTTP session for a given site.
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long

' Number of the TCP/IP port on the server to connect to.
Private Const INTERNET_DEFAULT_FTP_PORT = 21
Private Const INTERNET_DEFAULT_GOPHER_PORT = 70
Private Const INTERNET_DEFAULT_HTTP_PORT = 80
Private Const INTERNET_DEFAULT_HTTPS_PORT = 443
Private Const INTERNET_DEFAULT_SOCKS_PORT = 1080

' Type of service to access.
Private Const INTERNET_SERVICE_FTP = 1
Private Const INTERNET_SERVICE_GOPHER = 2
Private Const INTERNET_SERVICE_HTTP = 3

' Opens an HTTP request handle.
Private Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" (ByVal hHttpSession As Long, ByVal sVerb As String, ByVal sObjectName As String, ByVal sVersion As String, ByVal sReferer As String, ByVal something As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long

' Brings the data across the wire even if it locally cached.
Private Const INTERNET_FLAG_RELOAD = &H80000000

' Sends the specified request to the HTTP server.
Private Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, sOptional As Any, ByVal lOptionalLength As Long) As Integer

' Queries for information about an HTTP request.
Private Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" (ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long, ByRef lIndex As Long) As Integer

' The possible values for the lInfoLevel parameter include:
Private Const HTTP_QUERY_CONTENT_TYPE = 1
Private Const HTTP_QUERY_CONTENT_LENGTH = 5
Private Const HTTP_QUERY_EXPIRES = 10
Private Const HTTP_QUERY_LAST_MODIFIED = 11
Private Const HTTP_QUERY_PRAGMA = 17
Private Const HTTP_QUERY_VERSION = 18
Private Const HTTP_QUERY_STATUS_CODE = 19
Private Const HTTP_QUERY_STATUS_TEXT = 20
Private Const HTTP_QUERY_RAW_HEADERS = 21
Private Const HTTP_QUERY_RAW_HEADERS_CRLF = 22
Private Const HTTP_QUERY_FORWARDED = 30
Private Const HTTP_QUERY_SERVER = 37
Private Const HTTP_QUERY_USER_AGENT = 39
Private Const HTTP_QUERY_SET_COOKIE = 43
Private Const HTTP_QUERY_REQUEST_METHOD = 45

' Add this flag to the about flags to get request header.
Private Const HTTP_QUERY_FLAG_REQUEST_HEADERS = &H80000000

' Reads data from a handle opened by the HttpOpenRequest function.
Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer

' Closes a single Internet handle or a subtree of Internet handles.
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer

' Queries an Internet option on the specified handle
Private Declare Function InternetQueryOption Lib "wininet.dll" Alias "InternetQueryOptionA" (ByVal hInternet As Long, ByVal lOption As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long) As Integer

' Returns the version number of Wininet.dll.
Private Const INTERNET_OPTION_VERSION = 40

' Contains the version number of the DLL that contains the Windows Internet
' functions (Wininet.dll). This structure is used when passing the
' INTERNET_OPTION_VERSION flag to the InternetQueryOption function.
Private Type tWinInetDLLVersion
  lMajorVersion As Long
  lMinorVersion As Long
End Type

' Adds one or more HTTP request headers to the HTTP request handle.
Private Declare Function HttpAddRequestHeaders Lib "wininet.dll" Alias "HttpAddRequestHeadersA" (ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lModifiers As Long) As Integer

' Flags to modify the semantics of this function. Can be a combination of these values:

' Adds the header only if it does not already exist; otherwise, an error is returned.
Private Const HTTP_ADDREQ_FLAG_ADD_IF_NEW = &H10000000

' Adds the header if it does not exist. Used with REPLACE.
Private Const HTTP_ADDREQ_FLAG_ADD = &H20000000

' Replaces or removes a header. If the header value is empty and the header is found,
' it is removed. If not empty, the header value is replaced
Private Const HTTP_ADDREQ_FLAG_REPLACE = &H80000000

Public Property Get Forwarded() As String

  Forwarded = sForwarded

End Property

Function Connect(Url As String) As Boolean

  Dim vDllVersion As tWinInetDLLVersion
  Dim iRetVal     As Integer
  Dim UrlAddress As String
  Dim UrlServer  As String
  
  If (Left$(Url, 7) = "http://") Then Url = Mid$(Url, 8)
  
  If (InStr(Url, "/") = 0) Then
    UrlServer = Url
    UrlAddress = ""
  Else
    UrlServer = Left$(Url, InStr(Url, "/") - 1)
    UrlAddress = Mid$(Url, Len(UrlServer) + 1)
  End If
  
  sURL = ""
  'Screen.MousePointer = vbHourglass
  
  If CBool(hInternetSession) Then
    InternetQueryOption hInternetSession, INTERNET_OPTION_VERSION, vDllVersion, Len(vDllVersion)
    lMajor = vDllVersion.lMajorVersion
    lMinor = vDllVersion.lMinorVersion
    hInternetConnect = InternetConnect(hInternetSession, UrlServer, INTERNET_DEFAULT_HTTP_PORT, vbNullString, vbNullString, INTERNET_SERVICE_HTTP, 0, 0)
    
    If (hInternetConnect > 0) Then
      hHttpOpenRequest = HttpOpenRequest(hInternetConnect, "GET", UrlAddress, "HTTP/1.0", vbNullString, 0, INTERNET_FLAG_RELOAD, 0)
        
      If CBool(hHttpOpenRequest) Then
        iRetVal = HttpSendRequest(hHttpOpenRequest, vbNullString, 0, 0, 0)

        If iRetVal Then
          'Response headers
          sContentType = RemNull(GetQueryInfo(hHttpOpenRequest, HTTP_QUERY_CONTENT_TYPE))
          lContentLength = Val(GetQueryInfo(hHttpOpenRequest, HTTP_QUERY_CONTENT_LENGTH))
          sLastModified = GetQueryInfo(hHttpOpenRequest, HTTP_QUERY_LAST_MODIFIED)
          sVersion = GetQueryInfo(hHttpOpenRequest, HTTP_QUERY_VERSION)
          sStatusCode = GetQueryInfo(hHttpOpenRequest, HTTP_QUERY_STATUS_CODE)
          sStatusText = GetQueryInfo(hHttpOpenRequest, HTTP_QUERY_STATUS_TEXT)
          sRawHeaders = GetQueryInfo(hHttpOpenRequest, HTTP_QUERY_RAW_HEADERS)
          sResponseHeaders = GetQueryInfo(hHttpOpenRequest, HTTP_QUERY_RAW_HEADERS_CRLF)
          sForwarded = GetQueryInfo(hHttpOpenRequest, HTTP_QUERY_FORWARDED)
          sServer = GetQueryInfo(hHttpOpenRequest, HTTP_QUERY_SERVER)
          sRequestMethod = GetQueryInfo(hHttpOpenRequest, HTTP_QUERY_REQUEST_METHOD)
          sPragma = GetQueryInfo(hHttpOpenRequest, HTTP_QUERY_FLAG_REQUEST_HEADERS + HTTP_QUERY_PRAGMA)
          sRequestHeaders = GetQueryInfo(hHttpOpenRequest, HTTP_QUERY_FLAG_REQUEST_HEADERS + HTTP_QUERY_RAW_HEADERS_CRLF)
          sUserAgent = GetQueryInfo(hHttpOpenRequest, HTTP_QUERY_FLAG_REQUEST_HEADERS + HTTP_QUERY_USER_AGENT)
          sRequestMethod2 = GetQueryInfo(hHttpOpenRequest, HTTP_QUERY_FLAG_REQUEST_HEADERS + HTTP_QUERY_REQUEST_METHOD)
          Connect = True
        Else
          'HttpSendRequest failed
          sLastError = "HttpSendRequest failed"
          Connect = False
        End If
      Else
        'HttpOpenRequest failed
        sLastError = "HttpOpenRequest failed"
        Connect = False
        End If
    Else
      'InternetConnect failed
      sLastError = "InternetConnect failed"
      Connect = False
    End If
  Else
    'hInternetSession handle not allocated
    sLastError = "InternetSession failed"
    Connect = False
  End If
  
  sLastError = ""
  sURL = Url
  'Screen.MousePointer = vbDefault

End Function
Public Property Get ContentLength() As Long

  ContentLength = lContentLength
  
End Property
Private Function GetQueryInfo(ByVal hHttpRequest As Long, ByVal iInfoLevel As Long) As String

  ' Purpose    : Retrieves header information
  ' Parameters : HttpRequest handle, label to put information in, and infolevel flag
  ' Return val : Success/Failure
  ' Algorithm  : Calls the HttpQueryInfo function and copies result into label

  Dim sBuffer         As String * 1024
  Dim lBufferLength   As Long

  lBufferLength = Len(sBuffer)
  Call HttpQueryInfo(hHttpRequest, iInfoLevel, ByVal sBuffer, lBufferLength, 0)
  GetQueryInfo = Trim$(sBuffer)

End Function
Function GetFile(ByVal Url As String, Optional sFilename As String, Optional sContent As String, Optional ReadToString As Boolean) As Boolean

  ' Purpose    : Click event for button
  ' Parameters : NA
  ' Return val : NA
  ' Algorithm  : Gets resource identified in sUrl

  Dim bDoLoop             As Boolean
  Dim sReadBuffer         As String * 2048
  Dim sBuffer             As String
  Dim lNumberOfBytesRead  As Long
  Dim Procent             As Byte
  Dim ProcentOld          As Byte
  Dim sUrlAddress         As String
  Dim sUrlServer          As String
  Dim lContentRead        As Long
  Dim timerStart          As Currency
  Dim lLoops              As Byte
  Dim Filenr As Long
  Dim Text As String
  
  If Not Connect(Url) Then
    GetFile = False
    Exit Function
  End If
  DoEvents
  
  If (InStr(sURL, "/") = 0) Then
    sUrlServer = sURL
    sUrlAddress = ""
  Else
    sUrlServer = Left$(sURL, InStr(sURL, "/") - 1)
    sUrlAddress = Mid$(sURL, Len(sUrlServer) + 1)
  End If
  
  On Error GoTo Errorhandler
  sLastError = ""
  lContentRead = 0
  ProcentOld = 0
  DoEvents
  RaiseEvent Progress(0, 0, 0)
  timerStart = Timer
  If (Len(sFilename) = 0) Then ReadToString = True
  
  If ReadToString Then
    sContent = ""
  Else
    If FileExist(sFilename) Then Kill sFilename
    Filenr = FreeFile
    Open sFilename For Binary Access Write As #Filenr
  End If
  
  Do
    lLoops = lLoops + 1
    sReadBuffer = vbNullString
    bDoLoop = InternetReadFile(hHttpOpenRequest, sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
  
    If Not ReadToString Then
      Put #Filenr, , Left$(sReadBuffer, lNumberOfBytesRead)
    Else
      sContent = sContent & Left$(sReadBuffer, lNumberOfBytesRead)
    End If
    DoEvents
    
    lContentRead = lContentRead + lNumberOfBytesRead
    If (lContentLength > 0) Then
      Procent = Int((lContentRead / lContentLength) * 100)
      If Procent <> ProcentOld Then
        RaiseEvent Progress(Procent, lContentRead / (Timer - timerStart), lContentRead)
        ProcentOld = Procent
      End If
    ElseIf (lLoops >= 10) Then
      lLoops = 0
      RaiseEvent Progress(0, lContentRead / (Timer - timerStart), lContentRead)
    End If
    If Not CBool(lNumberOfBytesRead) Then bDoLoop = False
    DoEvents
  Loop While bDoLoop

  If Not ReadToString Then
    Close #Filenr
  End If
  
  On Error Resume Next
  Call InternetCloseHandle(hHttpOpenRequest)
  
  GetFile = True
  Exit Function
  
Errorhandler:
  On Error Resume Next
  If Not ReadToString Then
    Close #Filenr
  End If
  GetFile = False

End Function
Function FileExist(FileName As String) As Boolean

  On Error GoTo Errorhandler
  
  Call FileLen(FileName)
  FileExist = True
  Exit Function

Errorhandler:
  FileExist = False

End Function
Public Property Get LastError() As String

  LastError = sLastError

End Property
Public Property Get LastModified() As String

  LastModified = sLastModified
  
End Property

Private Function RemNull(Text As String) As String

  RemNull = Left$(Text, InStr(Text, Chr$(0)) - 1)
  
End Function

Private Sub Class_Initialize()

  hInternetSession = InternetOpen("http protocol", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
  If CBool(hInternetSession) Then
    sLastError = ""
  Else
    sLastError = "InternetOpen failed."
  End If
  
End Sub

Private Sub Class_Terminate()

  On Error Resume Next
  Call InternetCloseHandle(hHttpOpenRequest)
  Call InternetCloseHandle(hInternetSession)
  Call InternetCloseHandle(hInternetConnect)
  
End Sub


