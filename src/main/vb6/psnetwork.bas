Attribute VB_Name = "Network"
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2003
'
'****************************************************************************
'
' LANGUAGE:             Microsoft Visual Basic V6
'
' MODULE NAME:          Pivotal_Network
'
' MODULE TYPE:          BASIC Module
'
' FILE NAME:            PSNETWORK.BAS
'
' MODIFICATION HISTORY: Steve O'Hara    03 January 2003   First created for MediaWeb
'
' PURPOSE:              Provides an interface into the WININET DLL
'
'
'****************************************************************************
'
'****************************************************
' MODULE VARIABLE DECLARATIONS
'****************************************************
'
Option Explicit


    '
    ' Module error values
    '
    Public Const ERROR_OFFSET = 10000
    Public Const ERROR_SOURCE = "PSNETWORK"


Declare Function GetProcessHeap Lib "kernel32" () As Long
Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Public Const HEAP_ZERO_MEMORY = &H8
Public Const HEAP_GENERATE_EXCEPTIONS = &H4

Declare Sub CopyMemory1 Lib "kernel32" Alias "RtlMoveMemory" ( _
         hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Declare Sub CopyMemory2 Lib "kernel32" Alias "RtlMoveMemory" ( _
         hpvDest As Long, hpvSource As Any, ByVal cbCopy As Long)

Public Const MAX_PATH = 260
Public Const NO_ERROR = 0
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
Public Const FILE_ATTRIBUTE_OFFLINE = &H1000

Private Type WIN32_FIND_DATA
            dwFileAttributes As Long
            ftCreationTime As Currency
            ftLastAccessTime As Currency
            ftLastWriteTime As Currency
            nFileSizeHigh As Long
            nFileSizeLow As Long
            dwReserved0 As Long
            dwReserved1 As Long
            cFileName As String * 260
            cAlternate As String * 14
    End Type

Public Const ERROR_NO_MORE_FILES = 18

Public Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" _
    (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long
    
Public Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" _
(ByVal hFtpSession As Long, ByVal lpszSearchFile As String, _
      lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long

Public Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" _
(ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, _
      ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, _
      ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean

Public Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" _
(ByVal hFtpSession As Long, ByVal lpszLocalFile As String, _
      ByVal lpszRemoteFile As String, _
      ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean

Public Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" _
    (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Public Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" _
    (ByVal hFtpSession As Long, ByVal lpszDirectory As String, ByRef lpdwCurrentDirectory As Long) As Boolean
' Initializes an application's use of the Win32 Internet functions
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
(ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, _
ByVal sProxyByPass As String, ByVal lFlags As Long) As Long

' User agent constant.
Public Const scUserAgent = "vb wininet"

' Use registry access settings.
Public Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Public Const INTERNET_OPEN_TYPE_DIRECT = 1
Public Const INTERNET_OPEN_TYPE_PROXY = 3
Public Const INTERNET_INVALID_PORT_NUMBER = 0

Public Const FTP_TRANSFER_TYPE_ASCII = &H1
Public Const FTP_TRANSFER_TYPE_BINARY = &H2

' Opens a HTTP session for a given site.
Public Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" _
(ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, _
ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, _
ByVal lFlags As Long, ByVal lContext As Long) As Long
                
Public Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" ( _
    lpdwError As Long, _
    ByVal lpszBuffer As String, _
    lpdwBufferLength As Long) As Boolean

' Number of the TCP/IP port on the server to connect to.
Public Const INTERNET_DEFAULT_FTP_PORT = 21
Public Const INTERNET_DEFAULT_GOPHER_PORT = 70
Public Const INTERNET_DEFAULT_HTTP_PORT = 80
Public Const INTERNET_DEFAULT_HTTPS_PORT = 443
Public Const INTERNET_DEFAULT_SOCKS_PORT = 1080

Public Const INTERNET_OPTION_CONNECT_TIMEOUT = 2
Public Const INTERNET_OPTION_RECEIVE_TIMEOUT = 6
Public Const INTERNET_OPTION_SEND_TIMEOUT = 5

Public Const INTERNET_OPTION_SECURITY_FLAGS = 31
    
Public Const SECURITY_FLAG_IGNORE_CERT_CN_INVALID = &H1000
Public Const SECURITY_FLAG_IGNORE_CERT_DATE_INVALID = &H2000
Public Const SECURITY_FLAG_IGNORE_UNKNOWN_CA = &H100

Public Const INTERNET_OPTION_USERNAME = 28
Public Const INTERNET_OPTION_PASSWORD = 29
Public Const INTERNET_OPTION_PROXY_USERNAME = 43
Public Const INTERNET_OPTION_PROXY_PASSWORD = 44

' Type of service to access.
Public Const INTERNET_SERVICE_FTP = 1
Public Const INTERNET_SERVICE_GOPHER = 2
Public Const INTERNET_SERVICE_HTTP = 3

' Opens an HTTP request handle.
Public Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" _
(ByVal hHttpSession As Long, ByVal sVerb As String, ByVal sObjectName As String, ByVal sVersion As String, _
ByVal sReferer As String, ByVal something As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long

Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000

' Sends the specified request to the HTTP server.
Public Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal _
hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal sOptional As _
String, ByVal lOptionalLength As Long) As Integer


' Queries for information about an HTTP request.
Public Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" _
(ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, _
ByRef lBufferLength As Long, ByRef lIndex As Long) As Integer

' InternetErrorDlg
Public Declare Function InternetErrorDlg Lib "wininet.dll" _
(ByVal hWnd As Long, ByVal hInternet As Long, ByVal dwError As Long, ByVal dwFlags As Long, ByVal lppvData As Long) As Long

' InternetErrorDlg constants
Public Const FLAGS_ERROR_UI_FILTER_FOR_ERRORS = &H1
Public Const FLAGS_ERROR_UI_FLAGS_CHANGE_OPTIONS = &H2
Public Const FLAGS_ERROR_UI_FLAGS_GENERATE_DATA = &H4
Public Const FLAGS_ERROR_UI_FLAGS_NO_UI = &H8
Public Const FLAGS_ERROR_UI_SERIALIZE_DIALOGS = &H10

Public Declare Function GetDesktopWindow Lib "USER32.DLL" () As Long

' The possible values for the lInfoLevel parameter include:
Public Const HTTP_QUERY_CONTENT_TYPE = 1
Public Const HTTP_QUERY_CONTENT_LENGTH = 5
Public Const HTTP_QUERY_EXPIRES = 10
Public Const HTTP_QUERY_LAST_MODIFIED = 11
Public Const HTTP_QUERY_PRAGMA = 17
Public Const HTTP_QUERY_VERSION = 18
Public Const HTTP_QUERY_STATUS_CODE = 19
Public Const HTTP_QUERY_STATUS_TEXT = 20
Public Const HTTP_QUERY_RAW_HEADERS = 21
Public Const HTTP_QUERY_RAW_HEADERS_CRLF = 22
Public Const HTTP_QUERY_FORWARDED = 30
Public Const HTTP_QUERY_SERVER = 37
Public Const HTTP_QUERY_USER_AGENT = 39
Public Const HTTP_QUERY_SET_COOKIE = 43
Public Const HTTP_QUERY_REQUEST_METHOD = 45

' Add this flag to the about flags to get request header.
Public Const HTTP_QUERY_FLAG_REQUEST_HEADERS = &H80000000
Public Const HTTP_QUERY_FLAG_NUMBER = &H20000000
' Reads data from a handle opened by the HttpOpenRequest function.
Public Declare Function InternetReadFile Lib "wininet.dll" _
(ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, _
lNumberOfBytesRead As Long) As Integer

Public Declare Function InternetReadBinaryFile Lib "wininet.dll" Alias "InternetReadFile" _
(ByVal hFile As Long, ByRef bytearray_firstelement As Byte, ByVal lNumBytesToRead As Long, _
ByRef lNumberOfBytesRead As Long) As Integer

Public Type INTERNET_BUFFERS
    dwStructSize As Long        ' used for API versioning. Set to sizeof(INTERNET_BUFFERS)
    Next As Long                ' INTERNET_BUFFERS chain of buffers
    lpcszHeader As Long       ' pointer to headers (may be NULL)
    dwHeadersLength As Long     ' length of headers if not NULL
    dwHeadersTotal As Long      ' size of headers if not enough buffer
    lpvBuffer As Long           ' pointer to data buffer (may be NULL)
    dwBufferLength As Long      ' length of data buffer if not NULL
    dwBufferTotal As Long       ' total size of chunk, or content-length if not chunked
    dwOffsetLow As Long         ' used for read-ranges (only used in HttpSendRequest2)
    dwOffsetHigh As Long
End Type

Public Declare Function HttpSendRequestEx Lib "wininet.dll" Alias "HttpSendRequestExA" _
(ByVal hHttpRequest As Long, lpBuffersIn As INTERNET_BUFFERS, ByVal lpBuffersOut As Long, _
ByVal dwFlags As Long, ByVal dwContext As Long) As Long

Public Declare Function HttpEndRequest Lib "wininet.dll" Alias "HttpEndRequestA" _
(ByVal hHttpRequest As Long, ByVal lpBuffersOut As Long, _
ByVal dwFlags As Long, ByVal dwContext As Long) As Long


Public Declare Function InternetWriteFile Lib "wininet.dll" _
        (ByVal hFile As Long, ByVal sBuffer As String, _
        ByVal lNumberOfBytesToRead As Long, _
        lNumberOfBytesRead As Long) As Integer

Public Declare Function FtpOpenFile Lib "wininet.dll" Alias _
        "FtpOpenFileA" (ByVal hFtpSession As Long, _
        ByVal sFilename As String, ByVal lAccess As Long, _
        ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function FtpDeleteFile Lib "wininet.dll" _
    Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, _
    ByVal lpszFileName As String) As Boolean
Public Declare Function InternetSetOption Lib "wininet.dll" Alias "InternetSetOptionA" _
(ByVal hInternet As Long, ByVal lOption As Long, ByRef sBuffer As Any, ByVal lBufferLength As Long) As Integer
Public Declare Function InternetSetOptionStr Lib "wininet.dll" Alias "InternetSetOptionA" _
(ByVal hInternet As Long, ByVal lOption As Long, ByVal sBuffer As String, ByVal lBufferLength As Long) As Integer

' Closes a single Internet handle or a subtree of Internet handles.
Public Declare Function InternetCloseHandle Lib "wininet.dll" _
(ByVal hInet As Long) As Integer

' Queries an Internet option on the specified handle
Public Declare Function InternetQueryOption Lib "wininet.dll" Alias "InternetQueryOptionA" _
(ByVal hInternet As Long, ByVal lOption As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long) As Integer

' Returns the version number of Wininet.dll.
Public Const INTERNET_OPTION_VERSION = 40

' Contains the version number of the DLL that contains the Windows Internet
' functions (Wininet.dll). This structure is used when passing the
' INTERNET_OPTION_VERSION flag to the InternetQueryOption function.
Public Type tWinInetDLLVersion
    lMajorVersion As Long
    lMinorVersion As Long
End Type

' Adds one or more HTTP request headers to the HTTP request handle.
Public Declare Function HttpAddRequestHeaders Lib "wininet.dll" Alias "HttpAddRequestHeadersA" _
(ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, _
ByVal lModifiers As Long) As Integer

' Flags to modify the semantics of this function. Can be a combination of these values:

' Adds the header only if it does not already exist; otherwise, an error is returned.
Public Const HTTP_ADDREQ_FLAG_ADD_IF_NEW = &H10000000

' Adds the header if it does not exist. Used with REPLACE.
Public Const HTTP_ADDREQ_FLAG_ADD = &H20000000

' Replaces or removes a header. If the header value is empty and the header is found,
' it is removed. If not empty, the header value is replaced
Public Const HTTP_ADDREQ_FLAG_REPLACE = &H80000000

' Internet Errors
Public Const INTERNET_ERROR_BASE = 12000

Public Const ERROR_INTERNET_OUT_OF_HANDLES = (INTERNET_ERROR_BASE + 1)
Public Const ERROR_INTERNET_TIMEOUT = (INTERNET_ERROR_BASE + 2)
Public Const ERROR_INTERNET_EXTENDED_ERROR = (INTERNET_ERROR_BASE + 3)
Public Const ERROR_INTERNET_INTERNAL_ERROR = (INTERNET_ERROR_BASE + 4)
Public Const ERROR_INTERNET_INVALID_URL = (INTERNET_ERROR_BASE + 5)
Public Const ERROR_INTERNET_UNRECOGNIZED_SCHEME = (INTERNET_ERROR_BASE + 6)
Public Const ERROR_INTERNET_NAME_NOT_RESOLVED = (INTERNET_ERROR_BASE + 7)
Public Const ERROR_INTERNET_PROTOCOL_NOT_FOUND = (INTERNET_ERROR_BASE + 8)
Public Const ERROR_INTERNET_INVALID_OPTION = (INTERNET_ERROR_BASE + 9)
Public Const ERROR_INTERNET_BAD_OPTION_LENGTH = (INTERNET_ERROR_BASE + 10)
Public Const ERROR_INTERNET_OPTION_NOT_SETTABLE = (INTERNET_ERROR_BASE + 11)
Public Const ERROR_INTERNET_SHUTDOWN = (INTERNET_ERROR_BASE + 12)
Public Const ERROR_INTERNET_INCORRECT_USER_NAME = (INTERNET_ERROR_BASE + 13)
Public Const ERROR_INTERNET_INCORRECT_PASSWORD = (INTERNET_ERROR_BASE + 14)
Public Const ERROR_INTERNET_LOGIN_FAILURE = (INTERNET_ERROR_BASE + 15)
Public Const ERROR_INTERNET_INVALID_OPERATION = (INTERNET_ERROR_BASE + 16)
Public Const ERROR_INTERNET_OPERATION_CANCELLED = (INTERNET_ERROR_BASE + 17)
Public Const ERROR_INTERNET_INCORRECT_HANDLE_TYPE = (INTERNET_ERROR_BASE + 18)
Public Const ERROR_INTERNET_INCORRECT_HANDLE_STATE = (INTERNET_ERROR_BASE + 19)
Public Const ERROR_INTERNET_NOT_PROXY_REQUEST = (INTERNET_ERROR_BASE + 20)
Public Const ERROR_INTERNET_REGISTRY_VALUE_NOT_FOUND = (INTERNET_ERROR_BASE + 21)
Public Const ERROR_INTERNET_BAD_REGISTRY_PARAMETER = (INTERNET_ERROR_BASE + 22)
Public Const ERROR_INTERNET_NO_DIRECT_ACCESS = (INTERNET_ERROR_BASE + 23)
Public Const ERROR_INTERNET_NO_CONTEXT = (INTERNET_ERROR_BASE + 24)
Public Const ERROR_INTERNET_NO_CALLBACK = (INTERNET_ERROR_BASE + 25)
Public Const ERROR_INTERNET_REQUEST_PENDING = (INTERNET_ERROR_BASE + 26)
Public Const ERROR_INTERNET_INCORRECT_FORMAT = (INTERNET_ERROR_BASE + 27)
Public Const ERROR_INTERNET_ITEM_NOT_FOUND = (INTERNET_ERROR_BASE + 28)
Public Const ERROR_INTERNET_CANNOT_CONNECT = (INTERNET_ERROR_BASE + 29)
Public Const ERROR_INTERNET_CONNECTION_ABORTED = (INTERNET_ERROR_BASE + 30)
Public Const ERROR_INTERNET_CONNECTION_RESET = (INTERNET_ERROR_BASE + 31)
Public Const ERROR_INTERNET_FORCE_RETRY = (INTERNET_ERROR_BASE + 32)
Public Const ERROR_INTERNET_INVALID_PROXY_REQUEST = (INTERNET_ERROR_BASE + 33)
Public Const ERROR_INTERNET_NEED_UI = (INTERNET_ERROR_BASE + 34)

Public Const ERROR_INTERNET_HANDLE_EXISTS = (INTERNET_ERROR_BASE + 36)
Public Const ERROR_INTERNET_SEC_CERT_DATE_INVALID = (INTERNET_ERROR_BASE + 37)
Public Const ERROR_INTERNET_SEC_CERT_CN_INVALID = (INTERNET_ERROR_BASE + 38)
Public Const ERROR_INTERNET_HTTP_TO_HTTPS_ON_REDIR = (INTERNET_ERROR_BASE + 39)
Public Const ERROR_INTERNET_HTTPS_TO_HTTP_ON_REDIR = (INTERNET_ERROR_BASE + 40)
Public Const ERROR_INTERNET_MIXED_SECURITY = (INTERNET_ERROR_BASE + 41)
Public Const ERROR_INTERNET_CHG_POST_IS_NON_SECURE = (INTERNET_ERROR_BASE + 42)
Public Const ERROR_INTERNET_POST_IS_NON_SECURE = (INTERNET_ERROR_BASE + 43)
Public Const ERROR_INTERNET_CLIENT_AUTH_CERT_NEEDED = (INTERNET_ERROR_BASE + 44)
Public Const ERROR_INTERNET_INVALID_CA = (INTERNET_ERROR_BASE + 45)
Public Const ERROR_INTERNET_CLIENT_AUTH_NOT_SETUP = (INTERNET_ERROR_BASE + 46)
Public Const ERROR_INTERNET_ASYNC_THREAD_FAILED = (INTERNET_ERROR_BASE + 47)
Public Const ERROR_INTERNET_REDIRECT_SCHEME_CHANGE = (INTERNET_ERROR_BASE + 48)
Public Const ERROR_INTERNET_DIALOG_PENDING = (INTERNET_ERROR_BASE + 49)
Public Const ERROR_INTERNET_RETRY_DIALOG = (INTERNET_ERROR_BASE + 50)
Public Const ERROR_INTERNET_HTTPS_HTTP_SUBMIT_REDIR = (INTERNET_ERROR_BASE + 52)
Public Const ERROR_INTERNET_INSERT_CDROM = (INTERNET_ERROR_BASE + 53)

' FTP API errors

Public Const ERROR_FTP_TRANSFER_IN_PROGRESS = (INTERNET_ERROR_BASE + 110)
Public Const ERROR_FTP_DROPPED = (INTERNET_ERROR_BASE + 111)
Public Const ERROR_FTP_NO_PASSIVE_MODE = (INTERNET_ERROR_BASE + 112)

' gopher API errors

Public Const ERROR_GOPHER_PROTOCOL_ERROR = (INTERNET_ERROR_BASE + 130)
Public Const ERROR_GOPHER_NOT_FILE = (INTERNET_ERROR_BASE + 131)
Public Const ERROR_GOPHER_DATA_ERROR = (INTERNET_ERROR_BASE + 132)
Public Const ERROR_GOPHER_END_OF_DATA = (INTERNET_ERROR_BASE + 133)
Public Const ERROR_GOPHER_INVALID_LOCATOR = (INTERNET_ERROR_BASE + 134)
Public Const ERROR_GOPHER_INCORRECT_LOCATOR_TYPE = (INTERNET_ERROR_BASE + 135)
Public Const ERROR_GOPHER_NOT_GOPHER_PLUS = (INTERNET_ERROR_BASE + 136)
Public Const ERROR_GOPHER_ATTRIBUTE_NOT_FOUND = (INTERNET_ERROR_BASE + 137)
Public Const ERROR_GOPHER_UNKNOWN_LOCATOR = (INTERNET_ERROR_BASE + 138)

' HTTP API errors

Public Const ERROR_HTTP_HEADER_NOT_FOUND = (INTERNET_ERROR_BASE + 150)
Public Const ERROR_HTTP_DOWNLEVEL_SERVER = (INTERNET_ERROR_BASE + 151)
Public Const ERROR_HTTP_INVALID_SERVER_RESPONSE = (INTERNET_ERROR_BASE + 152)
Public Const ERROR_HTTP_INVALID_HEADER = (INTERNET_ERROR_BASE + 153)
Public Const ERROR_HTTP_INVALID_QUERY_REQUEST = (INTERNET_ERROR_BASE + 154)
Public Const ERROR_HTTP_HEADER_ALREADY_EXISTS = (INTERNET_ERROR_BASE + 155)
Public Const ERROR_HTTP_REDIRECT_FAILED = (INTERNET_ERROR_BASE + 156)
Public Const ERROR_HTTP_NOT_REDIRECTED = (INTERNET_ERROR_BASE + 160)
Public Const ERROR_HTTP_COOKIE_NEEDS_CONFIRMATION = (INTERNET_ERROR_BASE + 161)
Public Const ERROR_HTTP_COOKIE_DECLINED = (INTERNET_ERROR_BASE + 162)
Public Const ERROR_HTTP_REDIRECT_NEEDS_CONFIRMATION = (INTERNET_ERROR_BASE + 168)

' additional Internet API error codes

Public Const ERROR_INTERNET_SECURITY_CHANNEL_ERROR = (INTERNET_ERROR_BASE + 157)
Public Const ERROR_INTERNET_UNABLE_TO_CACHE_FILE = (INTERNET_ERROR_BASE + 158)
Public Const ERROR_INTERNET_TCPIP_NOT_INSTALLED = (INTERNET_ERROR_BASE + 159)
Public Const ERROR_INTERNET_DISCONNECTED = (INTERNET_ERROR_BASE + 163)
Public Const ERROR_INTERNET_SERVER_UNREACHABLE = (INTERNET_ERROR_BASE + 164)
Public Const ERROR_INTERNET_PROXY_SERVER_UNREACHABLE = (INTERNET_ERROR_BASE + 165)

Public Const ERROR_INTERNET_BAD_AUTO_PROXY_SCRIPT = (INTERNET_ERROR_BASE + 166)
Public Const ERROR_INTERNET_UNABLE_TO_DOWNLOAD_SCRIPT = (INTERNET_ERROR_BASE + 167)
Public Const ERROR_INTERNET_SEC_INVALID_CERT = (INTERNET_ERROR_BASE + 169)
Public Const ERROR_INTERNET_SEC_CERT_REVOKED = (INTERNET_ERROR_BASE + 170)

' InternetAutodial specific errors

Public Const ERROR_INTERNET_FAILED_DUETOSECURITYCHECK = (INTERNET_ERROR_BASE + 171)

Public Const INTERNET_ERROR_LAST = ERROR_INTERNET_FAILED_DUETOSECURITYCHECK

'
' flags common to open functions (not InternetOpen()):
'

Public Const INTERNET_FLAG_RELOAD = &H80000000             ' retrieve the original item

'
' flags for InternetOpenUrl():
'

Public Const INTERNET_FLAG_RAW_DATA = &H40000000           ' FTP/gopher find: receive the item as raw (structured) data
Public Const INTERNET_FLAG_EXISTING_CONNECT = &H20000000   ' FTP: use existing InternetConnect handle for server if possible

'
' flags for InternetOpen():
'

Public Const INTERNET_FLAG_ASYNC = &H10000000              ' this request is asynchronous (where supported)

'
' protocol-specific flags:
'

Public Const INTERNET_FLAG_PASSIVE = &H8000000             ' used for FTP connections

'
' additional cache flags
'

Public Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000      ' don't write this item to the cache
Public Const INTERNET_FLAG_DONT_CACHE = INTERNET_FLAG_NO_CACHE_WRITE
Public Const INTERNET_FLAG_MAKE_PERSISTENT = &H2000000     ' make this item persistent in cache
Public Const INTERNET_FLAG_FROM_CACHE = &H1000000          ' use offline semantics
Public Const INTERNET_FLAG_OFFLINE = INTERNET_FLAG_FROM_CACHE

'
' additional flags
'

Public Const INTERNET_FLAG_SECURE = &H800000               ' use PCT/SSL if applicable (HTTP)
Public Const INTERNET_FLAG_KEEP_CONNECTION = &H400000      ' use keep-alive semantics
Public Const INTERNET_FLAG_NO_AUTO_REDIRECT = &H200000     ' don't handle redirections automatically
Public Const INTERNET_FLAG_READ_PREFETCH = &H100000        ' do background read prefetch
Public Const INTERNET_FLAG_NO_COOKIES = &H80000            ' no automatic cookie handling
Public Const INTERNET_FLAG_NO_AUTH = &H40000               ' no automatic authentication handling
Public Const INTERNET_FLAG_CACHE_IF_NET_FAIL = &H10000     ' return cache file if net request fails

'
' Security Ignore Flags, Allow HttpOpenRequest to overide
'  Secure Channel (SSL/PCT) failures of the following types.
'

Public Const INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTP = &H8000       ' ex: https:// to http://
Public Const INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTPS = &H4000      ' ex: http:// to https://
Public Const INTERNET_FLAG_IGNORE_CERT_DATE_INVALID = &H2000      ' expired X509 Cert.
Public Const INTERNET_FLAG_IGNORE_CERT_CN_INVALID = &H1000        ' bad common name in X509 Cert.

'
' more caching flags
'

Public Const INTERNET_FLAG_RESYNCHRONIZE = &H800           ' asking wininet to update an item if it is newer
Public Const INTERNET_FLAG_HYPERLINK = &H400               ' asking wininet to do hyperlinking semantic which works right for scripts
Public Const INTERNET_FLAG_NO_UI = &H200                   ' no cookie popup
Public Const INTERNET_FLAG_PRAGMA_NOCACHE = &H100          ' asking wininet to add "pragma: no-cache"
Public Const INTERNET_FLAG_CACHE_ASYNC = &H80              ' ok to perform lazy cache-write
Public Const INTERNET_FLAG_FORMS_SUBMIT = &H40             ' this is a forms submit
Public Const INTERNET_FLAG_NEED_FILE = &H10                ' need a file for this request
Public Const INTERNET_FLAG_MUST_CACHE_REQUEST = INTERNET_FLAG_NEED_FILE

'
' flags for FTP
'

Public Const INTERNET_FLAG_TRANSFER_ASCII = FTP_TRANSFER_TYPE_ASCII       ' = &H00000001
Public Const INTERNET_FLAG_TRANSFER_BINARY = FTP_TRANSFER_TYPE_BINARY     ' = &H00000002

'
' flags field masks
'

Public Const SECURITY_INTERNET_MASK = INTERNET_FLAG_IGNORE_CERT_CN_INVALID Or _
                                 INTERNET_FLAG_IGNORE_CERT_DATE_INVALID Or _
                                 INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTPS Or _
                                 INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTP

Public Const INTERNET_FLAGS_MASK = INTERNET_FLAG_RELOAD Or _
                                 INTERNET_FLAG_RAW_DATA Or _
                                 INTERNET_FLAG_EXISTING_CONNECT Or _
                                 INTERNET_FLAG_ASYNC Or _
                                 INTERNET_FLAG_PASSIVE Or _
                                 INTERNET_FLAG_NO_CACHE_WRITE Or _
                                 INTERNET_FLAG_MAKE_PERSISTENT Or _
                                 INTERNET_FLAG_FROM_CACHE Or _
                                 INTERNET_FLAG_SECURE Or _
                                 INTERNET_FLAG_KEEP_CONNECTION Or _
                                 INTERNET_FLAG_NO_AUTO_REDIRECT Or _
                                 INTERNET_FLAG_READ_PREFETCH Or _
                                 INTERNET_FLAG_NO_COOKIES Or _
                                 INTERNET_FLAG_NO_AUTH Or _
                                 INTERNET_FLAG_CACHE_IF_NET_FAIL Or _
                                 SECURITY_INTERNET_MASK Or _
                                 INTERNET_FLAG_RESYNCHRONIZE Or _
                                 INTERNET_FLAG_HYPERLINK Or _
                                 INTERNET_FLAG_NO_UI Or _
                                 INTERNET_FLAG_PRAGMA_NOCACHE Or _
                                 INTERNET_FLAG_CACHE_ASYNC Or _
                                 INTERNET_FLAG_FORMS_SUBMIT Or _
                                 INTERNET_FLAG_NEED_FILE Or _
                                 INTERNET_FLAG_TRANSFER_BINARY Or _
                                 INTERNET_FLAG_TRANSFER_ASCII
                                

Public Const INTERNET_ERROR_MASK_INSERT_CDROM = &H1

Public Const INTERNET_OPTIONS_MASK = (Not INTERNET_FLAGS_MASK)

'
' common per-API flags (new APIs)
'

Public Const WININET_API_FLAG_ASYNC = &H1                  ' force async operation
Public Const WININET_API_FLAG_SYNC = &H4                   ' force sync operation
Public Const WININET_API_FLAG_USE_CONTEXT = &H8            ' use value supplied in dwContext (even if 0)

'
' INTERNET_NO_CALLBACK - if this value is presented as the dwContext parameter
' then no call-backs will be made for that API
'

Public Const INTERNET_NO_CALLBACK = 0

Public Declare Function InternetDial Lib "wininet.dll" _
(ByVal hWnd As Long, ByVal sConnectoid As String, _
 ByVal dwFlags As Long, lpdwConnection As Long, ByVal dwReserved As Long) As Long

Public Declare Function InternetHangUp Lib "wininet.dll" _
(ByVal dwConnection As Long, ByVal dwReserved As Long) As Long

Public Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" _
(ByVal hInternetSession As Long, ByVal sURL As String, _
ByVal sHeaders As String, ByVal lHeadersLength As Long, _
ByVal lFlags As Long, ByVal lContext As Long) As Long

Public Declare Function InternetAutodial Lib "wininet.dll" _
(ByVal dwFlags As Long, ByVal dwReserved As Long) As Long

Public Const INTERNET_AUTODIAL_FORCE_ONLINE = 1
Public Const INTERNET_AUTODIAL_FORCE_UNATTENDED = 2
Public Const INTERNET_DIAL_UNATTENDED = &H8000


Public Declare Function InternetAutodialHangup Lib "wininet.dll" _
(ByVal dwReserved As Long) As Long


Public Declare Function InternetAttemptConnect Lib "WinInet" (ByVal dwReserved As Long) As Long
Public Declare Function InternetGetConnectedStateEx Lib "wininet.dll" Alias "InternetGetConnectedStateExA" (lpdwFlags As Long, lpszConnectionName As Long, dwNameLen As Long, ByVal dwReserved As Long) As Long


Public Const INTERNET_CONNECTION_LAN As Long = &H2
Public Const INTERNET_CONNECTION_MODEM As Long = &H1

Public Declare Function InternetGetConnectedState _
        Lib "wininet.dll" (ByRef lpSFlags As Long, _
        ByVal dwReserved As Long) As Long

Public Const FLAG_ICC_FORCE_CONNECTION = &H1
Public Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Long

Declare Function WNetAddConnection2 Lib "mpr.dll" Alias _
      "WNetAddConnection2A" (lpNetResource As NETRESOURCE, _
      ByVal lpPassword As String, ByVal lpUserName As String, _
      ByVal dwFlags As Long) As Long



      Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias _
      "WNetCancelConnection2A" (ByVal lpName As String, _
      ByVal dwFlags As Long, ByVal fForce As Long) As Long



      Type NETRESOURCE
        dwScope As Long
        dwType As Long
        dwDisplayType As Long
        dwUsage As Long
        lpLocalName As String
        lpRemoteName As String
        lpComment As String
        lpProvider As String
      End Type

      Type NETRESOURCE_1
        dwScope As Long
        dwType As Long
        dwDisplayType As Long
        dwUsage As Long
        lpLocalName As Long
        lpRemoteName As Long
        lpComment As Long
        lpProvider As Long
      End Type

      Public Const CONNECT_UPDATE_PROFILE = &H1
      Public Const RESOURCETYPE_DISK = &H1
      Public Const RESOURCETYPE_PRINT = &H2
      Public Const RESOURCETYPE_ANY = &H0
      Public Const RESOURCE_CONNECTED = &H1
      Public Const RESOURCE_REMEMBERED = &H3
      Public Const RESOURCE_GLOBALNET = &H2
      Public Const RESOURCEDISPLAYTYPE_DOMAIN = &H1
      Public Const RESOURCEDISPLAYTYPE_GENERIC = &H0
      Public Const RESOURCEDISPLAYTYPE_SERVER = &H2
      Public Const RESOURCEDISPLAYTYPE_SHARE = &H3
      Public Const RESOURCEUSAGE_CONNECTABLE = &H1
      Public Const RESOURCEUSAGE_CONTAINER = &H2
      
      ' Error Constants:
      Public Const ERROR_ACCESS_DENIED = 5&
      Public Const ERROR_ALREADY_ASSIGNED = 85&
      Public Const ERROR_BAD_DEV_TYPE = 66&
      Public Const ERROR_BAD_DEVICE = 1200&
      Public Const ERROR_BAD_NET_NAME = 67&
      Public Const ERROR_BAD_PROFILE = 1206&
      Public Const ERROR_BAD_PROVIDER = 1204&
      Public Const ERROR_BUSY = 170&
      Public Const ERROR_CANCELLED = 1223&
      Public Const ERROR_CANNOT_OPEN_PROFILE = 1205&
      Public Const ERROR_DEVICE_ALREADY_REMEMBERED = 1202&
      Public Const ERROR_EXTENDED_ERROR = 1208&
      Public Const ERROR_INVALID_PASSWORD = 86&
      Public Const ERROR_NO_NET_OR_BAD_PATH = 1203&
      
   Public Declare Function WNetOpenEnum Lib "mpr.dll" Alias _
       "WNetOpenEnumA" ( _
       ByVal dwScope As Long, _
       ByVal dwType As Long, _
       ByVal dwUsage As Long, _
       lpNetResource As Any, _
       lphEnum As Long) As Long

   Public Declare Function WNetEnumResource Lib "mpr.dll" Alias _
       "WNetEnumResourceA" ( _
       ByVal hEnum As Long, _
       lpcCount As Long, _
       ByVal lpBuffer As Long, _
       lpBufferSize As Long) As Long

   Public Declare Function WNetCloseEnum Lib "mpr.dll" ( _
       ByVal hEnum As Long) As Long
      
      
Private Const GMEM_FIXED = &H0
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)


   Public Declare Function GlobalAlloc Lib "kernel32" ( _
      ByVal wFlags As Long, ByVal dwBytes As Long) As Long
   Public Declare Function GlobalFree Lib "kernel32" ( _
      ByVal hMem As Long) As Long

   Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
      hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

   Public Declare Function CopyPointer2String Lib "kernel32" _
      Alias "lstrcpyA" ( _
      ByVal NewString As String, ByVal OldString As Long) As Long

' Winsock API Type defs...
Private Type ICMP_OPTIONS
    Ttl                         As Byte
    Tos                         As Byte
    Flags                       As Byte
    OptionsSize                 As Byte
    OptionsData                 As Long
End Type

Private Type ICMP_ECHO_REPLY
    Address                     As Long
    Status                      As Long
    RoundTripTime               As Long
    DataSize                    As Long
    DataPointer                 As Long
    options                     As ICMP_OPTIONS
    data                        As String * 250
End Type

Private Const WS_VERSION_REQD   As Long = &H101
Private Const MIN_SOCKETS_REQD  As Long = 1
Private Const DATA_SIZE = 32
Private Const MAX_WSAD = 256
Private Const MAX_WSAS = 128
Private Const PING_TIMEOUT = 255
      
Private Type WSADATA
    wVersion                    As Integer
    wHighVersion                As Integer
    szDescription(MAX_WSAD)     As Byte
    szSystemStatus(MAX_WSAS)    As Byte
    wMaxSockets                 As Integer
    wMaxUDPDG                   As Integer
    dwVendorInfo                As Long
End Type

Private Type HostEnt
    hName                       As Long
    hAliases                    As Long
    hAddrType                   As Integer
    hLen                        As Integer
    hAddrList                   As Long
End Type

Private Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
Private Declare Function WSAStartup Lib "wsock32.dll" _
        (ByVal wVersionRequired As Long, lpWSAData As WSADATA) As Long
Private Declare Function WSACleanup Lib "wsock32.dll" () As Long
Private Declare Function GetHostName Lib "wsock32.dll" _
        Alias "gethostname" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Private Declare Function gethostbyname Lib "wsock32.dll" (ByVal szHost As String) As Long
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Private Declare Function IcmpSendEcho Lib "icmp.dll" _
        (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, _
        ByVal RequestData As String, ByVal RequestSize As Long, _
        ByVal RequestOptions As Long, ReplyBuffer As ICMP_ECHO_REPLY, _
        ByVal ReplySize As Long, ByVal Timeout As Long) As Long
      
              ' No more data is available.
      Const ERROR_NO_MORE_ITEMS = 259
      
      ' The data area passed to a system call is too small.
      Const ERROR_INSUFFICIENT_BUFFER = 122

      Public Declare Function InternetSetCookie Lib "wininet.dll" _
       Alias "InternetSetCookieA" _
       (ByVal lpszUrlName As String, _
       ByVal lpszCookieName As String, _
       ByVal lpszCookieData As String) As Boolean

      Public Declare Function InternetGetCookie Lib "wininet.dll" _
       Alias "InternetGetCookieA" _
       (ByVal lpszUrlName As String, _
       ByVal lpszCookieName As String, _
       ByVal lpszCookieData As String, _
       lpdwSize As Long) As Boolean
       
    Public Const HTTP_STATUS_OK = 200
    Public Const HTTP_STATUS_CREATED = 201
    Public Const HTTP_STATUS_ACCEPTED = 202
    Public Const HTTP_STATUS_PARTIAL = 203
    Public Const HTTP_STATUS_NO_CONTENT = 204
    Public Const HTTP_STATUS_RESET_CONTENT = 205
    Public Const HTTP_STATUS_PARTIAL_CONTENT = 206
    Public Const HTTP_STATUS_AMBIGUOUS = 300
    Public Const HTTP_STATUS_MOVED = 301
    Public Const HTTP_STATUS_REDIRECT = 302
    Public Const HTTP_STATUS_REDIRECT_METHOD = 303
    Public Const HTTP_STATUS_NOT_MODIFIED = 304
    Public Const HTTP_STATUS_USE_PROXY = 305
    Public Const HTTP_STATUS_REDIRECT_KEEP_VERB = 307
    Public Const HTTP_STATUS_BAD_REQUEST = 400
    Public Const HTTP_STATUS_DENIED = 401
    Public Const HTTP_STATUS_PAYMENT_REQ = 402
    Public Const HTTP_STATUS_FORBIDDEN = 403
    Public Const HTTP_STATUS_NOT_FOUND = 404
    Public Const HTTP_STATUS_BAD_METHOD = 405
    Public Const HTTP_STATUS_NONE_ACCEPTABLE = 406
    Public Const HTTP_STATUS_PROXY_AUTH_REQ = 407
    Public Const HTTP_STATUS_REQUEST_TIMEOUT = 408
    Public Const HTTP_STATUS_CONFLICT = 409
    Public Const HTTP_STATUS_GONE = 410
    Public Const HTTP_STATUS_LENGTH_REQUIRED = 411
    Public Const HTTP_STATUS_PRECOND_FAILED = 412
    Public Const HTTP_STATUS_REQUEST_TOO_LARGE = 413
    Public Const HTTP_STATUS_URI_TOO_LONG = 414
    Public Const HTTP_STATUS_UNSUPPORTED_MEDIA = 415
    Public Const HTTP_STATUS_SERVER_ERROR = 500
    Public Const HTTP_STATUS_NOT_SUPPORTED = 501
    Public Const HTTP_STATUS_BAD_GATEWAY = 502
    Public Const HTTP_STATUS_SERVICE_UNAVAIL = 503
    Public Const HTTP_STATUS_GATEWAY_TIMEOUT = 504
    Public Const HTTP_STATUS_VERSION_NOT_SUP = 505
       
       
Public Function PSINET_GetHttpCodeMessage$(ByVal lCode&)
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2005
'
'****************************************************************************
'
'                     NAME: Function PSINET_GetHttpCodeMessage
'
'                     iCode%             - HTTP Code
'
'                          ) As String
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    07 October 2005   First created for PivotalStock
'
'                  PURPOSE: Returns the message associated with the HTTP
'                           request code
'
'****************************************************************************
'
'
Dim sReturn$

    '
    ' Initialise error vector
    '
    On Error Resume Next
    Select Case lCode
        Case HTTP_STATUS_OK
            sReturn = "The request completed successfully."
        Case HTTP_STATUS_CREATED
            sReturn = "The request has been fulfilled and resulted in the creation of a new resource."
        Case HTTP_STATUS_ACCEPTED
            sReturn = "The request has been accepted for processing , but the processing has not been completed."
        Case HTTP_STATUS_PARTIAL
            sReturn = "The returned meta information in the entity-header is not the definitive set available from the origin server."
        Case HTTP_STATUS_NO_CONTENT
            sReturn = "The server has fulfilled the request, but there is no new information to send back."
        Case HTTP_STATUS_RESET_CONTENT
            sReturn = "The request has been completed, and the client program should reset the document view that caused the request to be sent to allow the user to easily initiate another input action."
        Case HTTP_STATUS_PARTIAL_CONTENT
            sReturn = "The server has fulfilled the partial GET request for the resource."
        Case HTTP_STATUS_AMBIGUOUS
            sReturn = "The server couldn't decide what to return."
        Case HTTP_STATUS_MOVED
            sReturn = "The requested resource has been assigned to a new permanent URI, and any future references to this resource should be done using one of the returned URIs."
        Case HTTP_STATUS_REDIRECT
            sReturn = "The requested resource resides temporarily under a different URI."
        Case HTTP_STATUS_REDIRECT_METHOD
            sReturn = "The response to the request can be found under a different URI and should be retrieved using a GET method on that resource."
        Case HTTP_STATUS_NOT_MODIFIED
            sReturn = "The requested resource has not been modified."
        Case HTTP_STATUS_USE_PROXY
            sReturn = "The requested resource must be accessed through the proxy given by the location field."
        Case HTTP_STATUS_REDIRECT_KEEP_VERB
            sReturn = "The redirected request keeps the same verb. HTTP/1.1 behavior."
        Case HTTP_STATUS_BAD_REQUEST
            sReturn = "The request could not be processed by the server due to invalid syntax."
        Case HTTP_STATUS_DENIED
            sReturn = "The requested resource requires user authentication."
        Case HTTP_STATUS_PAYMENT_REQ
            sReturn = "Not currently implemented in the HTTP protocol."
        Case HTTP_STATUS_FORBIDDEN
            sReturn = "The server understood the request, but is refusing to fulfill it."
        Case HTTP_STATUS_NOT_FOUND
            sReturn = "The server has not found anything matching the requested URI."
        Case HTTP_STATUS_BAD_METHOD
            sReturn = "The method used is not allowed."
        Case HTTP_STATUS_NONE_ACCEPTABLE
            sReturn = "No responses acceptable to the client were found."
        Case HTTP_STATUS_PROXY_AUTH_REQ
            sReturn = "Proxy authentication required."
        Case HTTP_STATUS_REQUEST_TIMEOUT
            sReturn = "The server timed out waiting for the request."
        Case HTTP_STATUS_CONFLICT
            sReturn = "The request could not be completed due to a conflict with the current state of the resource. The user should resubmit with more information."
        Case HTTP_STATUS_GONE
            sReturn = "The requested resource is no longer available at the server, and no forwarding address is known."
        Case HTTP_STATUS_LENGTH_REQUIRED
            sReturn = "The server refuses to accept the request without a defined content length."
        Case HTTP_STATUS_PRECOND_FAILED
            sReturn = "The precondition given in one or more of the request header fields evaluted to false when it was tested on the server."
        Case HTTP_STATUS_REQUEST_TOO_LARGE
            sReturn = "The server is refusing to process a request because the request entity is larger than the server is willing or able to process."
        Case HTTP_STATUS_URI_TOO_LONG
            sReturn = "The server is refusing to service the request because the request URI is longer than the server is willing to interpret."
        Case HTTP_STATUS_UNSUPPORTED_MEDIA
            sReturn = "The server is refusing to service the request because the entity of the request is in a format not supported by the requested resource for the requested method."
        Case HTTP_STATUS_SERVER_ERROR
            sReturn = "The server encountered an unexpected condition that prevented it from fulfilling the request."
        Case HTTP_STATUS_NOT_SUPPORTED
            sReturn = "The server does not support the functionality required to fulfill the request."
        Case HTTP_STATUS_BAD_GATEWAY
            sReturn = "The server, while acting as a gateway or proxy, received an invalid response from the upstream server it accessed in attempting to fulfill the request."
        Case HTTP_STATUS_SERVICE_UNAVAIL
            sReturn = "The service is temporarily overloaded."
        Case HTTP_STATUS_GATEWAY_TIMEOUT
            sReturn = "The request was timed out waiting for a gateway."
        Case HTTP_STATUS_VERSION_NOT_SUP
            sReturn = "The server does not support, or refuses to support, the HTTP protocol version that was used in the request message."
        Case Else
            sReturn = "Unknown HTTP error code"
    End Select

    '
    ' Return value to caller
    '
    PSINET_GetHttpCodeMessage = Format(lCode) + " " + sReturn

End Function



Public Function PSINET_TranslateErrorCode$(ByVal lErrorCode&)
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2002
'
'****************************************************************************
'
'                     NAME: Function PSINET_TranslateErrorCode
'
'                     lErrorCode&        - Error code returned from DLL
'
'                          ) As String
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    13 May 2002   First created for TeamPlayer
'
'                  PURPOSE: Translates the error code into something
'                           meaningful
'
'****************************************************************************
'
'
Dim sReturn$
Dim lError&, lLength&
Dim sError$

    '
    ' Determine the error text
    '
    On Error Resume Next
    Select Case lErrorCode
        Case 0
            lLength = 1000
            sError = String(lLength, vbNullChar)
            Call InternetGetLastResponseInfo(lError, sError, lLength)
            sReturn = Left$(sError, lLength)
            
        Case ERROR_ACCESS_DENIED: sReturn = "Access denied"
        Case ERROR_ALREADY_ASSIGNED: sReturn = "Drive letter already assigned"
        Case ERROR_BAD_DEV_TYPE: sReturn = "Bad device type specified"
        Case ERROR_BAD_DEVICE: sReturn = "Bad device specified"
        Case ERROR_BAD_NET_NAME: sReturn = "Bad network name specified"
        Case ERROR_BAD_PROFILE: sReturn = "Bad profile specified"
        Case ERROR_BAD_PROVIDER: sReturn = "Bad provider specified"
        Case ERROR_BUSY: sReturn = "System or resource is busy"
        Case ERROR_CANCELLED: sReturn = "Connection request cancelled"
        Case ERROR_CANNOT_OPEN_PROFILE: sReturn = "Cannot open profile"
        Case ERROR_DEVICE_ALREADY_REMEMBERED: sReturn = "Device already remembered"
        Case ERROR_EXTENDED_ERROR: sReturn = "Extended error"
        Case ERROR_INVALID_PASSWORD: sReturn = "Invalid password"
        Case ERROR_NO_NET_OR_BAD_PATH: sReturn = "No network or bad path specified"
        
        Case 12001: sReturn = "No more handles could be generated at this time"
        Case 12002: sReturn = "The request has timed out."
        Case 12003: sReturn = "An extended error was returned from the server."
        Case 12004: sReturn = "An internal error has occurred."
        Case 12005: sReturn = "The URL is invalid."
        Case 12006: sReturn = "The URL scheme could not be recognized, or is not supported."
        Case 12007: sReturn = "The server name could not be resolved."
        Case 12008: sReturn = "The requested protocol could not be located."
        Case 12009: sReturn = "A request to InternetQueryOption or InternetSetOption specified an invalid option value."
        Case 12010: sReturn = "The length of an option supplied to InternetQueryOption or InternetSetOption is incorrect for the type of option specified."
        Case 12011: sReturn = "The request option can not be set, only queried. "
        Case 12012: sReturn = "The Win32 Internet support is being shutdown or unloaded."
        Case 12013: sReturn = "The request to connect and login to an FTP server could not be completed because the supplied user name is incorrect."
        Case 12014: sReturn = "The request to connect and login to an FTP server could not be completed because the supplied password is incorrect. "
        Case 12015: sReturn = "The request to connect to and login to an FTP server failed."
        Case 12016: sReturn = "The requested operation is invalid. "
        Case 12017: sReturn = "The operation was canceled, usually because the handle on which the request was operating was closed before the operation completed."
        Case 12018: sReturn = "The type of handle supplied is incorrect for this operation."
        Case 12019: sReturn = "The requested operation can not be carried out because the handle supplied is not in the correct state."
        Case 12020: sReturn = "The request can not be made via a proxy."
        Case 12021: sReturn = "A required registry value could not be located. "
        Case 12022: sReturn = "A required registry value was located but is an incorrect type or has an invalid value."
        Case 12023: sReturn = "Direct network access cannot be made at this time. "
        Case 12024: sReturn = "An asynchronous request could not be made because a zero context value was supplied."
        Case 12025: sReturn = "An asynchronous request could not be made because a callback function has not been set."
        Case 12026: sReturn = "The required operation could not be completed because one or more requests are pending."
        Case 12027: sReturn = "The format of the request is invalid."
        Case 12028: sReturn = "The requested item could not be located."
        Case 12029: sReturn = "The attempt to connect to the server failed."
        Case 12030: sReturn = "The connection with the server has been terminated."
        Case 12031: sReturn = "The connection with the server has been reset."
        Case 12036: sReturn = "The request failed because the handle already exists."
        Case Else: sReturn = "Error details not available."
    End Select

    '
    ' Get any extended info
    '
    lLength = 1000
    sError = String(lLength, vbNullChar)
    Call InternetGetLastResponseInfo(lError, sError, lLength)
    If lLength > 0 Then sReturn = Left(sError, lLength)

    '
    ' Return value to caller
    '
    PSINET_TranslateErrorCode = sReturn

End Function

Private Function Z_PointerToString$(ByVal lPointer&)
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2002
'
'****************************************************************************
'
'                     NAME: Function Z_PointerToString
'
'                     lPointer&          - Pointer to a string in memory
'
'                          ) As String
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    13 May 2002   First created for TeamPlayer
'
'                  PURPOSE: Returns the string pointed to by lPointer
'
'****************************************************************************
'
'
Dim sReturn$


    '
    ' The values returned in the NETRESOURCE structures are pointers to
    ' ANSI strings so they need to be converted to Visual Basic Strings
    '
    On Error Resume Next
    sReturn = String(255, Chr$(0))
    CopyPointer2String sReturn, lPointer
    sReturn = Left(sReturn, InStr(sReturn, Chr$(0)) - 1)

    '
    ' Return value to caller
    '
    Z_PointerToString = sReturn

End Function

Public Function PSINET_GetConnectedShares$(Optional ByVal bIncludeDrives As Boolean)
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2002
'
'****************************************************************************
'
'                     NAME: Function PSINET_GetConnectedShares
'
'                           bIncludeDrives      - True, include the drive letter in brackets
'
'                          ) As String
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    13 May 2002   First created for TeamPlayer
'
'                  PURPOSE: Returns a comma separated list of connected share
'                           names
'
'****************************************************************************
'
'
Dim lEnum&, lBuff&
Dim stRes As NETRESOURCE_1
Dim lStringBuff&, lCount&
Dim lPointer&, lResource&, lCnt&
Dim sReturn$, sTmp$


    '
    ' Setup the NETRESOURCE input structure
    '
    On Error Resume Next
    stRes.dwUsage = RESOURCEUSAGE_CONTAINER
    stRes.lpRemoteName = 0
    lStringBuff = 1000
    lCount = &HFFFFFFFF

    '
    ' Open a Net enumeration operation handle: lEnum
    '
    lResource = WNetOpenEnum(RESOURCE_CONNECTED, RESOURCETYPE_ANY, 0, stRes, lEnum)
    If lResource = 0 Then
       
       '
       ' Create a buffer large enough for the results, 1000 bytes should be sufficient
       '
       lBuff = GlobalAlloc(GPTR, lStringBuff)
       
       '
       ' Call the enumeration function
       '
       lResource = WNetEnumResource(lEnum, lCount, lBuff, lStringBuff)
       If lResource = 0 Then
          lPointer = lBuff
          
          '
          ' WNetEnumResource fills the buffer with an array of NETRESOURCE
          ' structures. Walk through the list and print each local and remote name
          '
          For lCnt = 1 To lCount
             CopyMemory stRes, ByVal lPointer, LenB(stRes)
             lPointer = lPointer + LenB(stRes)
             sReturn = sReturn + IIf(lCnt = 1, "", ",") + Z_PointerToString(stRes.lpRemoteName)
             If bIncludeDrives Then
                sTmp = Z_PointerToString(stRes.lpLocalName)
                If sTmp <> "" Then sReturn = sReturn + " (" + sTmp + ")"
            End If
          Next lCnt
       End If
       
       '
       ' Free the memory and close the enumeration
       '
       If lBuff <> 0 Then Call GlobalFree(lBuff)
       Call WNetCloseEnum(lEnum)
    End If

    '
    ' Return value to caller
    '
    PSINET_GetConnectedShares = sReturn

End Function



      
      
      
      
Public Function PSINET_CloseNetworkShare(ByVal sShare$, Optional ByVal bReconnect As Boolean) As Boolean
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2002
'
'****************************************************************************
'
'                     NAME: Sub PSINET_CloseNetworkShare
'
'                     sShare$            - Share to disconnect
'                     bReconnect         - True, reconnect share after re-boot
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    13 May 2002   First created for TeamPlayer
'
'                  PURPOSE: Closes the connection defined by sResourse
'                           This can point to the share or the sahre drive
'                           name
'
'****************************************************************************
'
'
Dim bReturn As Boolean
Dim lErrInfo&


    '
    ' You may specify either the lpRemoteName or lpLocalName
    ' "\\ServerName\ShareName" or "Z:"
    '
    On Error Resume Next
    lErrInfo = WNetCancelConnection2(sShare, IIf(bReconnect, CONNECT_UPDATE_PROFILE, 0), False)
    If lErrInfo = NO_ERROR Then
        bReturn = True
    Else
        Err.Raise vbObjectError + ERROR_OFFSET, ERROR_SOURCE, PSINET_TranslateErrorCode(Err.LastDllError)
    End If

    PSINET_CloseNetworkShare = bReturn

End Function

Public Function PSINET_OpenNetworkShare(ByVal sServer$, ByVal sUsername$, ByVal sPassword$, ByVal sDrive$, Optional ByVal bPersistent As Boolean) As Boolean
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2002
'
'****************************************************************************
'
'                     NAME: Function PSINET_OpenNetworkShare
'
'                     sServer$           - Name of the share to connect to
'                     sUsername$         - Username to use (NULL means use current)
'                     sPassword$         - Password to use (NULL means use current)
'                     sDrive$            - Drive letter to connect to (NULL means none)
'                     bPersistent        - True, connection is re-created after re-boot
'
'                          ) As Boolean
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    13 May 2002   First created for TeamPlayer
'
'                  PURPOSE: Connects an intranet resource to the local
'                           connection list
'
'****************************************************************************
'
'
Dim bReturn As Boolean
Dim stNetResource As NETRESOURCE
Dim lErrInfo&


    '
    ' Initialise error vector
    '
    On Error Resume Next
    stNetResource.dwScope = RESOURCE_GLOBALNET
    stNetResource.dwType = RESOURCETYPE_DISK
    stNetResource.dwDisplayType = RESOURCEDISPLAYTYPE_SHARE
    stNetResource.dwUsage = RESOURCEUSAGE_CONNECTABLE
    If sDrive = "" Then sDrive = vbNullString
    stNetResource.lpLocalName = sDrive
    stNetResource.lpRemoteName = sServer

    '
    ' If the UserName and Password arguments are NULL, the user context
    ' for the process provides the default user name.
    '
    If sPassword = "" Then sPassword = vbNullString
    If sUsername = "" Then sUsername = vbNullString
    lErrInfo = WNetAddConnection2(stNetResource, sPassword, sUsername, IIf(bPersistent, CONNECT_UPDATE_PROFILE, 0))
    If lErrInfo = NO_ERROR Then
        bReturn = True
    Else
        Err.Raise vbObjectError + ERROR_OFFSET, ERROR_SOURCE, PSINET_TranslateErrorCode(Err.LastDllError)
    End If

    '
    ' Return value to caller
    '
    PSINET_OpenNetworkShare = bReturn

End Function

Public Function PSINET_GetHTTPFile(ByVal sURL$, sValue$, Optional ByVal sTitle$, Optional ByVal lFlags& = (INTERNET_FLAG_RELOAD + INTERNET_FLAG_PRAGMA_NOCACHE), Optional vCookies, Optional vHeaders, Optional ByVal lUseSession&, Optional sProxyName$ = "", Optional sProxyByPass$ = "", Optional lConnectionTimeout& = 30000, Optional lReadTimeout& = 30000, Optional iRetries = 0) As Boolean
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2002
'
'****************************************************************************
'
'                     NAME: Function PSINET_GetHTTPFile
'
'                     sURL$              - The URL of the file to get
'                     sValue$            - Contents of the file
'                     sTitle$            - Optional title
'                     vHeaders$          - Additional headers (array)
'                     lUseSession$       - Existing connection handle
'                     sProxyName         - Name of proxy server to use
'                     sProxyByPass       - List of addresses for the proxy to bypass
'                     iRetries           - Number of retries if the content is empty
'
'                          ) As Boolean
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    25 April 2002   First created for WebSneak
'
'                  PURPOSE: Returns the contents of a remote HTTP file in
'                           sValue
'                           If successful, then returns true
'                           The sURL should be expressed as a full spec HTTP
'                           URL
'
'****************************************************************************
'
'
Dim bReturn As Boolean
Dim iRetry%

    On Error Resume Next
    While iRetry <= iRetries And Not bReturn
        Err.Clear
        bReturn = Z_GetHTTPFile(sURL, sValue, sTitle, lFlags, vCookies, vHeaders, lUseSession, sProxyName, sProxyByPass, lConnectionTimeout, lReadTimeout)
        iRetry = iRetry + 1
    Wend

    If Err.Description <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + ERROR_OFFSET, ERROR_SOURCE, Err.Description
    End If
    PSINET_GetHTTPFile = bReturn
    
End Function

Private Function Z_GetHTTPFile(ByVal sURL$, sValue$, Optional ByVal sTitle$, Optional ByVal lFlags& = (INTERNET_FLAG_RELOAD + INTERNET_FLAG_PRAGMA_NOCACHE), Optional vCookies, Optional vHeaders, Optional ByVal lUseSession&, Optional sProxyName$ = "", Optional sProxyByPass$ = "", Optional lConnectionTimeout& = 30000, Optional lReadTimeout& = 30000) As Boolean
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2002
'
'****************************************************************************
'
'                     NAME: Function Z_GetHTTPFile
'
'                     sURL$              - The URL of the file to get
'                     sValue$            - Contents of the file
'                     sTitle$            - Optional title
'                     vHeaders$          - Additional headers (array)
'                     lUseSession$       - Existing connection handle
'                     sProxyName         - Name of proxy server to use
'                     sProxyByPass       - List of addresses for the proxy to bypass
'
'                          ) As Boolean
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    25 April 2002   First created for WebSneak
'
'                  PURPOSE: Returns the contents of a remote HTTP file in
'                           sValue
'                           If successful, then returns true
'                           The sURL should be expressed as a full spec HTTP
'                           URL
'
'****************************************************************************
'
'
Const CONNECTION_TIMEOUT = 5000&
Const RECEIVE_TIMEOUT = 30000&
Const SEND_TIMEOUT = 30000&

Dim bReturn As Boolean
Dim sBuffer As String * 32768
Dim lFile&, lLength&, lSession&, lTmp&
Dim bFinished As Boolean
Dim iCnt%
Dim dTimeOut As Date
Dim sError$

    '
    ' Set the cookies
    '
    On Error Resume Next
    Err.Clear
    sURL = Trim(sURL)
    sValue = ""
    If Not IsMissing(vCookies) Then
        For iCnt = 0 To UBound(vCookies)
            If Err <> 0 Then Exit For
            Call InternetSetCookie(sURL, PSGEN_GetItem(1, "=", vCookies(iCnt)), PSGEN_GetItem(2, "=", vCookies(iCnt)))
        Next iCnt
    End If

    '
    ' Connect to the session
    '
    If sTitle = "" Then sTitle = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    If lUseSession <> 0 Then
        lSession = lUseSession
    Else
        If sProxyName = "" Then
            lSession = InternetOpen(sTitle, INTERNET_OPEN_TYPE_DIRECT, "", "", 0)
        Else
            lSession = InternetOpen(sTitle, INTERNET_OPEN_TYPE_PROXY, sProxyName, sProxyByPass, 0)
        End If
    End If
    
    '
    ' Set the timeouts
    '
    Call InternetSetOption(lSession, INTERNET_OPTION_CONNECT_TIMEOUT, lConnectionTimeout, 4)
    Call InternetSetOption(lSession, INTERNET_OPTION_RECEIVE_TIMEOUT, lReadTimeout, 4)
    Call InternetSetOption(lSession, INTERNET_OPTION_SEND_TIMEOUT, SEND_TIMEOUT, 4)
    
    '
    ' Ignore Certificat errors
    '
    Call InternetSetOption(lSession, INTERNET_OPTION_SECURITY_FLAGS, SECURITY_FLAG_IGNORE_CERT_CN_INVALID Or SECURITY_FLAG_IGNORE_CERT_DATE_INVALID Or SECURITY_FLAG_IGNORE_UNKNOWN_CA, 4)
    
    '
    ' Get the file
    '
    If Not IsMissing(vHeaders) Then
        vHeaders = Join(vHeaders, vbNullChar)
        lFile = InternetOpenUrl(lSession, sURL, vHeaders, Len(vHeaders), lFlags, 0)
    Else
        lFile = InternetOpenUrl(lSession, sURL, vbNullString, 0, lFlags, 0)
    End If
    If lFile = 0 Then
        sError = PSINET_TranslateErrorCode(Err.LastDllError)
    Else
        '
        ' Check that we got the file we wanted
        '
        lLength = Len(sBuffer)
        Call HttpQueryInfo(lFile, HTTP_QUERY_STATUS_CODE, ByVal sBuffer, lLength, lTmp)
        If Left(sBuffer, 1) <> "2" Then
            sError = "STATUS:" + PSINET_GetHttpCodeMessage(Val(Left(sBuffer, lLength)))
        Else
            While Not bFinished
                sBuffer = vbNullString
                If InternetReadFile(lFile, sBuffer, Len(sBuffer), lLength) = 0 Then
                    sError = PSINET_TranslateErrorCode(Err.LastDllError)
                    bFinished = True
                Else
                    sValue = sValue & Left$(sBuffer, lLength)
                    bFinished = (lLength = 0)
                End If
            Wend
            bReturn = True
        End If
    End If
        
    '
    ' Close the handles
    '
    If lFile <> 0 Then Call InternetCloseHandle(lFile)
    If lUseSession = 0 And lSession <> 0 Then Call InternetCloseHandle(lSession)
    DoEvents

    '
    ' Return value to caller
    '
    If sError <> "" Then
        On Error GoTo 0
        Err.Raise vbObjectError + ERROR_OFFSET, ERROR_SOURCE, sError
    End If
    Z_GetHTTPFile = bReturn

End Function





Public Function PSINET_PostHTTPFile(ByVal sURL$, ByVal sContent$, ByVal sBoundary$, sValue$, Optional ByVal lFlags& = (INTERNET_FLAG_KEEP_CONNECTION + INTERNET_FLAG_RELOAD), Optional sProxyName$ = "", Optional sProxyByPass$ = "") As Boolean
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2002
'
'****************************************************************************
'
'                     NAME: Function PSINET_PostHTTPFile
'
'                     sURL$              - The URL of the file to get
'                     sContent$          - Contents of the file to send
'                     sBoundary$         - Boundary between data elements
'                     sValue$            - Returned response
'                     sProxyName         - Name of proxy server to use
'                     sProxyByPass       - List of addresses for the proxy to bypass
'
'                          ) As Boolean
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    25 April 2002   First created for WebSneak
'
'                  PURPOSE: Sends the value as a POST
'                           If successful, then returns true
'                           The sURL should be expressed as a full spec HTTP
'                           URL
'
'****************************************************************************
'
'
Dim bReturn As Boolean
Dim sBuffer As String * 65000
Dim sServer$, sService$, sHeader$
Dim lFile&, lLength&, lSession&, lConnection&, lTmp&
Dim bFinished As Boolean


    '
    ' Connect to the session
    '
    On Error Resume Next
    Err.Clear
    sValue = ""
    If sProxyName = "" Then
        lSession = InternetOpen(App.Title, INTERNET_OPEN_TYPE_DIRECT, "", "", 0)
    Else
        lSession = InternetOpen(App.Title, INTERNET_OPEN_TYPE_PROXY, sProxyName, sProxyByPass, 0)
    End If
    
    '
    ' Connect to the server
    '
    sServer = PSGEN_GetItem(3, "/", sURL)
    lConnection = InternetConnect(lSession, sServer, INTERNET_INVALID_PORT_NUMBER, vbNullString, vbNullString, INTERNET_SERVICE_HTTP, 0, 0)
    If lConnection = 0 Then
        Err.Raise vbObjectError + ERROR_OFFSET, ERROR_SOURCE, PSINET_TranslateErrorCode(Err.LastDllError)
    Else
    
        '
        ' Open a connection
        '
        sService = PSGEN_GetItem(2, sServer, sURL)
        lFile = HttpOpenRequest(lConnection, "POST", sService, vbNullString, vbNullString, 0, lFlags, 0)
    If lFile = 0 Then
        Err.Raise vbObjectError + ERROR_OFFSET, ERROR_SOURCE, PSINET_TranslateErrorCode(Err.LastDllError)
    Else
        
            '
            ' Send the data
            '
            sHeader = "Content-Type: multipart/form-data; boundary=" + sBoundary & vbCrLf
            Call HttpAddRequestHeaders(lFile, sHeader, Len(sHeader), HTTP_ADDREQ_FLAG_REPLACE + HTTP_ADDREQ_FLAG_ADD)
            If HttpSendRequest(lFile, vbNullString, 0, sContent, Len(sContent)) = 0 Then
                Err.Raise vbObjectError + ERROR_OFFSET, ERROR_SOURCE, PSINET_TranslateErrorCode(Err.LastDllError)
            Else
            
                '
                ' Check that we got a correct header
                '
                lLength = Len(sBuffer)
                Call HttpQueryInfo(lFile, HTTP_QUERY_STATUS_CODE, ByVal sBuffer, lLength, lTmp)
                If Left(sBuffer, 1) <> "2" Then
                    Err.Raise vbObjectError + ERROR_OFFSET, ERROR_SOURCE, "STATUS:" + PSINET_GetHttpCodeMessage(Val(Left(sBuffer, lLength)))
                Else
                    '
                    ' Read the response
                    '
        While Not bFinished
            sBuffer = vbNullString
            Call InternetReadFile(lFile, sBuffer, Len(sBuffer), lLength)
                        sValue = sValue & Left(sBuffer, lLength)
            bFinished = (lLength = 0)
        Wend
                    bReturn = True
                End If
            End If
        Call InternetCloseHandle(lFile)
    End If
        Call InternetCloseHandle(lConnection)
    End If
    Call InternetCloseHandle(lSession)
    DoEvents

    '
    ' Return value to caller
    '
    PSINET_PostHTTPFile = bReturn

End Function


Public Function PSINET_PutFtpFile(ByVal sServer$, ByVal sServerUid$, ByVal sServerPwd$, ByVal sSourceFile$, ByVal sDestFolder$, ByVal sDestFile$, ByVal bUseBinary As Boolean, Optional sProxyName$ = "", Optional sProxyByPass$ = "") As Boolean
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2002
'
'****************************************************************************
'
'                     NAME: Function PSINET_PutFtpFile
'
'                     sServer$           - Host name or IP address
'                     sServerUid$        - Server username
'                     sServerUpw$        - Server password
'                     sSourceFile$       - Source filename
'                     sDestFolder$       - Destination directory
'                     sDestFile$         - Destination filename
'                     sProxyName         - Name of proxy server to use
'                     sProxyByPass       - List of addresses for the proxy to bypass
'
'                          ) As Boolean
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    25 April 2002   First created for WebSneak
'
'                  PURPOSE: Sends the file using the FTP service
'
'****************************************************************************
'
'
Dim bReturn As Boolean
Dim lSession&, lConnection&

    '
    ' Connect to the session
    '
    On Error Resume Next
    Err.Clear
    If sProxyName = "" Then
        lSession = InternetOpen(App.Title, INTERNET_OPEN_TYPE_DIRECT, "", "", 0)
    Else
        lSession = InternetOpen(App.Title, INTERNET_OPEN_TYPE_PROXY, sProxyName, sProxyByPass, 0)
    End If
    
    '
    ' Connect to the server
    '
    lConnection = InternetConnect(lSession, sServer, INTERNET_INVALID_PORT_NUMBER, sServerUid, sServerPwd, INTERNET_SERVICE_FTP, 0, 0)
    If lConnection = 0 Then
        Err.Raise vbObjectError + ERROR_OFFSET, ERROR_SOURCE, PSINET_TranslateErrorCode(Err.LastDllError)
    Else
    
        '
        ' Change directory on the server
        '
        If sDestFolder <> "" Then
            If FtpSetCurrentDirectory(lConnection, sDestFolder) = 0 Then Err.Raise vbObjectError + ERROR_OFFSET, ERROR_SOURCE, "Cannot change directory to " + sDestFolder + vbCrLf + PSINET_TranslateErrorCode(Err.LastDllError)
        End If
        If Err = 0 Then
        
            '
            ' Send the file
            '
            If FtpPutFile(lConnection, sSourceFile, sDestFile, IIf(bUseBinary, INTERNET_FLAG_TRANSFER_BINARY, INTERNET_FLAG_TRANSFER_ASCII), 0) = 0 Then
                Err.Raise vbObjectError + ERROR_OFFSET, ERROR_SOURCE, PSINET_TranslateErrorCode(Err.LastDllError)
            Else
                bReturn = True
            End If
            Call InternetCloseHandle(lConnection)
        End If
    End If
    Call InternetCloseHandle(lSession)
    DoEvents

    '
    ' Return value to caller
    '
    PSINET_PutFtpFile = bReturn

End Function


Public Function PSINET_Ping(sAddress$, Optional ByVal lTimeout& = PING_TIMEOUT, Optional sRoundTripTime$ = "", Optional sDataSize$ = "", Optional sIPAddress$ = "", Optional bDataMatch As Boolean = False) As Boolean
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2002
'
'****************************************************************************
'
'                     NAME: Function PSINET_Ping
'
'                           sAddress$        - Host to ping
'                           lTimeout&        - Timeout for operation (msec)
'                           sRoundTripTime$  - Returned time for ping operation (msec)
'                           sDataSize$       - Returned size of data exchanged
'                           sIPAddress$      - Returned resolved IP address
'                           bDataMatch       - Returns true if data matched sent
'
'                          ) As Boolean
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    25 April 2002   First created for WebSneak
'
'                  PURPOSE: Tries to ping the specified host and returns True if OK
'                           Optionally returns time taken for the ping, the data
'                           exchanged and wether the sent matched the recieved.
'
'****************************************************************************
'
'
Dim stECHO As ICMP_ECHO_REPLY
Dim iPtr%
Dim sTmp$
Dim hPort&
Dim lAddress&
Dim abAddr(3) As Byte


    '
    ' Assume failure
    '
    On Error Resume Next
    PSINET_Ping = False

    '
    ' If passed a name, get the IP address
    '
    If Not Z_IsDottedQuad(sAddress) Then sAddress = PSINET_LookupIPAddress(sAddress)
    If sAddress = "" Then Exit Function

    '
    ' Init the sockets api
    '
    If Z_SocketsInitialize Then

        '
        ' Build string of random characters
        '
        For iPtr = 1 To DATA_SIZE
            sTmp = sTmp & Chr$(Rnd() * 254 + 1)
        Next

        '
        ' Ping an ip address, passing the address and the ECHO structure
        '
        lAddress = Z_AddressStringToLong(sAddress)
        hPort = IcmpCreateFile()
        IcmpSendEcho hPort, lAddress, sTmp, Len(sTmp), 0, stECHO, Len(stECHO), lTimeout
        IcmpCloseHandle hPort

        '
        ' Get the results from the ECHO structure
        '
        sRoundTripTime = stECHO.RoundTripTime
        CopyMemory abAddr(0), stECHO.Address, 4
        sIPAddress = CStr(abAddr(0)) & "." & CStr(abAddr(1)) & "." & CStr(abAddr(2)) & "." & CStr(abAddr(3))
        sDataSize = stECHO.DataSize & " bytes"

        iPtr = InStr(stECHO.data, vbNullChar)
        If iPtr > 1 Then bDataMatch = (Left$(stECHO.data, iPtr - 1) = sTmp)
        If stECHO.Status = 0 And stECHO.Address = lAddress Then PSINET_Ping = True

        '
        ' Clean up the sockets connection
        '
        WSACleanup
    End If

End Function



Private Function Z_IsDottedQuad(ByVal sIPString$) As Boolean
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2002
'
'****************************************************************************
'
'                     NAME: Function Z_IsDottedQuad
'
'                           sIPString$   - Dot format of hostname
'
'                          ) As Boolean
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    25 April 2002   First created for WebSneak
'
'                  PURPOSE: Verifies that the host string is in the correct
'                           xxx.xxx.xxx.xxx format
'
'****************************************************************************
'
'
Dim sSplit$()
Dim iCtr%

    '
    ' Split at the "."
    '
    sSplit = Split(sIPString, ".")

    '
    ' should be 4 elements
    '
    If UBound(sSplit) <> 3 Then Exit Function

    '
    ' Check each element
    '
    For iCtr = 0 To 3
        
        '
        ' Should be numeric
        '
        If Not IsNumeric(sSplit(iCtr)) Then Exit Function

        '
        ' range check
        '
        If iCtr = 0 Then
            If Val(sSplit(iCtr)) > 239 Then Exit Function
        Else
            If Val(sSplit(iCtr)) > 255 Then Exit Function
        End If
    Next
    
    Z_IsDottedQuad = True

End Function


Public Function PSINET_LookupIPAddress$(Optional ByVal sHostName$)
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2002
'
'****************************************************************************
'
'                     NAME: Function PSINET_LookupIPAddress
'
'                            sHostName$     - Name of the host
'
'                          ) As String
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    25 April 2002   First created for WebSneak
'
'                  PURPOSE: Uses DNS to return the IP address of the host
'
'****************************************************************************
'
'
Dim lpHost&
Dim stHost As HostEnt
Dim dwIPAddr&
Dim abIPAddr() As Byte
Dim iCnt%
Dim sIPAddr$


    '
    ' Init winsock api
    '
    If Not Z_SocketsInitialize() Then
        PSINET_LookupIPAddress = ""
        Exit Function
    End If
    
    '
    ' If no name given, use local host
    '
    If sHostName = "" Then sHostName = PSINET_GetLocalHostName
    sHostName = Trim(sHostName) & vbNullChar
    
    '
    ' Call api
    '
    lpHost = gethostbyname(sHostName)
    If lpHost Then
    
        '
        ' Extract the data...
        '
        CopyMemory stHost, ByVal lpHost, Len(stHost)
        CopyMemory dwIPAddr, ByVal stHost.hAddrList, 4
        ReDim abIPAddr(1 To stHost.hLen)
        CopyMemory abIPAddr(1), ByVal dwIPAddr, stHost.hLen

        '
        ' Convert format
        '
        For iCnt = 1 To stHost.hLen
            sIPAddr = sIPAddr & abIPAddr(iCnt) & "."
        Next

        '
        ' set the return value
        '
        PSINET_LookupIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
    Else
        WSAGetLastError
        PSINET_LookupIPAddress = ""
    End If
    
    '
    ' Close the sockets library
    '
    WSACleanup

End Function


Private Function Z_SocketsInitialize() As Boolean
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2002
'
'****************************************************************************
'
'                     NAME: Function Z_SocketsInitialize
'
'                          ) As Boolean
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    25 April 2002   First created for WebSneak
'
'                  PURPOSE: Opens the sockets library - returns true if OK
'
'****************************************************************************
'
'
Dim stWSAD As WSADATA

    '
    ' Initialize Windows sockets
    '
    Z_SocketsInitialize = False
    If WSAStartup(WS_VERSION_REQD, stWSAD) <> ERROR_SUCCESS Then Exit Function
    If stWSAD.wMaxSockets < MIN_SOCKETS_REQD Then Exit Function

    Z_SocketsInitialize = True

End Function


Private Function Z_AddressStringToLong&(ByVal sIPAddr$)
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2002
'
'****************************************************************************
'
'                     NAME: Function Z_AddressStringToLong
'
'                            sIPAddr$     - IP address in dot format
'
'                          ) As Long
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    25 April 2002   First created for WebSneak
'
'                  PURPOSE: Converts an IP address into a long
'
'****************************************************************************
'
'
Dim sParts$()

    '
    ' Convert an ip address string to a long value
    '
    sParts = Split(sIPAddr, ".")
    If UBound(sParts) <> 3 Then
        Z_AddressStringToLong = 0
        Exit Function
    End If

    '
    ' Build the long value out of the hex of the extracted strings
    '
    Z_AddressStringToLong = Val("&H" & Right$("00" & Hex$(sParts(3)), 2) & _
            Right$("00" & Hex$(sParts(2)), 2) & _
            Right$("00" & Hex$(sParts(1)), 2) & _
            Right$("00" & Hex$(sParts(0)), 2))

End Function



Public Function PSINET_GetLocalHostName$()
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2002
'
'****************************************************************************
'
'                     NAME: Function PSINET_GetLocalHostName
'
'                          ) As String
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    25 April 2002   First created for WebSneak
'
'                  PURPOSE: Returns the name of the local host
'
'****************************************************************************
'
'
Dim sHostName$
Dim iPtr%

    '
    ' Create a buffer
    '
    sHostName = String$(256, vbNullChar)

    '
    ' Init winsock api
    '
    If Not Z_SocketsInitialize() Then Exit Function

    '
    ' Get the local hosts name
    '
    If GetHostName(sHostName, Len(sHostName)) = ERROR_SUCCESS Then
        iPtr = InStr(sHostName, vbNullChar)
        If iPtr > 1 Then PSINET_GetLocalHostName = Mid$(sHostName, 1, iPtr - 1)
    End If
    
    '
    ' Cleanup the sockets library
    '
    WSACleanup

End Function




