Attribute VB_Name = "General"
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
' LANGUAGE:             Microsoft Visual Basic V5
'
' MODULE NAME:          PIVOTAL_General
'
' MODULE TYPE:          BASIC Module
'
' FILE NAME:            PSGEN.BAS
'
' MODIFICATION HISTORY: Steve O'Hara    03 February 1997   First created for DocBlazer
'
' PURPOSE:              Provides general purpose functions and procedures.
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
    ' Error base
    '
    Public Const PSGEN_ERROR_BASE = 10000
    
    '
    ' Common HTTP entities
    '
    Public Const PSGEN_HTTP_ENTITY_ALLOW = "Allow"
    Public Const PSGEN_HTTP_ENTITY_ENCODING = "Content-Encoding"
    Public Const PSGEN_HTTP_ENTITY_LENGTH = "Content-Length"
    Public Const PSGEN_HTTP_ENTITY_DISPOSITION = "Content-Disposition"
    Public Const PSGEN_HTTP_ENTITY_TYPE = "Content-Type"
    Public Const PSGEN_HTTP_ENTITY_EXPIRES = "Expires"
    Public Const PSGEN_HTTP_ENTITY_MODIFIED = "Last-Modified"
    Public Const PSGEN_HTTP_ENTITY_DATE = "Date"
    Public Const PSGEN_HTTP_ENTITY_FROM = "From"
    Public Const PSGEN_HTTP_ENTITY_IF_MODIFIED = "If-Modified-Since"
    Public Const PSGEN_HTTP_ENTITY_LOCATION = "Location"
    Public Const PSGEN_HTTP_ENTITY_PRAGMA = "Pragma"
    Public Const PSGEN_HTTP_ENTITY_REFERER = "Referer"
    Public Const PSGEN_HTTP_ENTITY_SERVER = "Server"
    Public Const PSGEN_HTTP_ENTITY_USER_AGENT = "User-Agent"
    Public Const PSGEN_HTTP_ENTITY_AUTHENTICATE = "WWW-Authenticate"

    '
    ' Mimetype file sections
    '
    Public Const PSGEN_MIMETYPE_FILE = "mimetypes.ini"
    Public Const PSGEN_MIMETYPE_SECTION = "MimeTypes"
    Public Const PSGEN_MIMETYPE_SECTION_REVERSE = "Reverse MimeTypes"
    
    '
    ' ImageMagick modules location
    '
    Public Const PSGEN_IMAGEMAGICK = "MAGICK_MODULE_PATH"
    
    '
    ' System setup constants
    '
    Public Const PSGEN_BASE_SCHEDULE = "SERVER_BASE_SCHEDULE"
    Public Const PSGEN_DEFAULT_SCHEDULE = 30
    Public Const PSGEN_TICK_PERIOD = "SERVER_TICK_PERIOD"
    Public Const PSGEN_DEFAULT_TICK_PERIOD = 10
    Dim miShutdownRequested As Boolean
    Public gbPollForTaskFinish As Boolean
    
    '
    ' Document definitions
    '
    Public Type PSGEN_FIELD_VALUES
        sName As String
        sValue As String
    End Type
    
    Public Type PSGEN_DOCUMENT_INFO
        astFields() As PSGEN_FIELD_VALUES  ' Field couplets for the source/destination data
        sFormat As String                  ' Format filename
        sOriginalFile As String            ' FTP, DIR or MAIL Attachment filename
        sCrystalPrint As String            ' Printable file version of a Crystal Report
        asTmpWorkFiles() As String         ' Temporary work filenames
        iNoOfTmpWorkFiles As Integer       ' Length of the dynamic array
        asBLOBFiles() As String            ' Temporary work filenames
        iNoOfBLOBFiles As Integer          ' Length of the dynamic array
        sMailFrom As String                ' Mail sender
        sMailSubject As String             ' Mail subject
        sMailMessage As String             ' Mail message body
        sMailID As String                  ' Mail message ID for MAPI
        lNode As Long                      ' Node ID
        iNodeSubType As Integer            ' Node sub-type (e-mail, ftp etc.)
        iStatus As Integer                 ' Status of the document
    End Type

    '
    ' Node sub-types
    '
    Public Const PSGEN_SUBTYPE_EMAIL = 0
    Public Const PSGEN_SUBTYPE_FTP = 1
    Public Const PSGEN_SUBTYPE_DIR = 2
    Public Const PSGEN_SUBTYPE_DEL = 3
    Public Const PSGEN_SUBTYPE_EDIT = 4
    
    '
    ' Encryption constants
    '
    Const ENCRYPT_START_CHAR = 127
    Const NO_OF_CHARS = 255
    
    '
    ' RTF to HTML Conversion
    '
    Declare Function ConvertRtfToHTML Lib "irun.dll" Alias "EXRTF2WEB" (ByVal sRtfFile$, sHtmlFile$, ByVal lOptions%, ByVal sBackColor As Any, ByVal sTitle As Any, ByVal lDPI%) As Integer
    
    Public Const EXO_RESULTS = 1
    Public Const EXO_INLINECSS = 2
    Public Const EXO_WMF2GIF = 4
    Public Const EXO_XML = 8
    Public Const EXO_HTML = 16
    Public Const EXO_MEMORY = 20
    Public Const EXO_NOHEADER = 64
    
    '
    ' Handles used by some of the functions
    '
    Dim mlDesktop&

    '
    ' Windows API calls
    '
    Type SHITEMID
        cb As Long
        abID() As Byte
    End Type
    
    Type ITEMIDLIST
        mkid As SHITEMID
    End Type

    Type BROWSEINFO
        hOwner As Long
        pidlRoot As Long
        pszDisplayName As String
        lpszTitle As String
        ulFlags As Long
        lpfn As Long
        lParam As Long
        iImage As Long
    End Type
    Dim mstBrowse As BROWSEINFO
    Dim msBrowseInitDir$
    Dim msBrowseTitle$

    Public Const WM_USER = &H400
    Public Const WM_INITDIALOG = &H110
    Public Const BFFM_INITIALIZED = 1
    Public Const BFFM_SELCHANGED = 2
    Public Const BFFM_VALIDATEFAILED = 4
    Public Const BFFM_ENABLEOK = (WM_USER + 101)
    Public Const BFFM_SETSELECTION = (WM_USER + 103)
    Public Const BFFM_SETSTATUSTEXT = (WM_USER + 104)
    
    Public Const CSIDL_DESKTOP = &H0
    Public Const CSIDL_PROGRAMS = &H2
    Public Const CSIDL_CONTROLS = &H3
    Public Const CSIDL_PRINTERS = &H4
    Public Const CSIDL_PERSONAL = &H5
    Public Const CSIDL_FAVORITES = &H6
    Public Const CSIDL_STARTUP = &H7
    Public Const CSIDL_RECENT = &H8
    Public Const CSIDL_SENDTO = &H9
    Public Const CSIDL_BITBUCKET = &HA
    Public Const CSIDL_STARTMENU = &HB
    Public Const CSIDL_DESKTOPDIRECTORY = &H10
    Public Const CSIDL_DRIVES = &H11
    Public Const CSIDL_NETWORK = &H12
    Public Const CSIDL_NETHOOD = &H13
    Public Const CSIDL_FONTS = &H14
    Public Const CSIDL_TEMPLATES = &H15
    Public Const CSIDL_HISTORY = &H22
    Public Const CSIDL_INTERNET = &H1
    Public Const CSIDL_INTERNET_CACHE = &H20

    
    Public Const BIF_RETURNONLYFSDIRS = &H1
    Public Const BIF_DONTGOBELOWDOMAIN = &H2
    Public Const BIF_STATUSTEXT = &H4
    Public Const BIF_RETURNFSANCESTORS = &H8
    Public Const BIF_BROWSEFORCOMPUTER = &H1000
    Public Const BIF_BROWSEFORPRINTER = &H2000

    Private Type ChooseColor
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        rgbResult As Long
        lpCustColors As String
        Flags As Long
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
    End Type
    
    Private Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        Flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
    End Type
    
    Type POINTAPI
        X As Long
        Y As Long
    End Type

    Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
    End Type

    Type PICBMP
       size As Long
       Type As Long
       hBmp As Long
       hPal As Long
       Reserved As Long
    End Type

    Type BITMAP '14 bytes
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
    End Type

    Type GUID
       Data1 As Long
       Data2 As Integer
       Data3 As Integer
       Data4(7) As Byte
    End Type

    Type PALETTEENTRY
       peRed As Byte
       peGreen As Byte
       peBlue As Byte
       peFlags As Byte
    End Type

    Type LOGPALETTE
       palVersion As Integer
       palNumEntries As Integer
       palPalEntry(255) As PALETTEENTRY  ' Enough for 256 colors
    End Type

    Public Const RASTERCAPS As Long = 38
    Public Const RC_PALETTE As Long = &H100
    Public Const SIZEPALETTE As Long = 104

    Type PROCESS_INFORMATION
        hProcess As Long
        hThread As Long
        dwProcessId As Long
        dwThreadId As Long
    End Type
    
    Type STARTUPINFO
        cb As Long
        lpReserved As String
        lpDesktop As String
        lpTitle As String
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Integer
        cbReserved2 As Integer
        lpReserved2 As Byte
        hStdInput As Long
        hStdOutput As Long
        hStdError As Long
    End Type
    
    Public Const OFN_ALLOWMULTISELECT = &H200
    Public Const OFN_CREATEPROMPT = &H2000
    Public Const OFN_ENABLEHOOK = &H20
    Public Const OFN_ENABLETEMPLATE = &H40
    Public Const OFN_ENABLETEMPLATEHANDLE = &H80
    Public Const OFN_EXPLORER = &H80000                         '  new look commdlg
    Public Const OFN_EXTENSIONDIFFERENT = &H400
    Public Const OFN_FILEMUSTEXIST = &H1000
    Public Const OFN_HIDEREADONLY = &H4
    Public Const OFN_LONGNAMES = &H200000                       '  force long names for 3.x modules
    Public Const OFN_NOCHANGEDIR = &H8
    Public Const OFN_NODEREFERENCELINKS = &H100000
    Public Const OFN_NOLONGNAMES = &H40000                      '  force no long names for 4.x modules
    Public Const OFN_NONETWORKBUTTON = &H20000
    Public Const OFN_NOREADONLYRETURN = &H8000
    Public Const OFN_NOTESTFILECREATE = &H10000
    Public Const OFN_NOVALIDATE = &H100
    Public Const OFN_OVERWRITEPROMPT = &H2
    Public Const OFN_PATHMUSTEXIST = &H800
    Public Const OFN_READONLY = &H1
    Public Const OFN_SHAREAWARE = &H4000
    Public Const OFN_SHAREFALLTHROUGH = 2
    Public Const OFN_SHARENOWARN = 1
    Public Const OFN_SHAREWARN = 0
    Public Const OFN_SHOWHELP = &H10
    
    Public Const INVALID_HANDLE_VALUE = -1
    Public Const MAX_PATH = 260
    Public Const NO_ERROR = 0
    Public Const PSFTP_FILE_ATTRIBUTE_READOPSY = &H1
    Public Const PSFTP_FILE_ATTRIBUTE_HIDDEN = &H2
    Public Const PSFTP_FILE_ATTRIBUTE_SYSTEM = &H4
    Public Const PSFTP_FILE_ATTRIBUTE_DIRECTORY = &H10
    Public Const PSFTP_FILE_ATTRIBUTE_ARCHIVE = &H20
    Public Const PSFTP_FILE_ATTRIBUTE_NORMAL = &H80
    Public Const PSFTP_FILE_ATTRIBUTE_TEMPORARY = &H100
    Public Const PSFTP_FILE_ATTRIBUTE_COMPRESSED = &H800
    Public Const PSFTP_FILE_ATTRIBUTE_OFFLINE = &H1000
    
    Type FILETIME
            dwLowDateTime As Long
            dwHighDateTime As Long
    End Type
    
    Type WIN32_FIND_DATA
            dwFileAttributes As Long
            ftCreationTime As Currency
            ftLastAccessTime As Currency
            ftLastWriteTime As Currency
            nFileSizeHigh As Long
            nFileSizeLow As Long
            dwReserved0 As Long
            dwReserved1 As Long
            cFileName As String * MAX_PATH
            cAlternate As String * 14
    End Type
    Public Const ERROR_NO_MORE_FILES = 18
    
    Public Const HH_DISPLAY_TOPIC = &H0
    Public Const HH_HELP_FINDER = &H0            ' WinHelp equivalent
    Public Const HH_DISPLAY_TOC = &H1            ' WinHelp equivalent
    Public Const HH_DISPLAY_INDEX = &H2         ' WinHelp equivalent
    Public Const HH_DISPLAY_SEARCH = &H3        ' not currently implemented
    Public Declare Function htmlhelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwdata As Long) As Long
    
    '
    ' User defined type required by Shell_NotifyIcon API call
    '
    Public Type PSGEN_NOTIFY_ICON_DATA
        lSize As Long
        lWnd As Long
        lID As Long
        lFlags As Long
        lCallBackMessage As Long
        lIcon As Long
        sTip As String * 255
    End Type
    
    Public Const NIM_ADD = &H0
    Public Const NIM_MODIFY = &H1
    Public Const NIM_DELETE = &H2
    Public Const NIF_MESSAGE = &H1
    Public Const NIF_ICON = &H2
    Public Const NIF_TIP = &H4
    
    Public Const WM_MOVE = &H3
    Public Const WM_CLOSE = &H10
    Public Const WM_SETTEXT = &HC
    Public Const WM_MOUSEMOVE = &H200
    Public Const WM_MOUSEWHEEL = &H20A
    Public Const WM_LBUTTONDOWN = &H201     'Button down
    Public Const WM_LBUTTONUP = &H202       'Button up
    Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
    Public Const WM_RBUTTONDOWN = &H204     'Button down
    Public Const WM_RBUTTONUP = &H205       'Button up
    Public Const WM_RBUTTONDBLCLK = &H206   'Double-click
    Public Const WM_CONTEXTMENU = &H7B
    Public Const WM_KEYDOWN = &H100
    Public Const WM_KEYUP = &H101
    
    Public Const EM_SETMODIFY = WM_USER + 9
    Public Const LB_SETTABSTOPS = &H192
    Public Const LB_SETHORIZONTALEXTENT = &H194
    Public Const LB_FINDSTRING = &H18F
    Public Const SWP_NOMOVE = 2
    Public Const SWP_NOSIZE = 1
    Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE
    Public Const HWND_BOTTOM = 0
    Public Const HWND_TOPMOST = -1
    Public Const HWND_NOTOPMOST = -2
    Public Const GW_OWNER = 4
    Public Const SWW_HPARENT = -8
    Public Const VK_ALTERNATE = &H12
    Public Const VK_CONTROL = &H11
    Public Const VK_SHIFT = &H10
    Public Const VK_ESCAPE = &H1B
    Public Const VK_F1 = &H70
    Public Const VK_F2 = &H71
    Public Const VK_F3 = &H72
    Public Const VK_F4 = &H73
    Public Const VK_F5 = &H74
    Public Const VK_F6 = &H75
    Public Const VK_F7 = &H76
    Public Const VK_F8 = &H77
    Public Const VK_F9 = &H78
    Public Const VK_F10 = &H79
    Public Const VK_F11 = &H7A
    Public Const VK_F12 = &H7B
    
    
    Public Const BM_CLICK = &HF5
    Public Const HWND_BROADCAST As Long = &HFFFF&
    Public Const WM_WININICHANGE As Long = &H1A

    '
    ' GetWindow () Constants
    '
    Public Const GW_CHILD = 5
    Public Const GW_HWNDFIRST = 0
    Public Const GW_HWNDLAST = 1
    Public Const GW_HWNDNEXT = 2
    Public Const GW_HWNDPREV = 3
    Public Const GWW_ID = (-12)
    
    Public Const DT_BOTTOM = &H8
    Public Const DT_CALCRECT = &H400
    Public Const DT_CENTER = &H1
    Public Const DT_CHARSTREAM = 4          '  Character-stream, PLP
    Public Const DT_DISPFILE = 6            '  Display-file
    Public Const DT_EXPANDTABS = &H40
    Public Const DT_EXTERNALLEADING = &H200
    Public Const DT_INTERNAL = &H1000
    Public Const DT_LEFT = &H0
    Public Const DT_METAFILE = 5            '  Metafile, VDM
    Public Const DT_NOCLIP = &H100
    Public Const DT_NOPREFIX = &H800
    Public Const DT_PLOTTER = 0             '  Vector plotter
    Public Const DT_RASCAMERA = 3           '  Raster camera
    Public Const DT_RASDISPLAY = 1          '  Raster display
    Public Const DT_RASPRINTER = 2          '  Raster printer
    Public Const DT_RIGHT = &H2
    Public Const DT_SINGLELINE = &H20
    Public Const DT_TABSTOP = &H80
    Public Const DT_TOP = &H0
    Public Const DT_VCENTER = &H4
    Public Const DT_WORDBREAK = &H10
    Public Const DT_END_ELLIPSIS = &H8000
    Public Const DT_WORD_ELLIPSIS = &H40000
    
    '
    ' GetWindowWord () Constants
    '
    Public Const GWW_HINSTANCE = (-6)
    
    Public Const SND_SYNC = &H0
    Public Const SND_ASYNC = &H1
    Public Const SND_NODEFAULT = &H2
    Public Const SND_LOOP = &H8
    Public Const SND_NOSTOP = &H10
    Public Const SND_MEMORY = &H4
    
    Public Const SRCCOPY = &HCC0020         ' (DWORD) dest = source
    Public Const SRCAND = &H8800C6          ' (DWORD) dest = source AND dest
    Public Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
    Public Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
    Public Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
    Public Const NOTSRCCOPY = &H330008      ' (DWORD) dest = (NOT source)
    
    Public Const NORMAL_PRIORITY_CLASS = &H20
    Public Const HIGH_PRIORITY_CLASS = &H80
    Public Const CREATE_NEW_CONSOLE = &H10
    Public Const CREATE_NEW_PROCESS_GROUP = &H200
    
    Public Const INFINITE = &HFFFFFFFF      '  Infinite timeout
    Public Const WAIT_TIMEOUT = &H102
    Public Const STARTF_FORCEOFFFEEDBACK = &H80
    Public Const STARTF_FORCEONFEEDBACK = &H40
    Public Const STARTF_RUNFULLSCREEN = &H20        '  ignored for non-x86 platforms
    Public Const STARTF_USECOUNTCHARS = &H8
    Public Const STARTF_USEFILLATTRIBUTE = &H10
    Public Const STARTF_USEPOSITION = &H4
    Public Const STARTF_USESHOWWINDOW = &H1
    Public Const STARTF_USESIZE = &H2
    Public Const STARTF_USESTDHANDLES = &H100
    Public Const CREATE_SEPARATE_WOW_VDM = &H800
    Public Const SW_ERASE = &H4
    Public Const SW_HIDE = 0
    Public Const SW_INVALIDATE = &H2
    Public Const SW_MAX = 10
    Public Const SW_MAXIMIZE = 3
    Public Const SW_MINIMIZE = 6
    Public Const SW_NORMAL = 1
    Public Const SW_OTHERUNZOOM = 4
    Public Const SW_OTHERZOOM = 2
    Public Const SW_PARENTCLOSING = 1
    Public Const SW_PARENTOPENING = 3
    Public Const SW_RESTORE = 9
    Public Const SW_SCROLLCHILDREN = &H1
    Public Const SW_SHOW = 5
    Public Const SW_SHOWDEFAULT = 10
    Public Const SW_SHOWMAXIMIZED = 3
    Public Const SW_SHOWMINIMIZED = 2
    Public Const SW_SHOWMINNOACTIVE = 7
    Public Const SW_SHOWNA = 8
    Public Const SW_SHOWNOACTIVATE = 4
    Public Const SW_SHOWNORMAL = 1
    
    Public Const R2_BLACK = 1       '   0
    Public Const R2_COPYPEN = 13    '  P
    Public Const R2_LAST = 16
    Public Const R2_MASKNOTPEN = 3  '  DPna
    Public Const R2_MASKPEN = 9     '  DPa
    Public Const R2_MASKPENNOT = 5  '  PDna
    Public Const R2_MERGENOTPEN = 12 '  DPno
    Public Const R2_MERGEPEN = 15   '  DPo
    Public Const R2_MERGEPENNOT = 14 '  PDno
    Public Const R2_NOP = 11        '  D
    Public Const R2_NOT = 6         '  Dn
    Public Const R2_NOTCOPYPEN = 4  '  PN
    Public Const R2_NOTMASKPEN = 8  '  DPan
    Public Const R2_NOTMERGEPEN = 2 '  DPon
    Public Const R2_NOTXORPEN = 10  '  DPxn
    Public Const R2_WHITE = 16      '   1
    Public Const R2_XORPEN = 7      '  DPx

    Public Const TRANSPARENT = 1
    Public Const OPAQUE = 2
        
    '
    ' Window placement
    '
    Public Const CW_USEDEFAULT = &H80000000
    
    '
    ' Registry API constants. The first, ERROR_SUCCESS, is used to check that API registry
    ' functions have completed successfully. The second indicates a 'base' registry key that
    ' the license key will be saved in.
    '
    ' The next two constants, respectively, define how a license key will be written to/
    ' retrieved from the registry (in binary form, so a Byte Array will be used), and how the
    ' registry value will be stored (non-volatile, meaning that the license key always remains
    ' in the registry after the machine is restarted).
    '
    Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
    End Type
    Public Const PROCESS_ALL_ACCESS = &H1F0FFF
    Public Const PROCESS_TERMINATE = &H1&
    
    Public Const ERROR_SUCCESS = 0&                ' Indicates that everything's okay
    Public Const ERROR_MORE_DATA = 234&            ' Indicates that there is more data to
                                                    ' retrieve from the registry
    Public Const ERROR_FILE_NOT_FOUND = 2&         ' Indicates that a license key (or license key value
                                                    ' does not exist)
    
    Public Const GWL_WNDPROC = -4
    Public Const GWL_USERDATA = -21
    Public Const GWL_STYLE = -16
    
    '
    ' Edge function values
    '
    Public Const BDR_RAISEDOUTER = &H1
    Public Const BDR_SUNKENOUTER = &H2
    Public Const BDR_RAISEDINNER = &H4
    Public Const BDR_SUNKENINNER = &H8
    Public Const BDR_OUTER = &H3
    Public Const BDR_INNER = &HC
    Public Const BDR_RAISED = &H5
    Public Const BDR_SUNKEN = &HA
    Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
    Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
    Public Const EDGE_RAISED_THIN = BDR_RAISEDINNER
    Public Const EDGE_SUNKEN_THIN = BDR_SUNKENOUTER
    Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
    Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
    Public Const BF_LEFT = &H1
    Public Const BF_TOP = &H2
    Public Const BF_RIGHT = &H4
    Public Const BF_BOTTOM = &H8
    Public Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
    Public Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
    Public Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
    Public Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
    Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
    Public Const BF_DIAGONAL = &H10       ' For diagonal lines, the BF_RECT flags specify the end point of
                                          ' the vector bounded by the rectangle parameter.
    Public Const BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
    Public Const BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
    Public Const BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
    Public Const BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)
    Public Const BF_MIDDLE = &H800    ' Fill in the middle.
    Public Const BF_SOFT = &H1000     ' Use for softer buttons.
    Public Const BF_ADJUST = &H2000   ' Calculate the space left over.
    Public Const BF_FLAT = &H4000     ' For flat rather than 3-D borders.
    Public Const BF_MONO = &H8000     ' For monochrome borders.
    
    Public Const CB_SETDROPPEDWIDTH = &H160
    
    '
    ' System Metrics
    '
    Public Const SM_CXVSCROLL = 2
    Public Const SM_CYHSCROLL = 3
    Public Const SM_CYVSCROLL = 20
    Public Const SM_CXHSCROLL = 21
    
    '
    ' Operating system version information
    '
    Public Const VER_PLATFORM_WIN32s = 0
    Public Const VER_PLATFORM_WIN32_WINDOWS = 1
    Public Const VER_PLATFORM_WIN32_NT = 2
    
    Public Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
    End Type
    
    Public Type WINDOWPLACEMENT
        length As Long
        Flags As Long
        showCmd As Long
        ptMinPosition As POINTAPI
        ptMaxPosition As POINTAPI
        rcNormalPosition As RECT
    End Type
    
    '
    ' Menu stuff
    '
    Public Const LB_RESETCONTENT = (WM_USER + 5)
    Public Const MF_INSERT = &H0
    Public Const MF_CHANGE = &H80
    Public Const MF_APPEND = &H100
    Public Const MF_DELETE = &H200
    Public Const MF_REMOVE = &H1000
    Public Const MF_BYCOMMAND = &H0
    Public Const MF_BYPOSITION = &H400
    Public Const MF_SEPARATOR = &H800
    Public Const MF_ENABLED = &H0
    Public Const MF_GRAYED = &H1
    Public Const MF_DISABLED = &H2
    Public Const MF_UNCHECKED = &H0
    Public Const MF_CHECKED = &H8
    Public Const MF_USECHECKBITMAPS = &H200
    Public Const MF_STRING = &H0
    Public Const MF_BITMAP = &H4
    Public Const MF_OWNERDRAW = &H100
    Public Const MF_POPUP = &H10
    Public Const MF_MENUBARBREAK = &H20
    Public Const MF_MENUBREAK = &H40
    Public Const MF_UNHILITE = &H0
    Public Const MF_HILITE = &H80
    Public Const MF_SYSMENU = &H2000
    Public Const MF_RIGHTJUSTIFY = &H4000&
    Public Const MF_HELP = &H4000
    Public Const MF_MOUSESELECT = &H8000
    Public Const MF_END = &H80
    Public Const MFS_CHECKED = MF_CHECKED
    Public Const MFS_ENABLED = MF_ENABLED
    Public Const MFS_GRAYED = &H3&
    Public Const MFS_HILITE = MF_HILITE
    Public Const MFS_DISABLED = MFS_GRAYED
    Public Const MFS_UNCHECKED = MF_UNCHECKED
    Public Const MFS_UNHILITE = MF_UNHILITE
    Public Const MFT_BITMAP = MF_BITMAP
    Public Const MFT_MENUBARBREAK = MF_MENUBARBREAK
    Public Const MFT_MENUBREAK = MF_MENUBREAK
    Public Const MFT_OWNERDRAW = MF_OWNERDRAW
    Public Const MFT_RADIOCHECK = &H200&
    Public Const MFT_RIGHTJUSTIFY = MF_RIGHTJUSTIFY
    Public Const MFT_RIGHTORDER = &H2000&
    Public Const MFT_SEPARATOR = MF_SEPARATOR
    Public Const MFT_STRING = MF_STRING
    Public Const TPM_LEFTBUTTON = &H0
    Public Const TPM_RIGHTBUTTON = &H2
    Public Const TPM_LEFTALIGN = &H0
    Public Const TPM_CENTERALIGN = &H4
    Public Const TPM_RIGHTALIGN = &H8
    Public Const TPM_BOTTOMALIGN = &H20&
    Public Const TPM_HORIZONTAL = &H0&
    Public Const TPM_NONOTIFY = &H80&
    Public Const TPM_RETURNCMD = &H100&
    Public Const TPM_TOPALIGN = &H0&
    Public Const TPM_VCENTERALIGN = &H10&
    Public Const TPM_VERTICAL = &H40&
    
    '
    ' Message hook types
    '
    Public Const WH_KEYBOARD = 2
    Public Const WH_CALLWNDPROC = 4
    Public Const WH_CBT = 5
    Public Const WH_DEBUG = 9
    Public Const WH_FOREGROUNDIDLE = 11
    Public Const WH_GETMESSAGE = 3
    Public Const WH_HARDWARE = 8
    Public Const WH_JOURNALPLAYBACK = 1
    Public Const WH_JOURNALRECORD = 0
    Public Const WH_MAX = 11
    Public Const WH_MIN = (-1)
    Public Const WH_MOUSE = 7
    Public Const WH_MSGFILTER = (-1)
    Public Const WH_SHELL = 10
    Public Const WH_SYSMSGFILTER = 6
    
    Public Const GWL_HWNDPARENT = (-8)
    
    Public Const STRETCH_ANDSCANS = 1
    Public Const STRETCH_DELETESCANS = 3
    Public Const STRETCH_HALFTONE = 4
    Public Const STRETCH_ORSCANS = 2
    
    Public Const SC_MAXIMIZE = &HF030&
    Public Const SC_MINIMIZE = &HF020&
    Public Const SC_RESTORE = &HF120&
    Public Const WM_SYSCOMMAND = &H112
    Public Const WM_COMMAND = &H111
    
    Public Const WS_MINIMIZE = &H20000000
    Public Const WS_MAXIMIZE = &H1000000
    Public Const WS_DLGFRAME = &H400000
    Public Const WS_POPUP = &H80000000
    
    Public Type CHARRANGE
        cpMin As Long     ' First character of range (0 for start of doc)
        cpMax As Long     ' Last character of range (-1 for end of doc)
    End Type

    Public Type FORMATRANGE
        hDC As Long       ' Actual DC to draw on
        hdcTarget As Long ' Target DC for determining text formatting
        rc As RECT        ' Region of the DC to draw to (in twips)
        rcPage As RECT    ' Region of the entire DC (page size) (in twips)
        chrg As CHARRANGE ' Range of text to draw (see above declaration)
    End Type
    
    Public Type BLEND_PROPS
        tBlendOp As Byte
        tBlendOptions As Byte
        tBlendAmount As Byte
        tAlphaType As Byte
    End Type
    
    Public Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
    Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
    Public Declare Function EnumProcesses Lib "PSAPI.DLL" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
    Public Declare Function GetModuleFileNameExA Lib "PSAPI.DLL" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
    Public Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
    Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
    Public Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
    
    Const TH32CS_SNAPPROCESS As Long = 2&
    
    Public Type PROCESSENTRY32
        dwSize As Long 'Specifies the length, in bytes, of the structure.
        cntUsage As Long 'Number of references to the process.
        th32ProcessID As Long 'Identifier of the process.
        th32DefaultHeapID As Long 'Identifier of the default heap for the process.
        th32ModuleID As Long 'Module identifier of the process. (Associated exe)
        cntThreads As Long 'Number of execution threads started by the process.
        th32ParentProcessID As Long 'Identifier of the process that created the process being examined.
        pcPriClassBase As Long 'Base priority of any threads created by this process.
        dwFlags As Long 'Reserved; do not use.
        szExeFile As String * MAX_PATH 'Path and filename of the executable file for the process.
    End Type

    Public Const WM_GETICON = &H7F
    Public Const WM_SETICON = &H80
    Public Const ICON_SMALL = 0
    Public Const ICON_BIG = 1

    
    Public Const FO_COPY As Long = &H2
    Public Const FO_MOVE = &H1
    Public Const FO_RENAME = &H4
    
    Public Const FOF_SILENT As Long = &H4
    Public Const FOF_RENAMEONCOLLISION As Long = &H8
    Public Const FOF_NOCONFIRMATION As Long = &H10
    
    Public Const SHARD_PATH = &H2&
    Public Const SHCNF_IDLIST = &H0
    Public Const SHCNE_ALLEVENTS = &H7FFFFFFF
    
    Public Type SHFILEOPSTRUCT
      hWnd      As Long
      wFunc      As Long
      pFrom      As String
      pTo        As String
      fFlags     As Integer
      fAborted   As Boolean
      hNameMaps  As Long
      sProgress  As String
    End Type
    
    Declare Function SHFileOperation Lib "shell32" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
    Declare Function SHAddToRecentDocs Lib "Shell32.dll" (ByVal dwFlags As Long, ByVal dwdata As String) As Long
    
    Public Const EM_FORMATRANGE As Long = WM_USER + 57
    Public Const EM_SETTARGETDEVICE As Long = WM_USER + 72
    Public Const PHYSICALOFFSETX As Long = 112
    Public Const PHYSICALOFFSETY As Long = 113

    Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
   
    ' Clipboard functions:
    Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
    Declare Function CloseClipboard Lib "user32" () As Long
    Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
    Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
    Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
    
    ' Memory functions:
    Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
    Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
    
    Declare Function ChooseColorAPI Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
    Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
    Declare Function SetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
    Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Declare Function SendMessageSpecial Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Integer) As Long
    Declare Function CreateCaret Lib "user32" (ByVal hWnd As Long, ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
    Declare Function ShowCaret Lib "user32" (ByVal hWnd As Long) As Long
    Declare Function SetParent Lib "user32" (ByVal hChild As Long, ByVal wParent As Long) As Long
    Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
    Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
    Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
    Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
    Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
    Declare Function SetWindowWord Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
    Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
    Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
    Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long
    Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
    Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
    Declare Function ReleaseCapture Lib "user32" () As Long
    Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
    Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
    Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
    Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
    Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
    Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
    Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
    Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
    Declare Function GetObjectA Lib "gdi32" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
    Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
    Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
    Declare Function ValidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
    Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
    Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
    Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
    Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
    Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
    Declare Function Polyline Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
    Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
    Declare Function SetROP2 Lib "gdi32" (ByVal hDC As Long, ByVal nDrawMode As Long) As Long
    Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long, ByVal blendFunct As Long) As Boolean
    Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
    Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
    Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal lMessage As Long, pNID As PSGEN_NOTIFY_ICON_DATA) As Boolean
    Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle&, ByVal dwMilliseconds&) As Long
    Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName$, ByVal lpCommandLine$, lpProcessAttributes As SECURITY_ATTRIBUTES, ByVal lpThreadAttributes&, ByVal bInheritHandles&, ByVal dwCreationFlags&, ByVal lpEnvironment&, ByVal lpCurrentDirectory$, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
    Declare Function CloseHandle Lib "kernel32" (ByVal hObject&) As Long
    Declare Function GetTickCount Lib "kernel32" () As Long
    Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
    Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
    Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
    Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
    Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
    Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
    Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
    Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)
    Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
    Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    Declare Function GetDesktopWindow Lib "user32" () As Long
    Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
    Declare Function GetWindowWord Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Integer
    Declare Function SHGetPathFromIDList Lib "Shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
    Declare Function SHGetSpecialFolderLocation Lib "Shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
    Declare Function SHGetSpecialFolderLocationLong Lib "Shell32.dll" Alias "SHGetSpecialFolderLocation" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
    Declare Sub CoTaskMemFree Lib "OLE32.DLL" (ByVal pv As Long)
    Declare Function SHBrowseForFolder Lib "Shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
    Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
    Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
    Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
    Declare Function SetWinFocus Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
    Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
    Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
    Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
    Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
    Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
    Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
    Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
    Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
    Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
    Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
    Declare Function GetProcessVersion Lib "kernel32" (ByVal lProcessId&) As Long
    Declare Function GetCurrentProcessId Lib "kernel32" () As Long
    Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long

    Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
    Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
    Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
    Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
    Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
    Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
    Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
    Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
    Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
    Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
    Declare Function RegGetKeySecurity Lib "advapi32.dll" (ByVal hKey As Long, ByVal lInfo&, lpSecurityAttributes As SECURITY_ATTRIBUTES, lLength&) As Long
    Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
    Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
    Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
    Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
    Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Declare Function WaitForInputIdle Lib "user32" (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long
    Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
    Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
    Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal lXVal&, ByVal lYVal&) As Long
    Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Declare Function FindExecutable Lib "Shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
    Declare Function IsZoomed Lib "user32" (ByVal hWnd As Long) As Long
    Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
    Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
    Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
    Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
   
    Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
    Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal oleColor&, ByVal hPalette&, pcColorRef&) As Long
    Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
    Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
    Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
    
    Declare Function AllocConsole Lib "kernel32" () As Long
    Declare Function FreeConsole Lib "kernel32" () As Long
    Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
    Declare Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" _
           (ByVal hConsoleOutput As Long, lpBuffer As Any, ByVal _
           nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, _
           lpReserved As Any) As Long
    Declare Function SetConsoleTitle Lib "kernel32" Alias "SetConsoleTitleA" (ByVal lpConsoleTitle As String) As Long
    
    Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
    
    Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
    Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
    Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
    
    Public Const STD_INPUT_HANDLE = -10&
    Public Const STD_OUTPUT_HANDLE = -11&
    Public Const STD_ERROR_HANDLE = -12&
    
    Declare Function WriteFile& Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite&, lpNumberOfBytesWritten&, ByVal lpOverlapped&)
    Declare Function CreateFile& Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName$, ByVal dwDesiredAccess&, ByVal dwShareMode&, ByVal lpSecurityAttributes&, ByVal dwCreationDisposition&, ByVal dwFlagsAndAttributes&, ByVal hTemplateFile&)
    Declare Function FlushFileBuffers& Lib "kernel32" (ByVal hFile&)
    
    Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
    Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
    Declare Function AppendMenuBynum Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Long) As Long
    Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As String) As Long
    Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
    Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
    Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
    Declare Function GetMenuCheckMarkDimensions Lib "user32" () As Long
    Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hWnd As Long, lprc As RECT) As Long
    Declare Function TrackPopupMenuBynum Lib "user32" Alias "TrackPopupMenu" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hWnd As Long, ByVal lprc As Long) As Long
    Declare Function ModifyMenuBynum Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Long) As Long
    Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
    Declare Function CreateBitmapIndirect Lib "gdi32" (lpBitmap As BITMAP) As Long
    Declare Function LoadMenu Lib "user32" Alias "LoadMenuA" (ByVal hInstance As Long, ByVal lpString As String) As Long
    Declare Function LoadMenuIndirect Lib "user32" Alias "LoadMenuIndirectA" (ByVal lpMenuTemplate As Long) As Long
    Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
    Declare Function SetMenu Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long) As Long
    Declare Function HiliteMenuItem Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long, ByVal wIDHiliteItem As Long, ByVal wHilite As Long) As Long
    Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
    Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
    Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
    Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
    Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Any) As Long
    Declare Function GetLastError Lib "kernel32" () As Long
    Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
    Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
    Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PICBMP, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
    Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
    Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
    Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hDC As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
    Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
    Declare Function SelectPalette Lib "gdi32" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
    Declare Function RealizePalette Lib "gdi32" (ByVal hDC As Long) As Long
    Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As Any, lpLocalFileTime As Any) As Long

    Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
    Public Const SPI_GETWORKAREA = 48

    Public Type LOGBRUSH
      lbStyle As Long
      lbColor As Long
      lbHatch As Long
    End Type
    
    Public Type WNDCLASS
      Style As Long
      lpfnWndProc As Long
      cbClsExtra As Long
      cbWndExtra As Long
      hInstance As Long
      hIcon As Long
      hCursor As Long
      hbrBackground As Long
      lpszMenuName As String
      lpszClassName As String
    End Type
    
    Public Type DRAWITEMSTRUCT
        CtlType As Long
        CtlID As Long
        itemID As Long
        itemAction As Long
        itemState As Long
        hwndItem As Long
        hDC As Long
        rcItem As RECT
        ItemData As Long
    End Type
    
    Public Const COLOR_HIGHLIGHT = 13
    Public Const COLOR_HIGHLIGHTTEXT = 14
    Public Const COLOR_WINDOW = 5
    Public Const COLOR_WINDOWTEXT = 8
    Public Const LB_GETTEXT = &H189
    Public Const WM_DRAWITEM = &H2B
    Public Const ODS_FOCUS = &H10
    Public Const ODT_LISTBOX = 2
    
    Public Const API_FALSE As Long = 0&
    Public Const API_TRUE As Long = 1&
    Public Const vbZLString As String = ""
    
    Public Const WS_EX_TOOLWINDOW As Long = &H80&
    
    Public Const WS_CHILD As Long = &H40000000
    
    Public Const WM_SIZE As Long = &H5&
    Public Const WM_NCLBUTTONDOWN As Long = &HA1&
    
    Public Const RDW_INVALIDATE As Long = &H1&
    Public Const RDW_INTERNALPAINT As Long = &H2&
    Public Const RDW_ERASE As Long = &H4&
    
    Public Const RDW_ALLCHILDREN As Long = &H80&
    
    Public Const RDW_UPDATENOW As Long = &H100&
    Public Const RDW_ERASENOW As Long = &H200&
    
    Public Const RDW_FRAME As Long = &H400&
    Public Const RDW_NOFRAME As Long = &H800&
    
    Public Const BS_SOLID As Long = 0&
    
    Public Const COLOR_BTNTEXT As Long = 18&
    
    Public Const CS_VREDRAW As Long = &H1&
    Public Const CS_HREDRAW As Long = &H2&
    Public Const CS_NOCLOSE As Long = &H200&
    Public Const CS_SAVEBITS As Long = &H800&
    
    Public Const IDC_SIZENWSE As Long = 32642&
    Public Const IDC_SIZENESW As Long = 32643&
    Public Const IDC_SIZEWE As Long = 32644&
    Public Const IDC_SIZENS As Long = 32645&
    
    Public Const HTLEFT As Long = 10&
    Public Const HTRIGHT As Long = 11&
    Public Const HTTOP As Long = 12&
    Public Const HTTOPLEFT As Long = 13&
    Public Const HTTOPRIGHT As Long = 14&
    Public Const HTBOTTOM As Long = 15&
    Public Const HTBOTTOMLEFT As Long = 16&
    Public Const HTBOTTOMRIGHT As Long = 17&
    
    ' SetWindowPos Flags
    Public Const SWP_NOZORDER As Long = &H4&
    Public Const SWP_NOREDRAW As Long = &H8&
    Public Const SWP_NOACTIVATE As Long = &H10&
    Public Const SWP_FRAMECHANGED As Long = &H20&
    Public Const SWP_SHOWWINDOW As Long = &H40&
    Public Const SWP_HIDEWINDOW As Long = &H80&
    Public Const SWP_NOCOPYBITS As Long = &H100&
    Public Const SWP_NOOWNERZORDER As Long = &H200&
    
    Public Const SWP_DRAWFRAME As Long = SWP_FRAMECHANGED
    Public Const SWP_NOREPOSITION As Long = SWP_NOOWNERZORDER
    
    Const PRINTER_ENUM_CONNECTIONS = &H4
    Const PRINTER_ENUM_LOCAL = &H2

    Type PRINTER_INFO_1
       Flags As Long
       pDescription As String
       PName As String
       PComment As String
    End Type

    Type PRINTER_INFO_4
       pPrinterName As String
       pServerName As String
       Attributes As Long
    End Type

    Declare Function EnumPrinters Lib "winspool.drv" Alias _
       "EnumPrintersA" (ByVal Flags As Long, ByVal Name As String, _
       ByVal Level As Long, pPrinterEnum As Long, ByVal cdBuf As Long, _
       pcbNeeded As Long, pcReturned As Long) As Long
    Declare Function PtrToStr Lib "kernel32" Alias "lstrcpyA" _
       (ByVal retval As String, ByVal Ptr As Long) As Long
    Declare Function StrLen Lib "kernel32" Alias "lstrlenA" _
       (ByVal Ptr As Long) As Long
    
    Public Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
    Public Declare Function ScreenToClientLong Lib "user32" Alias "ScreenToClient" (ByVal hWnd&, lpPoint&) As Long
    Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle&, ByVal lpClassName$, ByVal lpWindowName$, ByVal dwStyle&, ByVal X&, ByVal Y&, ByVal nWidth&, ByVal nHeight&, ByVal hwndParent&, ByVal hMenu&, ByVal hInstance&, lpParam As Any) As Long
    Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd&, ByVal wMsg&, ByVal wParam&, ByVal lParam&) As Long
    Public Declare Function LoadCursorBynum Lib "user32" Alias "LoadCursorA" (ByVal hInstance&, ByVal lpCursorName&) As Long
    Public Declare Function RedrawWindow Lib "user32" (ByVal hWnd&, lprcUpdate As RECT, ByVal hrgnUpdate&, ByVal fuRedraw&) As Long
    Public Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
    Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd&, ByVal lpString$, ByVal hData&) As Long
    Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd&, ByVal lpString$) As Long
    Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd&, ByVal lpString$) As Long
    Public Declare Function IsWindow Lib "user32" (ByVal hWnd&) As Long
    Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
    Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
    Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

    Public Declare Function ExtractIcon Lib "Shell32.dll" Alias "ExtractIconA" (ByVal hinst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
    Public Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long

    Public Declare Function HideCaret Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Function DestroyCaret Lib "user32" () As Long
    
    ' Retrieves information about an object in the file system, such as a file,
    ' a folder, a directory, or a drive root.
    Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" _
                                  (ByVal pszPath As Any, _
                                  ByVal dwFileAttributes As Long, _
                                  psfi As SHFILEINFO, _
                                  ByVal cbFileInfo As Long, _
                                  ByVal uFlags As Long) As Long
    
    Type SHFILEINFO
        hIcon As Long
        iIcon As Long
        dwAttributes As Long
        szDisplayName As String * MAX_PATH
        szTypeName As String * 80
    End Type
    
    ' uFlags:
    ' Flag that specifies the file information to retrieve. If uFlags includes the SHGFI_ICON
    ' or SHGFI_SYSICONINDEX value, the return value is the handle to the system image
    ' list that contains the large icon images. If the SHGFI_SMALLICON value is included
    ' with SHGFI_ICON or SHGFI_SYSICONINDEX, the return value is the handle to the
    ' image list that contains the small icon images.
    
    ' If uFlags does not include SHGFI_EXETYPE, SHGFI_ICON, SHGFI_SYSICONINDEX,
    ' or SHGFI_SMALLICON, the return value is nonzero if successful, or zero otherwise.
    
    ' This parameter can be a combination of the following values:
    
    ' Modifies SHGFI_ICON, causing the function to retrieve the file's large icon.
    Public Const SHGFI_LARGEICON = &H0&
    
    ' Modifies SHGFI_ICON, causing the function to retrieve the file's small icon.
    Public Const SHGFI_SMALLICON = &H1&
    
    ' Modifies SHGFI_ICON, causing the function to retrieve the file's open icon. A
    ' container object displays an open icon to indicate that the container is open.
    Public Const SHGFI_OPENICON = &H2&
    
    ' Modifies SHGFI_ICON, causing the function to retrieve a shell-sized icon. If
    ' this flag is not specified, the function sizes the icon according to the system
    ' metric values.
    ' The return value is *supposed to be* the handle of the system image list which
    ' could be passed to the ImageList_GetIconSize function to get the icon size.
    ' But the return value *is only* nonzero if successful, or zero otherwise.
    Public Const SHGFI_SHELLICONSIZE = &H4&
    
    ' Indicates that pszPath is the address of an ITEMIDLIST structure rather than
    ' a path name.
    Public Const SHGFI_PIDL = &H8&
    
    ' Indicates that the function should not attempt to access the file specified by
    ' pszPath. Rather, it should act as if the file specified by pszPath exists with
    ' the file attributes passed in dwFileAttributes. This flag cannot be combined
    ' with the SHGFI_ATTRIBUTES, SHGFI_EXETYPE, or SHGFI_PIDL flags.
    Public Const SHGFI_USEFILEATTRIBUTES = &H10&
    
    ' Retrieves the handle of the icon that represents the file and the index of the
    ' icon within the system image list. The handle is copied to the hIcon member
    ' of the structure specified by psfi, and the index is copied to the iIcon member.
    
    ' The return value is *supposed to be* the handle of the system image list,
    ' .....boolean instead...!!!
    ' ** SHGFI_ICON creates a copy of the icon in memory. The DestroyIcon **
    ' ** function must be called to free any memory the icon occupied.          **
    Public Const SHGFI_ICON = &H100&
    
    ' Retrieves the display name for the file. The name is copied to the szDisplayName
    ' member of the structure specified by psfi. The returned display name uses the
    ' long filename, if any, rather than the 8.3 form of the filename.
    Public Const SHGFI_DISPLAYNAME = &H200&
    
    ' Retrieves the string that describes the file's type. The string is copied to the
    ' szTypeName member of the structure specified by psfi.
    Public Const SHGFI_TYPENAME = &H400&
    
    ' Retrieves the file attribute flags. The flags are copied to the dwAttributes member
    ' of the structure specified by psfi. See the constants at the end of this file.
    Public Const SHGFI_ATTRIBUTES = &H800&
    
    ' Retrieves the name of the file that contains the icon representing the file. The
    ' name is copied to the szDisplayName member of the structure specified by psfi.
    Public Const SHGFI_ICONLOCATION = &H1000&
    
    ' Returns the type of the executable file if pszPath identifies an executable file.
    ' To retrieve the executable file type, uFlags must specify only SHGFI_EXETYPE.
    ' The return value specifies the type of the executable file:
    ' LowWord value       HighWord value         Type
    ' 0                                                           Nonexecutable file or an error condition.
    ' "NE" or "PE"          3.0, 3.5, or 4.0           Windows application
    ' "MZ"                      0                              MS-DOS .EXE, .COM or .BAT file
    ' "PE"                      0                              Win32 console application
    Public Const SHGFI_EXETYPE = &H2000&
    
    ' Constants to represent the strings as ASCII char codes
    Public Const EXE_WIN16 = &H454E    ' "NE"
    Public Const EXE_DOS16 = &H5A4D  ' "MZ"
    Public Const EXE_WIN32 = &H4550    ' "PE"
    'Public Const EXE_DOS32 = &H4543   ' "CE"
    
    ' Retrieves the index of the icon within the system image list. The index is copied to
    ' the iIcon member of the structure specified by psfi. The return value is the handle of
    ' the system image list.
    Public Const SHGFI_SYSICONINDEX = &H4000&
    
    ' Modifies SHGFI_ICON, causing the function to add the link overlay to the file's icon.
    Public Const SHGFI_LINKOVERLAY = &H8000&
    
    ' Modifies SHGFI_ICON, causing the function to blend the file's icon with the system
    ' highlight color.
    Public Const SHGFI_SELECTED = &H10000

    Public Declare Function DrawIconEx Lib "user32" _
                                (ByVal hDC As Long, _
                                 ByVal xLeft As Long, _
                                 ByVal yTop As Long, _
                                 ByVal hIcon As Long, _
                                 ByVal cxWidth As Long, _
                                 ByVal cyWidth As Long, _
                                 ByVal istepIfAniCur As Long, _
                                 ByVal hbrFlickerFreeDraw As Long, _
                                 ByVal diFlags As Long) As Boolean
    
    ' DrawIconEx() diFlags values:
    Public Const DI_MASK = &H1
    Public Const DI_IMAGE = &H2
    Public Const DI_NORMAL = &H3
    Public Const DI_COMPAT = &H4
    Public Const DI_DEFAULTSIZE = &H8
    
    Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
    Public Declare Function ImageList_GetIconSize Lib "comctl32" _
                            (ByVal himl As Long, _
                            cx As Long, _
                            cy As Long) As Boolean
                            
                            
    '
    ' Event log interface
    '
    Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
    
    Public Const EVENTLOG_ERROR_TYPE = 1
    Public Const EVENTLOG_WARNING_TYPE = 2
    Public Const EVENTLOG_INFORMATION_TYPE = 4
    Public Const EVENTLOG_AUDIT_SUCCESS = 8
    Public Const EVENTLOG_AUDIT_FAILURE = 10
    
    Public Enum LogEventTypes
        LogError = 1
        LogWarning = 2
        LogInformation = 4
    End Enum
    
    Public Declare Function OpenEventLog Lib "advapi32.dll" Alias "OpenEventLogA" (ByVal lpUNCServerName As String, ByVal lpSourceName As String) As Long
    Public Declare Function RegisterEventSource Lib "advapi32.dll" Alias "RegisterEventSourceA" (ByVal lpUNCServerName As String, ByVal lpSourceName As String) As Long
    Public Declare Function DeregisterEventSource Lib "advapi32.dll" (ByVal hEventLog As Long) As Long
    Public Declare Function CloseEventLog Lib "advapi32.dll" (ByVal hEventLog As Long) As Long
    Public Declare Function ReportEvent Lib "advapi32" Alias "ReportEventA" (ByVal hEventLog As Long, ByVal wType As Long, ByVal wCategory As Long, ByVal dwEventID As Long, ByVal lpUserSid As Long, ByVal wNumStrings As Long, ByVal dwDataSize As Long, lpStrings As Any, lpRawData As Any) As Long

    Public Type SYSTEMTIME
            wYear As Integer
            wMonth As Integer
            wDayOfWeek As Integer
            wDay As Integer
            wHour As Integer
            wMinute As Integer
            wSecond As Integer
            wMilliseconds As Integer
    End Type
    Public Type TIME_ZONE_INFORMATION
            Bias As Long
            StandardName(32) As Integer
            StandardDate As SYSTEMTIME
            StandardBias As Long
            DaylightName(32) As Integer
            DaylightDate As SYSTEMTIME
            DaylightBias As Long
    End Type
    
    Public Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
    Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long
    Public Declare Function GetCurrentThread Lib "kernel32" () As Long
    Public Declare Function GetThreadTimes Lib "Kernel32.dll" ( _
                 ByVal hThread As Long, _
                 lpCreationTime As Currency, _
                 lpExitTime As Currency, _
                 lpKernelTime As Currency, _
                 lpUserTime As Currency) As Long
    Public Declare Function GetProcessTimes Lib "Kernel32.dll" ( _
                 ByVal hProcess As Long, _
                 lpCreationTime As Currency, _
                 lpExitTime As Currency, _
                 lpKernelTime As Currency, _
                 lpUserTime As Currency) As Long
    
    Private Type RGBTRIPLE
        rgbRed As Byte
        rgbGreen As Byte
        rgbBlue As Byte
    End Type
    
    Private Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
    End Type
    
    Private Type BITMAPINFOHEADER '40 bytes
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
    End Type
    
    Private Type BITMAPINFO256
        bmiHeader As BITMAPINFOHEADER
        bmiColors(0 To 255) As RGBQUAD
    End Type
    
    Private Const BI_RGB = 0&
    Private Const DIB_RGB_COLORS = 0
    Public Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
    Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO256, ByVal wUsage As Long) As Long
    Private Declare Function CreateDIBSection256 Lib "gdi32" Alias "CreateDIBSection" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO256, ByVal un As Long, lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) As Long
    Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long

    Const FW_DONTCARE = 0
    Const FW_THIN = 100
    Const FW_EXTRALIGHT = 200
    Const FW_ULTRALIGHT = 200
    Const FW_LIGHT = 300
    Const FW_NORMAL = 400
    Const FW_REGULAR = 400
    Const FW_MEDIUM = 500
    Const FW_SEMIBOLD = 600
    Const FW_DEMIBOLD = 600
    Const FW_BOLD = 700
    Const FW_EXTRABOLD = 800
    Const FW_ULTRABOLD = 800
    Const FW_HEAVY = 900
    Const FW_BLACK = 900
    Const ANSI_CHARSET = 0
    Const ARABIC_CHARSET = 178
    Const BALTIC_CHARSET = 186
    Const CHINESEBIG5_CHARSET = 136
    Const DEFAULT_CHARSET = 1
    Const EASTEUROPE_CHARSET = 238
    Const GB2312_CHARSET = 134
    Const GREEK_CHARSET = 161
    Const HANGEUL_CHARSET = 129
    Const HEBREW_CHARSET = 177
    Const JOHAB_CHARSET = 130
    Const MAC_CHARSET = 77
    Const OEM_CHARSET = 255
    Const RUSSIAN_CHARSET = 204
    Const SHIFTJIS_CHARSET = 128
    Const SYMBOL_CHARSET = 2
    Const THAI_CHARSET = 222
    Const TURKISH_CHARSET = 162
    Const OUT_DEFAULT_PRECIS = 0
    Const OUT_DEVICE_PRECIS = 5
    Const OUT_OUTLINE_PRECIS = 8
    Const OUT_RASTER_PRECIS = 6
    Const OUT_STRING_PRECIS = 1
    Const OUT_STROKE_PRECIS = 3
    Const OUT_TT_ONLY_PRECIS = 7
    Const OUT_TT_PRECIS = 4
    Const CLIP_DEFAULT_PRECIS = 0
    Const CLIP_EMBEDDED = 128
    Const CLIP_LH_ANGLES = 16
    Const CLIP_STROKE_PRECIS = 2
    Const ANTIALIASED_QUALITY = 4
    Const DEFAULT_QUALITY = 0
    Const DRAFT_QUALITY = 1
    Const NONANTIALIASED_QUALITY = 3
    Const PROOF_QUALITY = 2
    Const DEFAULT_PITCH = 0
    Const FIXED_PITCH = 1
    Const VARIABLE_PITCH = 2
    Const FF_DECORATIVE = 80
    Const FF_DONTCARE = 0
    Const FF_MODERN = 48
    Const FF_ROMAN = 16
    Const FF_SCRIPT = 64
    Const FF_SWISS = 32
    
    Public Const CF_ANSIONLY = &H400
    Public Const CF_APPLY = &H200
    Public Const CF_BOTH = &H3
    Public Const CF_EFFECTS = &H100
    Public Const CF_ENABLEHOOK = &H8
    Public Const CF_ENABLETEMPLATE = &H10
    Public Const CF_ENABLETEMPLATEHANDLE = &H20
    Public Const CF_FIXEDPITCHONLY = &H4000
    Public Const CF_FORCEFONTEXIST = &H10000
    Public Const CF_INITTOLOGFONTSTRUCT = &H40
    Public Const CF_LIMITSIZE = &H2000
    Public Const CF_NOOEMFONTS = &H800
    Public Const CF_NOFACESEL = &H80000
    Public Const CF_NOSCRIPTSEL = &H800000
    Public Const CF_NOSIZESEL = &H200000
    Public Const CF_NOSIMULATIONS = &H1000
    Public Const CF_NOSTYLESEL = &H100000
    Public Const CF_NOVECTORFONTS = &H800
    Public Const CF_NOVERTFONTS = &H1000000
    Public Const CF_PRINTERFONTS = &H2
    Public Const CF_SCALABLEONLY = &H20000
    Public Const CF_SCREENFONTS = &H1
    Public Const CF_SCRIPTSONLY = &H400
    Public Const CF_SELECTSCRIPT = &H400000
    Public Const CF_SHOWHELP = &H4
    Public Const CF_TTONLY = &H40000
    Public Const CF_USESTYLE = &H80
    Public Const CF_WYSIWYG = &H8000
    Public Const BOLD_FONTTYPE = &H100
    Public Const ITALIC_FONTTYPE = &H200
    Public Const PRINTER_FONTTYPE = &H4000
    Public Const REGULAR_FONTTYPE = &H400
    Public Const SCREEN_FONTTYPE = &H2000
    Public Const SIMULATED_FONTTYPE = &H8000
    
    Type CHOOSEFONT_TYPE
      lStructSize As Long
      hwndOwner As Long
      hDC As Long
      lpLogFont As Long
      iPointSize As Long
      Flags As Long
      rgbColors As Long
      lCustData As Long
      lpfnHook As Long
      lpTemplateName As String
      hInstance As Long
      lpszStyle As String
      nFontType As Integer
      MISSING_ALIGNMENT As Integer
      nSizeMin As Long
      nSizeMax As Long
    End Type
    
    Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName As String * 32
    End Type
    Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (lpcf As CHOOSEFONT_TYPE) As Long

    Const GHND = &H40
    Const GMEM_DDESHARE = &H2000
    Const GMEM_DISCARDABLE = &H100
    Const GMEM_FIXED = &H0
    Const GMEM_MOVEABLE = &H2
    Const GMEM_NOCOMPACT = &H10
    Const GMEM_NODISCARD = &H20
    Const GMEM_SHARE = &H2000
    Const GMEM_ZEROINIT = &H40
    Const GPTR = &H42
    
    '
    ' GIF Stuff
    '
    Private Type GifScreenDescriptor
        logical_screen_width As Integer
        logical_screen_height As Integer
        Flags As Byte
        background_color_index As Byte
        pixel_aspect_ratio As Byte
    End Type
    
    Private Type GifImageDescriptor
        Left As Integer
        Top As Integer
        Width As Integer
        Height As Integer
        Format As Byte 'ImageFormat
    End Type
    '========Added by Wolfgang Goetz for transparent GIFs=====
    Private Type CONTROLBLOCK '(April 8., 2002 --> Wolfgang Goetz)
        Blocksize As Byte
        Flags As Byte
        Delay As Integer
        TransParent_Color As Byte
        Terminator As Byte
    End Type
    Private Const GIF89a = "GIF89a"
    Private Const CTRLINTRO As Byte = &H21
    Private Const CTRLLABEL As Byte = &HF9

    Const GIF87A = "GIF87a"
    
    Const GIFTERMINATOR As Byte = &H3B
    Const IMAGESEPARATOR As Byte = &H2C
    Const CHAR_BIT = 8
    Const CODESIZE As Byte = 9
    Const CLEARCODE = 256
    Const ENDCODE  As Integer = 257
    Const FIRSTCODE = 258
    Const LASTCODE As Integer = 511
    Const MAX_CODE = LASTCODE - FIRSTCODE

    Public Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (lpMutexAttributes&, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
    Public Declare Function OpenMutex Lib "kernel32" Alias "OpenMutexA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
    Public Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As Long

    Public Type PROCESS_MEMORY_COUNTERS
        cb As Long                           '// Size of the structure, in bytes.
        PageFaultCount As Long               '// Number of page faults.
        PeakWorkingSetSize As Long           '// Peak working set size.
        WorkingSetSize As Long               '// Current working set size.
        QuotaPeakPagedPoolUsage As Long      '// Peak paged pool usage.
        QuotaPagedPoolUsage As Long          '// Current paged pool usage.
        QuotaPeakNonPagedPoolUsage As Long   '// Peak nonpaged pool usage.
        QuotaNonPagedPoolUsage As Long       '// Current nonpaged pool usage.
        PagefileUsage As Long                '// Current space allocated for the pagefile.
                                             '// Those pages may or may not be in memory.
        PeakPagefileUsage As Long            '// Peak space allocated for the pagefile.
    End Type

    Private Declare Function GetProcessMemoryInfo Lib _
         "PSAPI.DLL" (ByVal lHandle As Long, lpStructure As _
         PROCESS_MEMORY_COUNTERS, ByVal lSize As Long) As Integer
    Public Declare Function GetCurrentProcess Lib "kernel32" () As Long


    ' Data type to hold version information.
    Public Type VersionInformationType
        StructureVersion As String
        FileVersion As String
        ProductVersion As String
        FileFlags As String
        TargetOperatingSystem As String
        FileType As String
        FileSubtype As String
    End Type
    
    ' API declarations.
    Private Type VS_FIXEDFILEINFO
        dwSignature As Long
        dwStrucVersionl As Integer     '  e.g. = &h0000 = 0
        dwStrucVersionh As Integer     '  e.g. = &h0042 = .42
        dwFileVersionMSl As Integer    '  e.g. = &h0003 = 3
        dwFileVersionMSh As Integer    '  e.g. = &h0075 = .75
        dwFileVersionLSl As Integer    '  e.g. = &h0000 = 0
        dwFileVersionLSh As Integer    '  e.g. = &h0031 = .31
        dwProductVersionMSl As Integer '  e.g. = &h0003 = 3
        dwProductVersionMSh As Integer '  e.g. = &h0010 = .1
        dwProductVersionLSl As Integer '  e.g. = &h0000 = 0
        dwProductVersionLSh As Integer '  e.g. = &h0031 = .31
        dwFileFlagsMask As Long        '  = &h3F for version "0.42"
        dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
        dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
        dwFileType As Long             '  e.g. VFT_DRIVER
        dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
        dwFileDateMS As Long           '  e.g. 0
        dwFileDateLS As Long           '  e.g. 0
    End Type
    
    Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
    Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
    Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
    Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal length As Long)
    
    ' dwFileFlags
    Private Const VS_FFI_SIGNATURE = &HFEEF04BD
    Private Const VS_FFI_structure_versionSION = &H10000
    Private Const VS_FFI_file_flagsMASK = &H3F&
    
    ' dwFileFlags
    Private Const VS_FF_DEBUG = &H1
    Private Const VS_FF_PRERELEASE = &H2
    Private Const VS_FF_PATCHED = &H4
    Private Const VS_FF_PRIVATEBUILD = &H8
    Private Const VS_FF_INFOINFERRED = &H10
    Private Const VS_FF_SPECIALBUILD = &H20
    
    ' dwFileOS
    Private Const VOS_UNKNOWN = &H0
    Private Const VOS_DOS = &H10000
    Private Const VOS_OS216 = &H20000
    Private Const VOS_OS232 = &H30000
    Private Const VOS_NT = &H40000
    Private Const VOS_DOS_WINDOWS16 = &H10001
    Private Const VOS_DOS_WINDOWS32 = &H10004
    Private Const VOS_OS216_PM16 = &H20002
    Private Const VOS_OS232_PM32 = &H30003
    Private Const VOS_NT_WINDOWS32 = &H40004
    
    ' dwFileType
    Private Const VFT_UNKNOWN = &H0
    Private Const VFT_APP = &H1
    Private Const VFT_DLL = &H2
    Private Const VFT_DRV = &H3
    Private Const VFT_FONT = &H4
    Private Const VFT_VXD = &H5
    Private Const VFT_STATIC_LIB = &H7
    
    ' dwFileSubtype for drivers
    Private Const VFT2_UNKNOWN = &H0
    Private Const VFT2_DRV_PRINTER = &H1
    Private Const VFT2_DRV_KEYBOARD = &H2
    Private Const VFT2_DRV_LANGUAGE = &H3
    Private Const VFT2_DRV_DISPLAY = &H4
    Private Const VFT2_DRV_MOUSE = &H5
    Private Const VFT2_DRV_NETWORK = &H6
    Private Const VFT2_DRV_SYSTEM = &H7
    Private Const VFT2_DRV_INSTALLABLE = &H8
    Private Const VFT2_DRV_SOUND = &H9
    Private Const VFT2_DRV_COMM = &HA

    Public Enum FileVersionTypes
        VersionStructure = 0
        VersionFile
        VersionProduct
    End Enum

    Private Const MAXPATH = 260
    
    Public Type FileInCabinetInfo
        NameInCabinet As Long
        FileSize      As Long
        Win32Error    As Long
        DosDate       As Integer
        DosTime       As Integer
        DosAttribs    As Integer
        FullTargetName(1 To MAXPATH) As Byte
    End Type

    Private Declare Function SetupIterateCabinet Lib "setupapi.dll" _
            Alias "SetupIterateCabinetA" (ByVal CabinetFile As String, _
            ByVal Reserved As Long, ByVal MsgHandler As Long, _
            ByVal Context As Long) As Long
    Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long

    Private Type FILEPATHS
        Target     As Long
        Source     As Long
        Win32Error As Integer
        Flags      As Long
    End Type

    '
    ' Local enum, indicating what action to
    ' take on each pass through the callback
    ' procedure.
    '
    Private Enum SetupIterateCabinetActions
        sicCount
        sicReport
        sicExtract
        sicGetXML
    End Enum
    
    '
    ' Notification messages, handled in the callback
    ' procedure. This class doesn't handle them all.
    '
    Private Const SPFILENOTIFY_FILEINCABINET = &H11
    Private Const SPFILENOTIFY_NEEDNEWCABINET = &H12
    Private Const SPFILENOTIFY_FILEEXTRACTED = &H13
    
    '
    ' Instructions sent out of the callback procedure.
    ' Tells Windows what to do next.
    '
    Private Enum FILEOP
        FILEOP_ABORT = 0
        FILEOP_DOIT = 1
        FILEOP_SKIP = 2
    End Enum

    '
    ' Used for the CAB extraction
    '
    Dim msFileToExtract$
    Dim msOutputFile$
    
    Public Const ABM_NEW = &H0
    Public Const ABM_REMOVE = &H1
    Public Const ABM_QUERYPOS = &H2
    Public Const ABM_SETPOS = &H3
    Public Const ABM_GETSTATE = &H4
    Public Const ABM_GETTASKBARPOS = &H5
    Public Const ABM_ACTIVATE = &H6
    Public Const ABM_GETAUTOHIDEBAR = &H7
    Public Const ABM_SETAUTOHIDEBAR = &H8
    Public Const ABM_WINDOWPOSCHANGED = &H9
    
    Public Const ABN_STATECHANGE = &H0
    Public Const ABN_POSCHANGED = &H1
    Public Const ABN_FULLSCREENAPP = &H2
    Public Const ABN_WINDOWARRANGE = &H3
    
    Public Const ABS_AUTOHIDE = &H1
    Public Const ABS_ALWAYSONTOP = &H2
    
    Public Const ABE_LEFT = 0
    Public Const ABE_TOP = 1
    Public Const ABE_RIGHT = 2
    Public Const ABE_BOTTOM = 3
    
    Public Type APPBARDATA
      cbSize As Long
      hWnd As Long
      uCallbackMessage As Long
      uEdge As Long
      rc As RECT
      lParam As Long      'message specific
    End Type
    
    Public Declare Function SHAppBarMessage Lib "shell32" (ByVal dwMessage As Long, pData As APPBARDATA) As Long
    
    
    Private Type LUID
  lowpart As Long
  highpart As Long
End Type

Private Type TOKEN_PRIVILEGES
   PrivilegeCount As Long
   LuidUDT As LUID
   Attributes As Long
End Type

Const TOKEN_ADJUST_PRIVILEGES = &H20
Const TOKEN_QUERY = &H8
Const SE_PRIVILEGE_ENABLED = &H2
    
Private Declare Function GetVersion _
   Lib "kernel32" () As Long
Private Declare Function OpenProcessToken _
   Lib "advapi32" (ByVal ProcessHandle As Long, _
   ByVal DesiredAccess As Long, _
   TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue _
   Lib "advapi32" Alias "LookupPrivilegeValueA" _
   (ByVal lpSystemName As String, _
   ByVal lpName As String, _
   lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges _
   Lib "advapi32" (ByVal TokenHandle As Long, _
   ByVal DisableAllPrivileges As Long, _
   NewState As TOKEN_PRIVILEGES, _
   ByVal BufferLength As Long, _
   PreviousState As Any, _
   ReturnLength As Any) As Long
   
Public Const SM_CXVIRTUALSCREEN = 78
Public Const SM_CYVIRTUALSCREEN = 79
Public Const SM_CMONITORS = 80
Public Const SM_SAMEDISPLAYFORMAT = 81
   

Public Function Z_CabinetCallback(ByVal Context As Long, ByVal Notification As Long, ByRef Param1 As FileInCabinetInfo, ByVal Param2 As Long) As Long
    
Dim fp As FILEPATHS
Dim lLen&
Dim sBuffer$
Dim bytTemp() As Byte
Dim i%
    
    '
    ' Callback procedure for SetupIterateCabinet
    ' Handle the callback for the CAB file.
    '
    On Error Resume Next
    Z_CabinetCallback = NO_ERROR
    Select Case Notification
        Case SPFILENOTIFY_NEEDNEWCABINET
            
        Case SPFILENOTIFY_FILEEXTRACTED
            '
            ' Copy the bytes passed into a FILEPATHS structure.
            ' Although this procedure gets a parameter of
            ' type FileCabinetInfo, you want to cast it as a
            ' FILEPATHS structure. The LSET statement does that
            ' for you. You can also use the CopyMemory API function,
            ' but this is simpler.
            '
            LSet fp = Param1
            Z_CabinetCallback = fp.Win32Error
        
        Case SPFILENOTIFY_FILEINCABINET
            '
            ' Given a string pointer, copy the value
            ' of the string into a new, safe location.
            '
            lLen = lstrlen(Param1.NameInCabinet)
            sBuffer = Space(lLen)
            Call CopyMemory(ByVal sBuffer, ByVal Param1.NameInCabinet, lLen)
            If StrComp(sBuffer, msFileToExtract, vbTextCompare) = 0 Then
                '
                ' Copy the byte array to the output location.
                '
                bytTemp = StrConv(msOutputFile, vbFromUnicode)
                For i = LBound(bytTemp) To UBound(bytTemp)
                    Param1.FullTargetName(i + 1) = bytTemp(i)
                Next i
                Param1.FullTargetName(i + 1) = 0
                Z_CabinetCallback = FILEOP_DOIT
            Else
                Z_CabinetCallback = FILEOP_SKIP
            End If
    End Select

End Function



Public Function PSGEN_ExtractFileFromCabinet(ByVal sSourceFilename$, ByVal sCabinetFilename$, ByVal sDestinationFilename$) As Boolean
Attribute PSGEN_ExtractFileFromCabinet.VB_Description = "Extracts the file from the given cabinet and places it in the destination file"
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2005
'
'****************************************************************************
'
'                     NAME: Function PSGEN_ExtractFileFromCabinet
'
'                     sSourceFilename$   - Name of the file to extract
'                     sCabinetFilename$  - Filename of the CAB
'
'                          ) As Boolean
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    27 October 2005   First created for MediaWeb
'
'                  PURPOSE: Extracts the file from the given cabinet and
'                           places it in the destination file.
'                           Returns true if it worked OK
'
'****************************************************************************
'
'
Dim bReturn As Boolean
Dim lReturn&

    '
    ' Extracts the file(s) from the cabinet. FileToExtract can specify
    ' the file to extract or if ommitted all files will be extracted.
    ' OutputPath can specify the folder to extract to. The default is the
    ' same folder as the cab file. When extracting a single file,
    ' OutputFile can specify the extract file name. The default is the
    ' original file name.
    '
    On Error Resume Next
    If PSGEN_FileExists(sCabinetFilename) Then
        
        '
        ' Set up the module-level variables
        ' tracking which file(s) you want to extract,
        ' and where you want to put them.
        '
        msFileToExtract = sSourceFilename
        msOutputFile = sDestinationFilename
        lReturn = SetupIterateCabinet(sCabinetFilename, 0, AddressOf Z_CabinetCallback, sicExtract)
        
        '
        ' If the return value is 0, the call to SetupIterateCabinet failed.
        '
        bReturn = (lReturn <> 0)
    End If

    '
    ' Return value to caller
    '
    PSGEN_ExtractFileFromCabinet = bReturn

End Function


Public Function PSGEN_GetFileVersion$(ByVal sFilename$, ByVal eType As FileVersionTypes)
Attribute PSGEN_GetFileVersion.VB_Description = "Returns a dot separated version string for the filename"
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2005
'
'****************************************************************************
'
'                     NAME: Function PSGEN_GetFileVersion
'
'                     sFilename$         - Full path filename to get version info.
'
'                          ) As String
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    27 October 2005   First created for MediaWeb
'
'                  PURPOSE: Returns a dot separated version string for the
'                           filename
'
'****************************************************************************
'
'
Dim lHandle&, lInfoSize&, lFixedInfoSize&, lInfoAddress&
Dim abBuffer() As Byte
Dim stFileVer As VS_FIXEDFILEINFO
Dim stVersionInfo As VersionInformationType

Dim sReturn$


    '
    ' Get the version information buffer size
    '
    On Error Resume Next
    lInfoSize = GetFileVersionInfoSize(sFilename, lHandle)
    If lInfoSize > 0 Then

        '
        ' Load the fixed file information into a buffer
        '
        ReDim abBuffer(1 To lInfoSize)
        If GetFileVersionInfo(sFilename, 0&, lInfoSize, abBuffer(1)) <> 0 Then
            If VerQueryValue(abBuffer(1), "\", lInfoAddress, lFixedInfoSize) <> 0 Then
        
                '
                ' Copy the information from the buffer into a usable structure
                '
                If lInfoAddress <> 0 Then
                    MoveMemory stFileVer, lInfoAddress, Len(stFileVer)
                
                    '
                    ' Get the version information
                    '
                    With stFileVer
                    If eType = VersionStructure Then
                        sReturn = Format$(.dwStrucVersionh) & "." & Format$(.dwStrucVersionl)
            
                    ElseIf eType = VersionFile Then
                        sReturn = Format$(.dwFileVersionMSh) & "." & Format$(.dwFileVersionMSl) & "." & Format$(.dwFileVersionLSh) & "." & Format$(.dwFileVersionLSl)
            
                    ElseIf eType = VersionProduct Then
                        sReturn = Format$(.dwProductVersionMSh) & "." & Format$(.dwProductVersionMSl) & "." & Format$(.dwProductVersionLSh) & "." & Format$(.dwProductVersionLSl)
                    End If
                    End With
                End If
            End If
        End If
    End If

    '
    ' Return value to caller
    '
    PSGEN_GetFileVersion = sReturn

End Function

    

Public Sub PSGEN_ShowInTaskBar(ByVal lHwnd&, ByVal bShow As Boolean)
Attribute PSGEN_ShowInTaskBar.VB_Description = "Shows or hides a form from the toolbar"
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2005
'
'****************************************************************************
'
'                     NAME: Sub PSGEN_ShowInTaskBar
'
'                     lHwnd&                       - Handle of the window to show in taskbar
'                     bShow As Boolean             - True if the app should show in toolbar
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    07 October 2005   First created for PivotalStock
'
'                  PURPOSE: Shows or hides a form from the toolbar
'
'****************************************************************************
'
'
Const GWL_EXSTYLE = (-20)
Const WS_EX_APPWINDOW = &H40000

Dim lStyle As Long

    '
    ' We have to hide the window first while we make the change
    '
    On Error Resume Next
    Call ShowWindow(lHwnd, SW_HIDE)
    
    '
    ' Get the current style and modify it
    '
    lStyle = GetWindowLong(lHwnd, GWL_EXSTYLE)
    If bShow Then
        lStyle = lStyle Or WS_EX_APPWINDOW
    Else
        If lStyle And WS_EX_APPWINDOW Then lStyle = lStyle - WS_EX_APPWINDOW
    End If
    
    '
    ' Set the style
    '
    Call SetWindowLong(lHwnd, GWL_EXSTYLE, lStyle)
    App.TaskVisible = bShow
    
    '
    ' Show the form
    '
    Call ShowWindow(lHwnd, SW_NORMAL)

End Sub

Public Function PSGEN_RemoveDuplicates$(ByVal sValue$, ByVal sSep$)
Attribute PSGEN_RemoveDuplicates.VB_Description = "Returns a string where duplicate values have been removed\r\nValues are seperated by sSep"
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2005
'
'****************************************************************************
'
'                     NAME: Function PSGEN_RemoveDuplicates
'
'                     sValue$            - String of sSep seperated values
'                     sSep$              - Seperator string
'
'                          ) As String
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    29 July 2005   First created for MediaWeb
'
'                  PURPOSE: Returns a string where duplicate values have been
'                           removed
'                           Values are seperated by sSep
'
'****************************************************************************
'
'
Dim sReturn$
Dim asValues$()
Dim iCnt%

    '
    ' Loop through the list of values
    '
    On Error Resume Next
    asValues = Split(sValue, sSep)
    For iCnt = 0 To UBound(asValues)
        
        '
        ' If the value is in the list then ignore it
        '
        If InStr(1, sSep + sReturn + sSep, asValues(iCnt), vbTextCompare) = 0 Then
            sReturn = sReturn + IIf(sReturn = "", "", sSep) + asValues(iCnt)
        End If
    Next iCnt

    '
    ' Return value to caller
    '
    PSGEN_RemoveDuplicates = sReturn

End Function


Public Function PSGEN_GetProcessMemoryCount&()
Attribute PSGEN_GetProcessMemoryCount.VB_Description = "Returns the number of bytes used by the current process"
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2005
'
'****************************************************************************
'
'                     NAME: Function PSGEN_GetProcessMemoryCount
'
'                          ) As Long
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    13 March 2005   First created for PivotalBASISSlave
'
'                  PURPOSE: Returns the number of bytes used by the current
'                           process
'
'****************************************************************************
'
'
Dim stMemory As PROCESS_MEMORY_COUNTERS

    '
    ' Call the API
    '
    On Error Resume Next
    Call GetProcessMemoryInfo(GetCurrentProcess, stMemory, Len(stMemory))
    
    '
    ' Return value to caller
    '
    PSGEN_GetProcessMemoryCount = stMemory.WorkingSetSize

End Function



Public Function PSGEN_WorkDays&(ByVal dStart As Date, ByVal dEnd As Date)
Attribute PSGEN_WorkDays.VB_Description = "Returns the number of work days between the two dates"
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2004
'
'****************************************************************************
'
'                     NAME: Function PSGEN_WorkDays
'
'                     dStart As Date            - Start date
'                     dEnd As Date              - End date
'
'                          ) As Long
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    27 July 2004   First created for PSMediaServices
'
'                  PURPOSE: Returns the number of work days between the two
'                           dates
'
'****************************************************************************
'
'
Dim lReturn&, lWholeWeeks&
Dim dDateCnt As Date


    '
    ' Determine how many weeks there are between the dates
    '
    On Error Resume Next
    lWholeWeeks = DateDiff("w", dStart, dEnd)
    dDateCnt = DateAdd("ww", lWholeWeeks, dStart)
    lReturn = 0
    Do While dDateCnt < dEnd
       If LCase(Format(dDateCnt, "ddd")) <> "sun" And LCase(Format(dDateCnt, "ddd")) <> "sat" Then lReturn = lReturn + 1
       dDateCnt = DateAdd("d", 1, dDateCnt)
    Loop
    lReturn = lWholeWeeks * 5 + lReturn

    '
    ' Return value to caller
    '
    PSGEN_WorkDays = lReturn

End Function




Public Function PSGEN_StartMutex&(ByVal sName$)
Attribute PSGEN_StartMutex.VB_Description = "Creates and waits on the named mutex returning true if the mutex is owned by the caller"
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2004
'
'****************************************************************************
'
'                     NAME: Function PSGEN_StartMutex
'
'                     sName$             - Name of the mutex
'
'                          ) As Boolean
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    24 July 2004   First created for MediaWeb
'
'                  PURPOSE: Creates and waits on the named mutex returning
'                           the mutex handle if the mutex is owned by the caller
'
'****************************************************************************
'
'
Const MUTEX_TIMEOUT = 5000
Dim lReturn&

    '
    ' Create the mutex and then wait on it
    '
    On Error Resume Next
    sName = Replace(sName, "\", "")
    sName = Replace(sName, ".", "")
    sName = Replace(sName, ":", "")
    sName = Replace(sName, " ", "")
    lReturn = CreateMutex(0, True, sName)
    If lReturn <> 0 Then

        '
        ' Now wait on the mutex for a period
        '
        If WaitForSingleObject(lReturn, MUTEX_TIMEOUT) = WAIT_TIMEOUT Then
            
            '
            ' Couldn't lock the mutex so release our handle anyway
            '
            Call ReleaseMutex(lReturn)
            lReturn = 0
            App.LogEvent "MutEx failed on file " + sName
        End If
    End If
    
    '
    ' Return value to caller
    '
    PSGEN_StartMutex = lReturn

End Function


Public Sub PSGEN_EndMutex(ByVal lMutex&)
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2004
'
'****************************************************************************
'
'                     NAME: Function PSGEN_EndMutex
'
'                     lMutex&             - Handle of the mutex
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    24 July 2004   First created for MediaWeb
'
'                  PURPOSE: Releases the mutex
'
'****************************************************************************
'
'

    '
    ' Close the mute if it is valid
    '
    On Error Resume Next
    Call ReleaseMutex(lMutex)
    Call CloseHandle(lMutex)

End Sub




Public Function PSGEN_ChooseFont(ByVal frmOwner As Form, ByVal objDefault As Font, Optional ByVal lFlags) As Font
Attribute PSGEN_ChooseFont.VB_Description = "Returns a font from the user"
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2004
'
'****************************************************************************
'
'                     NAME: Function PSGEN_ChooseFont
'
'                     objDefault As Font        - Default font to use
'
'                          ) As Font
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    25 May 2004   First created for PivotalDesktop
'
'                  PURPOSE: Returns a font from the user
'
'****************************************************************************
'
'
Dim stFontStructure As CHOOSEFONT_TYPE
Dim stFont As LOGFONT
Dim hMem&, pMem&, lStatus&
Dim sFontName$
Dim objReturn As New StdFont


    '
    ' Initialize the default selected font: Times New Roman, regular, black, 12 point.
    ' (Note that some of that information is in the CHOOSEFONT_TYPE structure instead.)
    '
    On Error Resume Next
    stFont.lfHeight = (objDefault.size / 72) * frmOwner.ScaleY(1, vbInches, vbPixels) ' determine default height
    stFont.lfWidth = 0  ' determine default width
    stFont.lfEscapement = 0  ' angle between baseline and escapement vector
    stFont.lfOrientation = 0  ' angle between baseline and orientation vector
    stFont.lfWeight = IIf(objDefault.Bold, FW_BOLD, FW_NORMAL)
    stFont.lfItalic = IIf(objDefault.Italic, 1, 0)
    stFont.lfUnderline = IIf(objDefault.Underline, 1, 0)
    stFont.lfStrikeOut = IIf(objDefault.Strikethrough, 1, 0)
    stFont.lfCharSet = objDefault.Charset  ' use default character set
    stFont.lfOutPrecision = OUT_DEFAULT_PRECIS  ' default precision mapping
    stFont.lfClipPrecision = CLIP_DEFAULT_PRECIS  ' default clipping precision
    stFont.lfQuality = DEFAULT_QUALITY  ' default quality setting
    stFont.lfPitchAndFamily = DEFAULT_PITCH Or FF_DONTCARE  ' default pitch, proportional with serifs
    stFont.lfFaceName = objDefault.Name & vbNullChar  ' string must be null-terminated
    
    '
    ' Create the memory block which will act as the LOGFONT structure buffer
    '
    hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(stFont))
    pMem = GlobalLock(hMem)  ' lock and get pointer
    CopyMemory ByVal pMem, stFont, Len(stFont)  ' copy structure's contents into block
    
    '
    ' Initialize dialog box: Screen and printer fonts, point size between 10 and 72
    '
    stFontStructure.lStructSize = Len(stFontStructure)  ' size of structure
    stFontStructure.hwndOwner = frmOwner.hWnd  ' window Form1 is opening this dialog box
    stFontStructure.hDC = Printer.hDC  ' device context of default printer (using VB's mechanism)
    stFontStructure.lpLogFont = pMem  ' pointer to LOGFONT memory block buffer
    stFontStructure.iPointSize = Int(objDefault.size) * 10 ' 12 point font (in units of 1/10 point)
    If IsMissing(lFlags) Then
        stFontStructure.Flags = CF_PRINTERFONTS Or CF_EFFECTS Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT Or CF_LIMITSIZE Or CF_NOSCRIPTSEL
    Else
        stFontStructure.Flags = lFlags
    End If
    stFontStructure.rgbColors = RGB(0, 0, 0)  ' black
    stFontStructure.lCustData = 0  ' we don't use this here...
    stFontStructure.lpfnHook = 0  ' ...or this...
    stFontStructure.lpTemplateName = ""  ' ...or this...
    stFontStructure.hInstance = 0  ' ...or this...
    stFontStructure.lpszStyle = ""  ' ...or this
    stFontStructure.nFontType = REGULAR_FONTTYPE  ' regular font type i.e. not bold or anything
    stFontStructure.nSizeMin = 8  ' minimum point size
    stFontStructure.nSizeMax = 72  ' maximum point size
    
    '
    ' Now, call the function.  If successful, copy the LOGFONT structure back into the structure
    ' and then print out the attributes we mentioned earlier that the user selected.
    '
    lStatus = ChooseFont(stFontStructure)  ' open the dialog box
    objReturn.Charset = objDefault.Charset
    objReturn.Weight = objDefault.Weight
    objReturn.Name = objDefault.Name
    objReturn.size = objDefault.size
    objReturn.Bold = objDefault.Bold
    objReturn.Italic = objDefault.Italic
    objReturn.Underline = objDefault.Underline
    objReturn.Strikethrough = objDefault.Strikethrough
    
    If lStatus <> 0 Then  ' success
        CopyMemory stFont, ByVal pMem, Len(stFont)  ' copy memory back
        DoEvents
        objReturn.Name = Left(stFont.lfFaceName, InStr(stFont.lfFaceName, vbNullChar) - 1)
        objReturn.size = stFontStructure.iPointSize / 10
        objReturn.Bold = stFont.lfWeight >= FW_BOLD
        objReturn.Italic = stFont.lfItalic <> 0
        objReturn.Underline = stFont.lfUnderline <> 0
        objReturn.Strikethrough = stFont.lfStrikeOut <> 0
    End If
    
    '
    ' Deallocate the memory block we created earlier.  Note that this must
    ' be done whether the function succeeded or not.
    '
    lStatus = GlobalUnlock(hMem)  ' destroy pointer, unlock block
    lStatus = GlobalFree(hMem)  ' free the allocated memory

    '
    ' Return value to caller
    '
    Set PSGEN_ChooseFont = objReturn

End Function

    

Public Function PSGEN_CreateMimeTypeIniFile(ByVal sFilename$) As Boolean
Attribute PSGEN_CreateMimeTypeIniFile.VB_Description = "Creates a mimetype file by reading the IIS default list"
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2003
'
'****************************************************************************
'
'                     NAME: Sub PSGEN_CreateMimeTypeIniFile
'
'                     sFilename$         - Filename to use
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    12 November 2003   First created for PivotalBASISSlave
'
'                  PURPOSE: Creates a mimetype file by reading the IIS
'                           default list
'
'****************************************************************************
'
'
Dim bReturn As Boolean
Dim objMimeMap As Object
Dim vMimeMap As Variant
Dim iCnt

    '
    ' Create the external objects
    '
    On Error Resume Next
    Set objMimeMap = GetObject("IIS://localhost/mimemap")
    If Err = 0 Then
    
        '
        ' Get the array of mimetypes
        '
        vMimeMap = objMimeMap.Get("MimeMap")
        If IsArray(vMimeMap) Then
        
            '
            ' Start the file off
            '
            If PSGEN_WriteTextFile("[" + PSGEN_MIMETYPE_SECTION + "]" + vbCrLf, sFilename) Then
            
                '
                ' Output the mimetypes
                '
                For iCnt = LBound(vMimeMap) To UBound(vMimeMap)
                    Call PSGEN_WriteTextFile(Mid$(vMimeMap(iCnt).Extension, 2) + "=" + vMimeMap(iCnt).MIMEType + vbCrLf, sFilename, , True)
                Next iCnt
                
                '
                ' Now put out the file types
                '
                Call PSGEN_WriteTextFile("[" + PSGEN_MIMETYPE_SECTION_REVERSE + "]" + vbCrLf, sFilename, , True)
                For iCnt = LBound(vMimeMap) To UBound(vMimeMap)
                    Call PSGEN_WriteTextFile(vMimeMap(iCnt).MIMEType + "=" + Mid$(vMimeMap(iCnt).Extension, 2) + vbCrLf, sFilename, , True)
                Next iCnt
                bReturn = True
            End If
        End If
        Set objMimeMap = Nothing
    End If
    
    PSGEN_CreateMimeTypeIniFile = bReturn
    
End Function



Public Sub PSGEN_LaunchBrowser(ByVal sURL$)
Attribute PSGEN_LaunchBrowser.VB_Description = "Launches the default browser using the URL"
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2003
'
'****************************************************************************
'
'                     NAME: Sub PSGEN_LaunchBrowser
'
'                     sURL$              - URL to launch
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    10 June 2003   First created for DialogLoadManager
'
'                  PURPOSE: Launches the default browser using the URL
'
'****************************************************************************
'
'
Dim sBrowserExec As String * 255
Dim sTmp$

    '
    ' Get the default browser app
    '
    On Error Resume Next
    sBrowserExec = Space(Len(sBrowserExec))
    sTmp = PSGEN_GetTempPathFilename("htm")
    Call FindExecutable(sTmp, vbNullString, sBrowserExec)
    sBrowserExec = PSGEN_GetItem(1, vbNullChar, sBrowserExec)
    Kill sTmp
    
    '
    ' Launch it with the URL as the argument
    '
    Call ShellExecute(0&, "open", sBrowserExec, sURL, "", SW_SHOW)

End Sub



Public Function PSGEN_GetHTMLFromClipboard$(ByVal lWindow&)
Attribute PSGEN_GetHTMLFromClipboard.VB_Description = "Returns the HTML data from the clipboard"
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2003
'
'****************************************************************************
'
'                     NAME: Function PSGEN_GetHTMLFromClipboard
'
'                     lWindow&           - Handle of window to own the clipboard
'
'                          ) As String
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    03 February 2003   First created for PivotalDesktop
'
'                  PURPOSE: Returns the HTML data from the clipboard
'
'****************************************************************************
'
'
Dim abData() As Byte
Dim lMem&, lSize&, lPtr&, lFormat&
Dim sReturn$
    
    '
    ' Get the HTML custom format ID
    '
    On Error Resume Next
    lFormat = RegisterClipboardFormat("HTML Format" + vbNullChar)
    If lFormat > &HC000& Then
    
        '
        ' Open the clipboard for access
        '
        If OpenClipboard(lWindow) Then
            
            '
            ' Check if this data format is available
            '
            If (IsClipboardFormatAvailable(lFormat) <> 0) Then
            
                '
                ' Get the memory handle to the data
                '
                lMem = GetClipboardData(lFormat)
                If (lMem <> 0) Then
                    
                    '
                    ' Get the size of this memory block
                    '
                    lSize = GlobalSize(lMem)
                    If (lSize > 0) Then
                        
                        '
                        ' Get a pointer to the memory
                        '
                        lPtr = GlobalLock(lMem)
                        If (lPtr <> 0) Then
                            
                            '
                            ' Resize the byte array to hold the data
                            '
                            ReDim abData(0 To lSize - 1) As Byte
                            
                            '
                            ' Copy from the pointer into the array
                            '
                            CopyMemory abData(0), ByVal lPtr, lSize
                            
                            '
                            ' Unlock the memory block
                            '
                            GlobalUnlock lMem
                            
                            '
                            ' Now return the data as a string
                            '
                            sReturn = StrConv(abData, vbUnicode)
                        End If
                    End If
                End If
            End If
            CloseClipboard
        End If
    End If

    '
    ' Return value to caller
    '
    PSGEN_GetHTMLFromClipboard = sReturn

End Function

Public Function PSGEN_GetSpecialFolderLocation$(ByVal lFolderID&)
Attribute PSGEN_GetSpecialFolderLocation.VB_Description = "Returns the directory of the specified special folder"
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2002
'
'****************************************************************************
'
'                     NAME: Function PSGEN_GetSpecialFolderLocation
'
'                     lFolderID&         - Special CSIDL folder identity
'
'                          ) As String
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    19 November 2002   First created for PivotalScan2PDF
'
'                  PURPOSE: Returns the directory of the specified special
'                           folder
'
'****************************************************************************
'
'
Dim lRet&
Dim sLocation$, sReturn$
Dim lPID&


    '
    ' Retrieve a PIDL for the specified location
    '
    On Error Resume Next
    If SHGetSpecialFolderLocationLong(0&, lFolderID, lPID) = 0 Then
        
        '
        ' Convert the pidl to a physical path
        '
        sLocation = Space$(MAX_PATH)
        If SHGetPathFromIDList(lPID, sLocation) <> 0 Then
            
            '
            ' If successful, return the location
            '
            sReturn = Left$(sLocation, InStr(sLocation, vbNullChar) - 1)
        End If
        Call CoTaskMemFree(lPID)
    End If

    '
    ' Return value to caller
    '
    PSGEN_GetSpecialFolderLocation = sReturn

End Function



Public Function PSGEN_ConvertRtfToHtml$(ByVal sRTF$)
Attribute PSGEN_ConvertRtfToHtml.VB_Description = "Converts the bit segment of RTF into HTML and returns it"
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2002
'
'****************************************************************************
'
'                     NAME: Function PSGEN_ConvertRtfToHtml
'
'                     sRTF$              - RTF segment to convert
'
'                          ) As String
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    04 October 2002   First created for TeamPlayer
'
'                  PURPOSE: Converts the bit segment of RTF into HTML and
'                           returns it
'
'****************************************************************************
'
'
Dim sReturn$, sTmpRTF$, sTmpHTML$


    '
    ' Create the temporary files
    '
    On Error Resume Next
    sTmpRTF = PSGEN_GetTempPathFilename
    sTmpHTML = PSGEN_GetTempPathFilename
    Kill sTmpHTML
    
    '
    ' Put the text into the RTF
    '
    Call PSGEN_WriteTextFile(sRTF, sTmpRTF)
    
    '
    ' Run the conversion
    '
    If ConvertRtfToHTML(ByVal sTmpRTF, ByVal sTmpHTML, EXO_RESULTS + EXO_INLINECSS + EXO_WMF2GIF + EXO_HTML, 0&, 0&, 96) = 1 Then
        Call PSGEN_ReadTextFile(sTmpHTML, sReturn)
    End If
    
    '
    ' Remove the temporary files
    '
    Kill sTmpRTF
    Kill sTmpHTML

    '
    ' Return value to caller
    '
    PSGEN_ConvertRtfToHtml = sReturn

End Function





Public Function PSGEN_GetHTMLColor$(ByVal lColor#)
Attribute PSGEN_GetHTMLColor.VB_Description = "Returns the HTML version of a windows colour"
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2002
'
'****************************************************************************
'
'                     NAME: Function PSGEN_GetHTMLColor
'
'                     lColor&            - Windows colour to convert
'
'                          ) As String
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    14 May 2002   First created for TeamPlayer
'
'                  PURPOSE: Returns the HTML version of a windows colour
'
'****************************************************************************
'
'
Dim sReturn$, sTmp$

    '
    ' Convert the color to a non-system version
    '
    On Error Resume Next
    lColor = PSGEN_GetAPIColor(lColor)
    
    '
    ' Now re-arrange the bits into a proper HTML colour
    '
    sTmp = Replace(Format$(Hex(lColor), "@@@@@@"), " ", "0")
    sReturn = Mid$(sTmp, 5, 2) + Mid$(sTmp, 3, 2) + Mid$(sTmp, 1, 2)

    '
    ' Return value to caller
    '
    PSGEN_GetHTMLColor = sReturn

End Function

Public Sub PSGEN_SetDefaultPrinter(ByVal sPrinter$)
Attribute PSGEN_SetDefaultPrinter.VB_Description = "Sets the system wide default printer to that defined by the printer object"
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2002
'
'****************************************************************************
'
'                     NAME: Sub PSGEN_SetDefaultPrinter
'
'                     sPrinter$        - Printer to be the default
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    13 May 2002   First created for TeamPlayer
'
'                  PURPOSE: Sets the system wide default printer to that
'                           defined by the printer object
'
'****************************************************************************
'
'
Dim sDeviceLine$
Dim objTmp As Printer

    '
    ' Initialise the line to go in the INI file
    '
    On Error Resume Next
    For Each objTmp In Printers
        If StrComp(objTmp.DeviceName, sPrinter, vbTextCompare) = 0 Then
            sDeviceLine = objTmp.DeviceName & "," & objTmp.DriverName & "," & objTmp.Port
            Exit For
        End If
    Next objTmp
    If sDeviceLine <> "" Then
   
        '
        ' Store the new printer information in the [WINDOWS] section
        ' of the WIN.INI file for the DEVICE= item
        '
        Call WriteProfileString("windows", "Device", sDeviceLine)
          
        '
        ' Cause all applications to reload the INI file
        '
        Call SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, ByVal "windows")
    End If

End Sub

Public Function PSGEN_GetListofPrinters$()
Attribute PSGEN_GetListofPrinters.VB_Description = "Returns a null list of printers known to this server"
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2002
'
'****************************************************************************
'
'                     NAME: Function PSGEN_GetListofPrinters
'
'                          ) As String
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    04 May 2002   First created for TeamPlayer
'
'                  PURPOSE: Returns a null list of printers known to this
'                           server
'
'****************************************************************************
'
'
Dim sReturn$
Dim bSuccess As Boolean
Dim lRequired&, lBuffer&
Dim alBuffer&()
Dim lEntries&
Dim lCnt&
Dim sPName$, sSName$
Dim lAttrib, lTemp&


    '
    ' Initialise error vector
    '
    On Error Resume Next
    lBuffer = 3072
    ReDim alBuffer((lBuffer \ 4) - 1)
    bSuccess = EnumPrinters(PRINTER_ENUM_CONNECTIONS Or _
                          PRINTER_ENUM_LOCAL, _
                          vbNullString, _
                          4, _
                          alBuffer(0), _
                          lBuffer, _
                          lRequired, _
                          lEntries)
    If bSuccess Then
        If lRequired > lBuffer Then
            lBuffer = lRequired
            
            '
            ' Buffer is too small so try again with correct size
            '
            ReDim Buffer(lBuffer \ 4)
            bSuccess = EnumPrinters(PRINTER_ENUM_CONNECTIONS Or _
                                PRINTER_ENUM_LOCAL, _
                                vbNullString, _
                                4, _
                                alBuffer(0), _
                                lBuffer, _
                                lRequired, _
                                lEntries)
        End If
        
        '
        ' Produce return list
        '
        If bSuccess Then
            For lCnt = lEntries - 1 To 0 Step -1
                sPName = Space$(StrLen(alBuffer(lCnt * 3)))
                lTemp = PtrToStr(sPName, alBuffer(lCnt * 3))
                sReturn = sReturn + IIf(sReturn = "", "", vbNullChar) + sPName
                sSName = Space$(StrLen(alBuffer(lCnt * 3 + 1)))
                lTemp = PtrToStr(sSName, alBuffer(lCnt * 3 + 1))
                lAttrib = alBuffer(lCnt * 3 + 2)
            Next lCnt
        End If
    End If

    '
    ' Return value to caller
    '
    PSGEN_GetListofPrinters = sReturn

End Function

Public Function PSGEN_GetPostCodeInfo$(ByVal sPostCode$, Optional sCounty$)
Attribute PSGEN_GetPostCodeInfo.VB_Description = "Returns the postal region for the given post code.\r\nWill optionally also return the county"
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2002
'
'****************************************************************************
'
'                     NAME: Function PSGEN_GetPostCodeInfo
'
'                     sPostCode$         - UK Post code to check
'                     sCounty$           - Returned county name
'
'                          ) As String
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    25 April 2002   First created for TeamPlayer
'
'                  PURPOSE: Returns the postal region for the given post
'                           code.
'                           Will optionally also return the county
'
'****************************************************************************
'
'
Const POST_LOOKUP = "AB,Aberdeen,,AL,St Albans,,B,Birmingham,,BA,Bath,,BB,Blackburn,,BD,Bradford,,BH,Bournemouth,,BL,Bolton,,BN,Brighton,,BR,Bromley,,BS,Bristol,,BT,Belfast,,CA,Carlisle,,CB,Cambridge,," + _
                    "CF,Cardiff,,CH,Chester,,CM,Chelmsford,,CO,Colchester,,CR,Croydon,,CT,Canterbury,,CV,Coventry,,CW,Crewe,,DA,Dartford,,DD,Dundee,,DE,Derby,,DG,Dumfries & Galloway,,DH,Durham,," + _
                    "DL,Darlington,,DN,Doncaster,,DT,Dorchester,,DY,Dudley,,E,London (East),,EC,London (East Central),,EH,Edinburgh,,EN,Enfield,,EX,Exeter,,FK,Falkirk,,FY,Fylde (Blackpool),,G,Glasgow,,GL,Gloucester,," + _
                    "GU,Guildford,,GY,Guernsey & Alderney,,HA,Harrow,,HD,Huddersfield,,HG,Harrogate,,HP,Hemel Hempstead,,HR,Hereford,,HS,Western Isles (Harris),,HU,Hull,,HX,Halifax,,IG,Ilford,,IM,Isle of Man,,IP,Ipswich,," + _
                    "IV,Inverness,,JE,Jersey,,KA,Kilmarnock,,KT,Kingston upon Thames,,KW,Orkney Isles (Kirkwall),,KY,Kirkcaldy,,L,Liverpool,,LA,Lancaster,,LD,Llandrindod Wells,,LE,Leicester,,LL,Llandudno,,LN,Lincoln,," + _
                    "LS,Leeds,,LU,Luton,,M,Manchester,,ME,Medway (Rochester),,MK,Milton Keynes,,ML,Motherwell,,N,London (North),,NE,Newcastle on Tyne,,NG,Nottingham,,NN,Northampton,,NP,Newport,,NR,Norwich,," + _
                    "NW,London (North West),,OL,Oldham,,OX,Oxford,,PA,Paisley,,PE,Peterborough,,PH,Perth,,PL,Plymouth,,PO,Portsmouth,,PR,Preston,,RG,Reading,,RH,Redhill,,RM,Romford,,S,Sheffield,,SA,Swansea,," + _
                    "SE,London (South East),,SG,Stevenage,,SK,Stockport,,SL,Slough,,SM,Sutton,,SN,Swindon,,SO,Southampton,,SP,Salisbury,,SR,Sunderland,,SS,Southend on Sea,,ST,Stoke on Trent,,SW,London (South West),," + _
                    "SY,Shrewsbury,,TA,Taunton,,TD,Tweed (Galashiels),,TF,Telford,,TN,Tunbridge Wells,,TQ,Torquay,,TR,Truro,,TS,Teesside (Middlesbrough),,TW,Twickenham,,UB,Uxbridge,,W,London (West),,WA,Warrington,," + _
                    "WC,London (West Central),,WD,Watford,,WF,Wakefield,,WN,Wigan,,WR,Worcester,,WS,Walsall,,WV,Wolverhampton,,YO,York,,ZE,Shetland Isles (Lerwick)"
Const UNKNOWN_CODE = "Unknown"

Dim sReturn$, sCode$
Dim asCodes$()
Dim iCnt%


    '
    ' Sanitise the post code and get the lookup value
    '
    On Error Resume Next
    sReturn = UNKNOWN_CODE
    sCounty = UNKNOWN_CODE
    sPostCode = Replace(Trim$(UCase(sPostCode)), " ", "")
    For iCnt = 1 To Len(sPostCode)
        If IsNumeric(Mid$(sPostCode, iCnt, 1)) Then
            sCode = Left$(sPostCode, iCnt - 1)
            Exit For
        End If
    Next iCnt
    If iCnt > Len(sPostCode) Then sCode = sPostCode
    
    '
    ' Check that we have something to lookup
    '
    If sCode <> "" Then
        
        '
        ' Create an array of the strings to search through
        '
        asCodes = Split(POST_LOOKUP, ",")
        For iCnt = 0 To UBound(asCodes) \ 3
            If asCodes(iCnt * 3) = sCode Then
                sReturn = asCodes((iCnt * 3) + 1)
                sCounty = asCodes((iCnt * 3) + 2)
                Exit For
            End If
        Next iCnt
    End If

    '
    ' Return value to caller
    '
    PSGEN_GetPostCodeInfo = sReturn

End Function


Public Sub PSGEN_ShowFloatingInfo(Optional ByVal sInfo$, Optional ByVal iXOffset%, Optional ByVal iYOffset%)
Attribute PSGEN_ShowFloatingInfo.VB_Description = "Draws the floating information at the current cursor location\r\nIf the info is empty then it actually deletes it"
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2002
'
'****************************************************************************
'
'                     NAME: Sub PSGEN_ShowFloatingInfo
'
'                     iXOffset%          - X offset of the information
'                     iYOffset%          - Y offset of the information
'                     sInfo$             - Information
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    04 January 2002   First created for TeamPlayer
'
'                  PURPOSE: Draws the floating information at the current
'                           cursor location
'                           If the info is empty then it actually deletes it
'
'****************************************************************************
'
'
Dim stPoint As POINTAPI
Dim lDCSrc&, lBitmapTmp&
Dim lDesktop&

Static stRect As RECT
Static ssInfo$
Static objPic As Picture


    '
    ' Determine if we are finishing the display
    '
    On Error Resume Next
    lDesktop = GetDC(ByVal 0&)
    Call GetCursorPos(stPoint)
    If sInfo = "" Or (sInfo <> ssInfo And ssInfo <> "") Then
        lDCSrc = CreateCompatibleDC(lDesktop)
        lBitmapTmp = SelectObject(lDCSrc, objPic.Handle)
        Call BitBlt(lDesktop, stRect.Left, stRect.Top, stRect.Right - stRect.Left, stRect.Bottom - stRect.Top, lDCSrc, 0, 0, SRCCOPY)
        Call SelectObject(lDCSrc, lBitmapTmp)
        Call DeleteDC(lDCSrc)
        ssInfo = ""
        Set objPic = Nothing
        Screen.MousePointer = vbNormal
    End If
    
    '
    ' Now check if we are actually putting something out
    '
    If sInfo <> "" And ((stRect.Left - iXOffset) <> stPoint.X Or (stRect.Top - iYOffset) <> stPoint.Y) Then
        
        '
        ' If we already have a DC then replace the bit of screen with our copy
        '
        If Not objPic Is Nothing Then
            lDCSrc = CreateCompatibleDC(lDesktop)
            lBitmapTmp = SelectObject(lDCSrc, objPic.Handle)
            Call BitBlt(lDesktop, stRect.Left, stRect.Top, stRect.Right - stRect.Left, stRect.Bottom - stRect.Top, lDCSrc, 0, 0, SRCCOPY)
            Call SelectObject(lDCSrc, lBitmapTmp)
            Call DeleteDC(lDCSrc)
        End If
        
        '
        ' Create a rectangle to put the text in and output it
        '
        Call SetBkMode(lDesktop, TRANSPARENT)
        stRect.Left = stPoint.X + iXOffset
        stRect.Top = stPoint.Y + iYOffset
        Call DrawText(lDesktop, sInfo, Len(sInfo), stRect, DT_CALCRECT)
        Set objPic = CaptureWindow(0, False, stRect.Left, stRect.Top, stRect.Right - stRect.Left, stRect.Bottom - stRect.Top)
        Call DrawText(lDesktop, sInfo, Len(sInfo), stRect, 0)
        ssInfo = sInfo
    End If
    Call ReleaseDC(0, lDesktop)

End Sub




Public Function PSGEN_GetGreyScale&(ByVal lColor&)
Attribute PSGEN_GetGreyScale.VB_Description = "Returns the greyscale value of the color"
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2001
'
'****************************************************************************
'
'                     NAME: Function PSGEN_GetGreyScale
'
'                     lColor&            - color value
'
'                          ) As Long
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    05 December 2001   First created for MediaWeb
'
'                  PURPOSE: Returns the greyscale value of the color
'
'****************************************************************************
'
'

    '
    ' Initialise error vector
    '
    On Error Resume Next
    PSGEN_GetGreyScale = ((77& * (lColor And &HFF&) + _
                 152& * (lColor And &HFF00&) \ &H100& + _
                  28& * ((lColor And &HFF0000) \ &H10000)) \ 256&) * &H10101

End Function

Public Function PSGEN_TranslateComErrorCode$(Optional ByVal lError)
Attribute PSGEN_TranslateComErrorCode.VB_Description = "Returns the message for the given error code returned after a COM object failure"
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2001
'
'****************************************************************************
'
'                     NAME: Function PSGEN_TranslateComErrorCode
'
'                     lError&            - Normal VB Err.Number
'
'                          ) As String
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    20 September 2001   First created for MediaWeb
'
'                  PURPOSE: Returns the message for the given error code
'                           returned after a COM object failure
'
'****************************************************************************
'
'
Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000

Dim lRet&
Dim sReturn$


    '
    ' Initialise error vector
    '
    On Error Resume Next
    If IsMissing(lError) Then lError = GetLastError
    sReturn = Space(256)
    lRet = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0&, lError, 0&, sReturn, 256&, 0&)
    If lRet > 0 Then
       sReturn = Left$(sReturn, lRet)
    Else
       sReturn = "Error not found"
    End If

    '
    ' Return value to caller
    '
    PSGEN_TranslateComErrorCode = sReturn

End Function

Public Function PSGEN_SwapWebCharacters$(ByVal sValue$)
Attribute PSGEN_SwapWebCharacters.VB_Description = "Changes all non alphanumeric characters to their web equivalent"
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2001
'
'****************************************************************************
'
'                     NAME: Function PSGEN_SwapWebCharacters
'
'                     sValue$            - Value to change
'
'                          ) As String
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    03 July 2001   First created for BikeSwapMonitor
'
'                  PURPOSE: Changes all non alphanumeric characters to their
'                           web equivalent
'
'****************************************************************************
'
'
Dim sReturn$, sTmp$
Dim iCnt%


    '
    ' Loop round each character
    '
    On Error Resume Next
    sValue = Replace(sValue, vbCrLf, vbCr)
    For iCnt = 1 To Len(sValue)
        Select Case Mid$(sValue, iCnt, 1)
            Case "a" To "z", "A" To "Z", "0" To "9", "%", vbCr
                sReturn = sReturn + Mid$(sValue, iCnt, 1)
            Case Else
                sTmp = Hex(Asc(Mid$(sValue, iCnt, 1)))
                If Len(sTmp) = 1 Then sTmp = "0" + sTmp
                sReturn = sReturn + "%" + sTmp
        End Select
    Next iCnt
    sReturn = Replace(sReturn, vbCr, "%0D%0A")

    '
    ' Return value to caller
    '
    PSGEN_SwapWebCharacters = sReturn

End Function

Public Function PSGEN_SwapWebEntities$(ByVal sValue$)
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2001
'
'****************************************************************************
'
'                     NAME: Function PSGEN_SwapWebEntities
'
'                     sValue$            - Value to change
'
'                          ) As String
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    03 July 2001   First created for BikeSwapMonitor
'
'                  PURPOSE: Changes all non alphanumeric characters to their
'                           web equivalent
'
'****************************************************************************
'
'
Dim sReturn$, sTmp$
Dim iCnt%


    '
    ' Loop round each character
    '
    On Error Resume Next
    sReturn = Replace(sValue, "+", " ")
    If InStr(sReturn, "%") > 0 Then
        For iCnt = 1 To 255
            sTmp = Hex(iCnt)
            If Len(sTmp) = 1 Then sTmp = "0" + sTmp
            sReturn = Replace(sReturn, "%" + sTmp, Chr(iCnt), compare:=vbTextCompare)
        Next iCnt
    End If
    
    '
    ' Return value to caller
    '
    PSGEN_SwapWebEntities = sReturn

End Function

Public Sub PSGEN_AssignPictureToMenu(ByVal objPicture As Object, ByVal objMenu As Menu)
Attribute PSGEN_AssignPictureToMenu.VB_Description = "Assigns the picture given by the picture handle to the menu unchecked bitmap"
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2001
'
'****************************************************************************
'
'                     NAME: Sub PSGEN_AssignPictureToMenu
'
'                     objPicture As Picture        - Picture object
'                     objMenu As Menu           - Menu to assign picture to
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    01 April 2001   First created for TeamPlayer
'
'                  PURPOSE: Assigns the picture given by the picture handle
'                           to the menu unchecked bitmap
'
'****************************************************************************
'
'

Dim lMenu&, lMenuID&


    '
    ' Find the menu in question
    '
    On Error Resume Next
    lMenu = PSGEN_GetMenuHandle(objMenu, lMenuID)
    If lMenu <> 0 Then
    
        '
        ' We've found it so now assign the picture to the menu
        '
        Call SetMenuItemBitmaps(lMenu, lMenuID, MF_BYCOMMAND, objPicture, 0)
    End If

End Sub



Public Function PSGEN_GetMenuHandle&(ByVal objMenu As Menu, ByRef lRetMenuID&, Optional lMenu& = 0)
Attribute PSGEN_GetMenuHandle.VB_Description = "Returns the API ID of the menu supplied"
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2001
'
'****************************************************************************
'
'                     NAME: Function PSGEN_GetMenuID
'
'                     objMenu As menu           - The menu to get the ID from
'                     lRetMemuID&               - Menu to which the item belongs
'                     lMenu&                    - Menu to search within
'
'                          ) As Long
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    31 March 2001   First created for TeamPlayer
'
'                  PURPOSE: Returns the API ID of the menu supplied
'
'****************************************************************************
'
'
Dim lReturn&, lMenuID&
Dim iMenus%, iCnt%
Dim sCaption$


    '
    ' Get the top level menu value if it hasn't been provided
    '
    On Error Resume Next
    If lMenu = 0 Then lMenu = GetMenu(objMenu.Parent.hWnd)

    '
    ' Loop through all the submenus looking for our caption
    '
    iMenus = GetMenuItemCount(lMenu)
    For iCnt = 0 To iMenus - 1
        
        '
        ' Get the menu caption
        '
        lMenuID = GetMenuItemID(lMenu, iCnt)
        sCaption = String(127, vbNullChar)
        Call GetMenuString(lMenu, iCnt, sCaption, 127, MF_BYPOSITION)
        sCaption = PSGEN_GetItem(1, vbNullChar, sCaption)
        If StrComp(sCaption, objMenu.Caption, vbTextCompare) = 0 Then
            lRetMenuID = lMenuID
            lReturn = lMenu
            Exit For
        Else
            
            '
            ' If this is a popup then get it's handle and call ourselves again
            '
            If lMenuID = -1 And GetSubMenu(lMenu, iCnt) <> 0 Then
                lReturn = PSGEN_GetMenuHandle(objMenu, lRetMenuID, GetSubMenu(lMenu, iCnt))
                If lReturn <> 0 Then Exit For
            End If
        End If
    Next iCnt


    '
    ' Return value to caller
    '
    PSGEN_GetMenuHandle = lReturn

End Function



Public Sub PSGEN_PositionForm(ByVal frmSubject As Form, Optional ByVal frmParent As Form = Nothing)
Attribute PSGEN_PositionForm.VB_Description = "Centers the form on the parent and if that isn't provided then it centres it on the screen"
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2001
'
'****************************************************************************
'
'                     NAME: Sub PSGEN_PositionForm
'
'                     frmSubject As Form        - Form to position
'                     frmParent As Form         - Form to centre it on
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    19 March 2001   First created for TeamPlayer
'
'                  PURPOSE: Centers the form on the parent and if that isn't
'                           provided then it centres it on the screen
'
'****************************************************************************
'
'


    '
    ' Determine what to centre on
    '
    On Error Resume Next
    If frmParent Is Nothing Then
        frmSubject.Move (Screen.Width - frmSubject.Width) / 2, (Screen.Height - frmSubject.Height) / 3
    Else
        frmSubject.Move frmParent.Left + (frmParent.Width - frmSubject.Width) / 2, frmParent.Top + (frmParent.Height - frmSubject.Height) / 2
        If frmSubject.Top < 0 Then frmSubject.Top = 0
        If frmSubject.Left < 0 Then frmSubject.Left = 0
        If frmSubject.Top + frmSubject.Height > Screen.Height Then frmSubject.Top = Screen.Height - frmSubject.Height
        If frmSubject.Left + frmSubject.Width > Screen.Width Then frmSubject.Left = Screen.Width - frmSubject.Width
    End If

End Sub

Public Function PSGEN_GetItemNumber%(ByVal sSep$, ByVal sItemValue$, ByVal sList$)
Attribute PSGEN_GetItemNumber.VB_Description = "Returns the item number for the value and separator supplied"
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME: Function PSGEN_GetItemNumber
'
'                     sSep$              - Separators between items
'                     sItemValue$        - Item to look for
'                     sList$             - List of items
'
'                          ) As Integer
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    14 December 2000   First created for PivotalBASISSlave
'
'                  PURPOSE: Returns the item number for the value and
'                           separator supplied
'
'****************************************************************************
'
'
Dim iReturn%, iCnt%
Dim asTmp$()


    '
    ' Get a list of items and loop through them all
    '
    iReturn = 0
    On Error Resume Next
    asTmp = Split(sList, sSep, compare:=vbTextCompare)
    For iCnt = 0 To UBound(asTmp)
        If StrComp(asTmp(iCnt), sItemValue, vbTextCompare) = 0 Then
            iReturn = iCnt + 1
            Exit For
        End If
    Next iCnt

    '
    ' Return value to caller
    '
    PSGEN_GetItemNumber = iReturn

End Function

Public Sub PSGEN_DrawText(ByVal sValue$, ByVal lDC&, stRect As RECT)
Attribute PSGEN_DrawText.VB_Description = "Draws the text using the rectangle supplied onto the device context supplied.\r\nExpands the bottom of the box to accomodate the text"
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME: Sub PSGEN_DrawText
'
'                     sValue$            - String to draw
'                     lDC&               - Drawing plane
'                     stRect As RECT     - Drawing bounding rectangle
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    13 September 2000   First created for GURU
'
'                  PURPOSE: Draws the text using the rectangle supplied onto
'                           the device context supplied.
'                           Expands the bottom of the box to accomodate the
'                           text
'
'****************************************************************************
'
'


    '
    ' Adjust the rectangle and then draw the text
    '
    On Error Resume Next
    Call DrawText(lDC, sValue, Len(sValue), stRect, DT_CALCRECT + DT_WORDBREAK)
    Call DrawText(lDC, sValue, Len(sValue), stRect, DT_WORDBREAK)

End Sub

Public Function PSGEN_ConvertImage$(ByVal sFile$, ByVal lWidth&, ByVal lHeight&, ByVal sType$, Optional ByVal sExtra$ = "")
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME: Function PSGEN_ConvertImage
'
'                     sFile$             - File to convert
'                     lWidth&            - Width in pixels
'                     lHeight&           - Height in pixels
'                     sType$             - Extension depicting the type
'                     sExtra$            - Extra options to add
'
'                          ) As String
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    12 June 2000   First created for BikeSwap
'
'                  PURPOSE: Converts the file to the type specified and
'                           scales it to the dimensions passed.
'                           This functions uses the ImageMagick set of
'                           utilities and so expects the neccersary DLLs and
'                           EXEs to be in the path.
'                           Returns a temporary file if successful otherwise
'                           rasies an error.
'
'****************************************************************************
'
'
Dim sReturn$, sFeedback$, sCommand$, sPath$
Dim sError$, sBatch$, sGeometery$, sCmd$
Dim lActWidth&, lActHeight&

    
    '
    ' We need to get the size of the file
    '
    On Error Resume Next
    sPath = Environ(PSGEN_IMAGEMAGICK)
    If sPath = "" Then
        Err.Raise vbObjectError + PSGEN_ERROR_BASE, "PSGEN_ConvertImage", "Cannot find the path to image conversion library - missing env. variable " + PSGEN_IMAGEMAGICK
    Else
        sCommand = PSGEN_GetFileAttributes(sFile, True)
        If Err = 0 Then
        
            '
            ' The width and height are after the name
            '
            sCommand = Trim$(PSGEN_GetItem(1, vbCrLf, PSGEN_GetItem(2, "Geometry:", sCommand)))
            lActWidth = CLng(Trim$(PSGEN_GetItem(1, "x", sCommand)))
            lActHeight = CLng(Trim$(PSGEN_GetItem(2, "x", sCommand)))
            sGeometery = Format$(CLng(Format$((lWidth / lActWidth) * 100))) + "x" + Format$(CLng(Format$((lHeight / lActHeight) * 100))) + "%%"
            
            '
            ' Build the required command line
            '
            sFeedback = PSGEN_GetTempPathFilename
            Kill sFeedback
            sBatch = PSGEN_GetTempPathFilename("bat")
            Kill sBatch
            sReturn = PSGEN_GetTempPathFilename(sType)
            Kill sReturn
            If PSGEN_IsSystemNT Then
                Call PSGEN_WriteTextFile("cmd.exe /c """ + sPath + "\convert"" -geometery " + sGeometery + IIf(sExtra = "", " ", " " + sExtra + " ") + sFile + " " + sReturn + " >" + sFeedback, sBatch, sError)
                sCmd = "cmd.exe /c " + sBatch
            Else
                Call PSGEN_WriteTextFile("""" + sPath + "\convert"" -geometery " + sGeometery + IIf(sExtra = "", " ", " " + sExtra + " ") + sFile + " " + sReturn + " >" + sFeedback, sBatch, sError)
                sCmd = "command.com /c " + sBatch
            End If
        
            '
            ' Run the command and gather the output
            '
            If PSGEN_WaitForTask(sCmd, "Converting " + sFile + " to " + sReturn, True, 20000) <> 0 Then
                Kill sReturn
                Call PSGEN_ReadTextFile(sFeedback, sCommand, sError)
                If sCommand = "" Then
                    Err.Raise vbObjectError + PSGEN_ERROR_BASE, "PSGEN_ConvertImage", "Conversion taking too long"
                Else
                    Err.Raise vbObjectError + PSGEN_ERROR_BASE, "PSGEN_ConvertImage", sCommand
                End If
            Else
                '
                ' Check that there was no feedback
                '
                Call PSGEN_ReadTextFile(sFeedback, sCommand, sError)
                If sCommand <> "" Then
                    Kill sReturn
                    Err.Raise vbObjectError + PSGEN_ERROR_BASE, "PSGEN_ConvertImage", sCommand
                End If
            End If
            Kill sFeedback
            Kill sBatch
        Else
            Err.Raise Err.Number, Err.Source, Err.Description
        End If
    End If

    '
    ' Return value to caller
    '
    PSGEN_ConvertImage = sReturn

End Function

Public Function PSGEN_InDevelopment(Optional bSetMode As Boolean = False) As Boolean
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME: Public Function PSGEN_InDevelopment
'
'                           oClassName      - Class name of the project
'
'             DEPENDENCIES: None
'
'     MODIFICATION HISTORY: Steve O'Hara    126th January 1998   First created for RTE from Desktop
'
'                  PURPOSE: This function returns true if we are in the
'                           VB IDE, otherwise false.
'
'****************************************************************************
'
'
Static sbReturn As Boolean
 
    sbReturn = bSetMode

    If Not sbReturn Then
        Debug.Assert PSGEN_InDevelopment(True)
    End If

    PSGEN_InDevelopment = sbReturn

End Function


Public Function PSGEN_GetFolder$(ByVal sTitle$, Optional ByVal lOwner&, Optional sInstructions$, Optional ByVal sStartFolder$, Optional ByVal iSystemFolder%, Optional ByVal lRestrictions&)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME: Function PSGEN_GetFolder
'
'                     sTitle$            - Title of dialog
'                     lOwner&            - Owner window handle
'                     sInstructions$     - Instructions to the user
'                     sStartFolder$      - Starting folder to look in
'                     iSystemFolder%     - System folder to look inside
'                     iRestrictions%     - Restrictions on browse
'
'                          ) As String
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    28 October 1998   First created for Willow
'
'                  PURPOSE: Returns te name of the folder selected by the user
'
'****************************************************************************
'
'
Const MAX_PATH = 260

Dim sReturn$
Dim lFolder&, lIDL&
Dim stIDL As ITEMIDLIST

    '
    ' Initialise error vector and set up the browse information
    '
    sReturn = ""
    On Error Resume Next
    With mstBrowse
        .hOwner = IIf(IsMissing(lOwner), 0&, lOwner)
        lFolder = IIf(IsMissing(iSystemFolder), 0, iSystemFolder)
        .pidlRoot = 0
        If SHGetSpecialFolderLocation(mstBrowse.hOwner, ByVal lFolder, stIDL) = 0 Then .pidlRoot = stIDL.mkid.cb
        .pszDisplayName = String$(MAX_PATH, 0)
        .lpszTitle = sInstructions
        .ulFlags = IIf(IsMissing(lRestrictions), 0, lRestrictions)
        .lpfn = PSGEN_AddressOf(AddressOf Z_FolderCallback)
    End With
    msBrowseInitDir = sStartFolder
    msBrowseTitle = sTitle
    
  
    '
    ' Show the selection dialog
    '
    lIDL = SHBrowseForFolder(mstBrowse)
  
    '
    ' If the dialog wasn't cancelled
    '
    If lIDL <> 0 Then
    
        '
        ' Get the path from the ID list
        '
        sReturn = String$(MAX_PATH, 0)
        Call SHGetPathFromIDList(ByVal lIDL, ByVal sReturn)
        sReturn = Left$(sReturn, InStr(sReturn, vbNullChar) - 1)
    End If
  
    '
    ' Frees the memory SHBrowseForFolder() allocated for the pointer to the item id list
    '
    Call CoTaskMemFree(lIDL)

    '
    ' Return value to caller
    '
    PSGEN_GetFolder = sReturn

End Function



Public Function PSGEN_FindFileInPath$(ByVal sFilename$)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME: Function PSGEN_FindFileInPath
'
'                     sFilename$         - Name of file to look for
'
'                          ) As String
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    22 April 1998   First created for Willow
'
'                  PURPOSE: Returns the fully specified path for the filename
'                            by looking for the file in all the directories
'                            specified in the PATH environment variable
'
'****************************************************************************
'
'
Dim sReturn$, sPath$, sDir$
Dim iCnt%, iItems%


    '
    ' Initialise error vector and get the path
    '
    On Error Resume Next
    sReturn = ""
    sFilename = Trim$(sFilename)
    sPath = Environ("PATH")
    If sPath <> "" Then

        '
        ' Loop round each of the directories in the path looking to see if the file is there
        '
        iItems = PSGEN_GetNoOfItems(";", sPath)
        For iCnt = 1 To iItems
            sDir = Trim$(PSGEN_GetItem(iCnt, ";", sPath))
            If sDir <> "" Then
                
                '
                ' Check the directory and file
                '
                If Right$(sDir, 1) <> "\" Then sDir = sDir + "\"
                If PSGEN_FileExists(sDir + sFilename) Then sReturn = sDir + sFilename
            End If
        Next iCnt
    End If

    '
    ' Return value to caller
    '
    PSGEN_FindFileInPath = sReturn

End Function
Public Function PSGEN_DirectoryExists(ByVal sDirname$) As Boolean
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME: Function PSGEN_DirectoryExists
'
'                     sDirname$         - Name of directory to check
'
'                          ) As Integer
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    22 April 1998   First created for Willow
'
'                  PURPOSE: Returns true if the the fully specified path
'                           exists
'
'****************************************************************************
'
'
Dim bReturn As Boolean
Dim sTmp$


    '
    ' Initialise error vector
    '
    On Error Resume Next
    sDirname = Trim$(PSGEN_GetItem(1, vbNullChar, sDirname))
    If sDirname <> "" Then

        '
        ' Check for the name
        '
        sTmp = Dir(sDirname, vbDirectory)
        bReturn = (Err = 0) And (sTmp <> "")
    End If

    '
    ' Return value to caller
    '
    PSGEN_DirectoryExists = bReturn

End Function



Public Function PSGEN_FileExists%(ByVal sFilename$)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME: Function PSGEN_FileExists
'
'                     sfilename$         - Name of file to check
'
'                          ) As Integer
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    22 April 1998   First created for Willow
'
'                  PURPOSE: Returns true if the the fully specified path
'                            filename exists and can be opened for read access
'
'****************************************************************************
'
'
Dim iReturn%, iFile%


    '
    ' Initialise error vector
    '
    On Error Resume Next
    iReturn = False
    sFilename = Trim$(sFilename)

    '
    ' First check to see if it exists using the directory command
    '
    If Left$(sFilename, 1) = "'" Then sFilename = Right$(sFilename, Len(sFilename) - 1)
    If Right$(sFilename, 1) = "'" Then sFilename = Left$(sFilename, Len(sFilename) - 1)
    If Trim$(sFilename) <> "" And Dir(sFilename, vbNormal Or vbHidden Or vbReadOnly) <> "" Then
    
        '
        ' Try and open the file for read and shared access
        '
        iFile = FreeFile
        Open sFilename For Input Access Read As iFile
        If Err = 0 Then
            iReturn = True
            Close iFile
        End If
    End If
    Err.Clear

    '
    ' Return value to caller
    '
    PSGEN_FileExists = iReturn

End Function


Public Function PSGEN_FindWindowLike%(hWndArray&(), ByVal hWndStart&, sWindowText$, sClassName$, vID As Variant)
'****************************************************************************
'
'       Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'       NAME:           Public Function PSGEN_FindWindowLike (
'
'                       hWndArray&()    - A long array used to return the window handles
'                       hWndStart&      - The handle of the window to search under.
'                       sWindowText$    - The pattern used with the Like operator to compare window's text.
'                       sClassName$     - The pattern used with the Like operator to compare window's class name.
'                       vID             - A child vID number used to identify a window.
'
'                       )       - Integer
'
'       DEPENDENCIES:   NONE
'
'       PURPOSE:        Finds the window handles of the windows matching the specified
'                       parameters.  The routine searches through all of this hWndStart's
'                       children and their children recursively. If hWndStart = 0 then
'                       the routine searches through all windows.  vID can be a decimal
'                       number or a hex string. Prefix hex strings with "&H" or an error
'                       will occur. To ignore the vID pass the Visual Basic Null function.
'                       Returns the number of windows that matched the parameters.
'                       Also returns the window handles in hWndArray()
'
'       MODIFICATION HISTORY: Steve O'Hara  10th July 1996   First created for Pivotal Addin
'
'****************************************************************************
'

Dim hWnd&
Dim sLocalWindowText$
Dim sLocalClassName$
Dim vIdentity
Dim lTmp&
Static iLevel%
Static iFound%

    '
    ' Initialize first time through
    '
    On Error Resume Next
    If iLevel = 0 Then
        iFound = 0
        ReDim hWndArray(0 To 0)
        If hWndStart = 0 Then hWndStart = GetDesktopWindow()
    End If

    '
    ' Increase recursion counter:
    '
    iLevel = iLevel + 1

    '
    ' Get first child window:
    '
    hWnd = GetWindow(hWndStart, GW_CHILD)
    Do Until hWnd = 0
        
        '
        ' Search children by recursion:
        '
        lTmp = PSGEN_FindWindowLike(hWndArray(), hWnd, sWindowText, sClassName, vID)

        '
        ' Get the window text and class name:
        '
        sLocalWindowText = Space(255)
        lTmp = GetWindowText(hWnd, sLocalWindowText, 255)
        sLocalWindowText = Left$(sLocalWindowText, lTmp)
        sLocalClassName = Space(255)
        lTmp = GetClassName(hWnd, sLocalClassName, 255)
        sLocalClassName = Left$(sLocalClassName, lTmp)

        '
        ' If window is a child get the vID:
        '
        If GetParent(hWnd) <> 0 Then
            lTmp = GetWindowWord(hWnd, GWW_ID)
            vIdentity = CLng("&H" & Hex(lTmp))
        Else
            vIdentity = Null
        End If

        '
        ' Check that window matches the search parameters:
        '
        If (sLocalWindowText Like sWindowText) And ((sLocalClassName Like sClassName) Or (sClassName = "")) Then
            If IsNull(vID) Then
            
                '
                ' If find a match, increment counter and add handle to array:
                '
                iFound = iFound + 1
                ReDim Preserve hWndArray(0 To iFound)
                hWndArray(iFound) = hWnd
            
            ElseIf Not IsNull(vIdentity) Then
                If vIdentity = CLng(vID) Then
                    '
                    ' If find a match increment counter and add handle to array:
                    '
                    iFound = iFound + 1
                    ReDim Preserve hWndArray(0 To iFound)
                    hWndArray(iFound) = hWnd
                End If
            End If
        End If

        '
        ' Get next child window:
        '
         hWnd = GetWindow(hWnd, GW_HWNDNEXT)
    Loop

    '
    ' Decrement recursion counter:
    '
    iLevel = iLevel - 1

    '
    ' Return the number of windows found:
    '
    PSGEN_FindWindowLike = iFound

End Function







Public Function PSGEN_ShutdownRequested%()
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME:   Function PSGEN_ShutdownRequested
'
'
'                          ) As Integer
'
'             DEPENDENCIES:   NONE
'
'     MODIFICATION HISTORY:   Steve O'Hara    21 January 1998   First created for DocBlazer
'
'                  PURPOSE:   Returns the state of the shutdown flag
'
'****************************************************************************
'
'
Dim iReturn%


    '
    ' Initialise error vector
    '
    On Error Resume Next

    '
    ' Return value to caller
    '
    PSGEN_ShutdownRequested = miShutdownRequested

End Function

Public Sub PSGEN_RequestShutdown(ByVal iValue%)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME:   Sub PSGEN_RequestShutdown
'
'                             iValue%     - Value to set flag to
'
'             DEPENDENCIES:   NONE
'
'     MODIFICATION HISTORY:   Steve O'Hara    21 January 1998   First created for DocBlazer
'
'                  PURPOSE:   Sets the module flag that tells the system that a
'                             shutdown has been requested by the user.
'
'****************************************************************************
'
'


    '
    ' Initialise error vector
    '
    On Error Resume Next
    miShutdownRequested = iValue

End Sub



Public Function PSGEN_ReadTextFile(ByVal sFilename$, sValue$, Optional sError$) As Boolean
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME:   Function PSGEN_ReadTextFile
'
'                     sFilename$         - Filename to get data from
'                     sValue$            - Variable to populate
'                     sError$            - Error description
'
'             DEPENDENCIES:   NONE
'
'     MODIFICATION HISTORY:   Steve O'Hara    29 January 1998   First created for DocBlazer
'
'                  PURPOSE:   Gets the contents of the text file into sValue.
'                             The limitation to this routine is that the file
'                             cannot be greater than 32k bytes long.
'
'****************************************************************************
'
'
Dim iFile%
Dim bReturn As Boolean


    '
    ' Initialise error vector and check data
    '
    On Error Resume Next
    bReturn = False
    If PSGEN_FileExists(sFilename) Then
        iFile = FreeFile
        Open sFilename For Binary Access Read As #iFile
        If Err <> 0 Then
            sError = "Cannot open file: '" + sFilename + "' - " + Err.Description
        Else
        
            '
            ' Determine how many chunks there are
            '
            sValue = Space(LOF(iFile))
            Get iFile, , sValue
            
            '
            ' Check for any errors
            '
            If Err <> 0 Then
                sError = "Cannot read from text file: '" + sFilename + "' - " + Err.Description
                bReturn = False
            Else
                bReturn = True
            End If
            Close #iFile
        End If
    Else
        sError = "File does not exist or cannot be opened for reading - " + sFilename
    End If
    
    '
    ' Return to caller
    '
    PSGEN_ReadTextFile = bReturn

End Function

    







Function PSGEN_SetStringWidth$(frmback As Form, ByVal iWidth%, ByVal sVal$)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME:   Function PSGEN_SetStringwidth
'
'                     frmBack as Form    - Form to run font checks on
'                     iWidth%            - Width in pixels to space the string to
'                     sVal$              - String to stretch
'
'                          ) As String
'
'             DEPENDENCIES:   NONE
'
'     MODIFICATION HISTORY:   steve o'hara    11 December 1997   First created for docblazer
'
'                  PURPOSE:   Stretches the string to the desired width based
'                             upon the font characteristics of the specified form
'
'****************************************************************************
'
'

Dim sReturn$

    '
    ' Loop round adding spaces to the string to get it to the desired width
    '
    sReturn = sVal
    On Error Resume Next
    Do While frmback.TextWidth(sReturn) < iWidth
        sReturn = sReturn + " "
    Loop

    PSGEN_SetStringWidth = sReturn$
    
End Function


Public Function PSGEN_WriteTextFile%(ByVal sMessage$, ByVal sFilename$, Optional sError$, Optional bAppendFile As Boolean = False)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME:   Function PSGEN_WriteTextFile
'
'                     sMessage$          - Data to output to the file
'                     sFilename$         - Filename to create
'                     sError$            - Errors encounterd
'                     oAppendFile as Variant - File to append to
'
'                          ) As Integer
'
'             DEPENDENCIES:   NONE
'
'     MODIFICATION HISTORY:   steve o'hara    11 December 1997   First created for docblazer
'
'                  PURPOSE:   Creates a file that contains the data supplied in
'                             the message.  Returns any errors encountered.
'
'****************************************************************************
'
'
Dim iReturn%, iFile%

    '
    ' Initialise error vector and file handle
    '
    On Error Resume Next
    iFile = FreeFile
    sError = ""
    iReturn = False
        
    '
    ' Open a file for output
    '
    If bAppendFile Then
        Open sFilename For Binary Access Read Write As #iFile
        If LOF(iFile) > 0 Then Seek iFile, LOF(iFile) + 1
    Else
        If PSGEN_FileExists(sFilename) Then Kill sFilename
        Open sFilename For Binary Access Write As #iFile
    End If
    If Err <> 0 Then
        sError = "Cannot create data file: '" + sFilename + "' - " + Err.Description
    Else

        '
        ' Write message body out to the file
        '
        Put #iFile, , sMessage
        If Err <> 0 Then
            sError = "Cannot write to data file: '" + sFilename + "' - " + Err.Description
        Else
            iReturn = True
        End If
    End If
    Close iFile
    
    '
    ' Return value to caller
    '
    PSGEN_WriteTextFile = iReturn

End Function



Public Function PSGEN_GetKeyStrokes$(ByVal iKeyCode%)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME:   Function PSGEN_GetKeyStrokes
'
'                     iKeycode%          - Key just pressed
'
'                          ) As String
'
'             DEPENDENCIES:   NONE
'
'     MODIFICATION HISTORY:   Steve O'Hara    09 December 1997   First created for DocBlazer
'
'                  PURPOSE:   Returns the currently typed in string of
'                             characters based upon how close together the keys
'                             have been pressed.  If more than KEY_WAIT
'                             milliseconds has elapsed since the last call to
'                             this routine, then only the passed keystroke is returned.
'
'****************************************************************************
'
'
Const KEY_WAIT = 1000

Static sCurrentKeys$
Static lCurrentTime&


    '
    ' Initialise error vector
    '
    On Error Resume Next
    If iKeyCode = 0 Then
        sCurrentKeys = ""
    Else
        If (GetTickCount - lCurrentTime) > KEY_WAIT Then
            sCurrentKeys = Chr(iKeyCode)
        Else
            sCurrentKeys = sCurrentKeys + Chr(iKeyCode)
        End If
    End If
    lCurrentTime = GetTickCount

    '
    ' Return value to caller
    '
    PSGEN_GetKeyStrokes = sCurrentKeys

End Function


    

Public Sub PSGEN_SelectText(ByVal ctlText As Object)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME:   Sub PSGEN_SelectText
'
'                     ctlText As Control           - Textbox to use
'
'             DEPENDENCIES:   NONE
'
'     MODIFICATION HISTORY:   Steve O'Hara    27 November 1997   First created for DocBlazer
'
'                  PURPOSE:   Selects all the text in a textbox
'
'****************************************************************************
'
'


    '
    ' Initialise error vector
    '
    On Error Resume Next
    ctlText.SelStart = 0
    ctlText.SelLength = Len(ctlText.Text)
    
End Sub


Public Sub PSGEN_ViewInNotepad(sValue$, Optional ByVal bWait As Boolean = True)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME:   Sub PSGEN_ViewInNotepad
'
'                     sValue$            - Value to edit
'
'             DEPENDENCIES:   NONE
'
'     MODIFICATION HISTORY:   Steve O'Hara    04 November 1997   First created for DocBlazer
'
'                  PURPOSE:   Views the specified string in Notepad by writing
'                             the string out to a temporary text file and
'                             setting Notepad off.  It then waits for the
'                             Notepad session to complete before it returns the
'                             string to the user.
'
'****************************************************************************
'
'
Dim sFile$, sTmp$


    '
    ' Create a temporary filename and fill it with the data
    '
    On Error Resume Next
    sFile = PSGEN_GetTempPathFilename()
    Call PSGEN_WriteTextFile(sValue, sFile, sTmp)

    '
    ' Now launch notepad and wait for it to finish
    '
    If bWait Then
        Call PSGEN_WaitForTask("notepad.exe " & sFile, "Edit " + sFile, False)
        
        '
        ' Now get back the data from the file
        '
        Call PSGEN_ReadTextFile(sFile, sValue, sTmp)
        If PSGEN_FileExists(sFile) Then Kill sFile
    Else
        Shell "notepad.exe " & sFile, vbNormalFocus
    End If
    
End Sub


Public Function PSGEN_WaitForTask&(ByVal sTask$, ByVal sTitle$, ByVal iHidden%, Optional vTimeout As Variant)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME:   Function PSGEN_WaitForTask
'
'                     sTask$             - Task to run
'                     sTitle$            - Title to use
'                     iHidden%           - Boolean True=Hidden, False=Normal
'                     vTimeout as variant- Optional timeout
'
'                     ) as long          - Wait failure code
'
'             DEPENDENCIES:   NONE
'
'     MODIFICATION HISTORY:   Steve O'Hara    04 November 1997   First created for DocBlazer
'
'                  PURPOSE:   Runs the shell task and waits for it to complete.
'                             If specified will wait vTimeout seconds otherwise
'                             the wait is infinite.
'
'****************************************************************************
'
'
Dim stProc As PROCESS_INFORMATION
Dim stStart As STARTUPINFO
Dim stSecurity As SECURITY_ATTRIBUTES
Dim lRet&, lTimeout&
Dim iStatus%
Dim frmTmp As Form


    '
    ' Initialise error vector
    '
    On Error Resume Next
    If IsMissing(vTimeout) Then
        lTimeout = INFINITE
    Else
        If vTimeout <= 0 Then
            lTimeout = INFINITE
        Else
            lTimeout = vTimeout
        End If
    End If

    '
    ' Initialize the stStartupInfo structure:
    '
    Set frmTmp = Screen.ActiveForm
    Call PSGEN_SetFormsEnabled(False)
    stStart.lpTitle = sTitle
    stStart.cb = Len(stStart)
    If iHidden Then
        stStart.dwFlags = STARTF_USESHOWWINDOW
        stStart.wShowWindow = SW_HIDE
    End If
    stSecurity.nLength = Len(stSecurity)
    stSecurity.bInheritHandle = 0
    stSecurity.lpSecurityDescriptor = PROCESS_ALL_ACCESS
    '
    ' Start the shelled application:
    '
    lRet = CreateProcessA(vbNullString, sTask, stSecurity, 0&, 1&, NORMAL_PRIORITY_CLASS + CREATE_SEPARATE_WOW_VDM, 0&, vbNullString, stStart, stProc)
      
    '
    ' Wait for the shelled application to finish:
    '
    lRet = WaitForSingleObject(stProc.hProcess, lTimeout)
    If lRet <> 0 Then Call TerminateProcess(stProc.hProcess, 0&)
    Call CloseHandle(stProc.hProcess)
    DoEvents
    Call PSGEN_SetFormsEnabled(True)
    If iHidden And Not frmTmp Is Nothing Then frmTmp.SetFocus
    
    PSGEN_WaitForTask = lRet
    
End Function

Public Function PSGEN_PollForTask&(ByVal sTask$, ByVal sTitle$, ByVal iHidden%, Optional vTimeout As Variant)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME:   Function PSGEN_PollForTask
'
'                     sTask$             - Task to run
'                     sTitle$            - Title to use
'                     iHidden%           - Boolean True=Hidden, False=Normal
'                     vTimeout as variant- Optional timeout
'
'                     ) as long          - Wait failure code
'
'             DEPENDENCIES:   NONE
'
'     MODIFICATION HISTORY:   Steve O'Hara    04 November 1997   First created for DocBlazer
'
'                  PURPOSE:   Runs the shell task and waits for it to complete
'                             by polling to see if it is still running.
'                             If specified will wait vTimeout seconds otherwise
'                             the wait is infinite.
'
'****************************************************************************
'
'
Const POLL_TIMEOUT = 1000
Dim stProc As PROCESS_INFORMATION
Dim stStart As STARTUPINFO
Dim stSecurity As SECURITY_ATTRIBUTES
Dim lRet&, lTimeout&, lTmp&, lStartTime&
Dim iStatus%, iFinished%


    '
    ' Initialise error vector
    '
    On Error Resume Next
    gbPollForTaskFinish = False
    If IsMissing(vTimeout) Then
        lTimeout = INFINITE
    Else
        If vTimeout <= 0 Then
            lTimeout = INFINITE
        Else
            lTimeout = vTimeout
        End If
    End If
    '
    ' Initialize the stStartupInfo structure:
    '
    stStart.lpTitle = sTitle
    stStart.cb = Len(stStart)
    If iHidden Then
        stStart.dwFlags = STARTF_USESHOWWINDOW
        stStart.wShowWindow = SW_HIDE
    End If
    stSecurity.nLength = Len(stSecurity)
    stSecurity.bInheritHandle = 0
    ' Set stSecurity.lpSecurityDescriptor to null to overcome problem at RTE on new XP machine - 31 March 2007
    'stSecurity.lpSecurityDescriptor = PROCESS_ALL_ACCESS
    stSecurity.lpSecurityDescriptor = Null
    
    '
    ' Start the shelled application:
    '
    lRet = CreateProcessA(vbNullString, sTask, stSecurity, 0&, 1&, NORMAL_PRIORITY_CLASS + CREATE_SEPARATE_WOW_VDM, 0&, vbNullString, stStart, stProc)
  
    '
    ' Wait for the shelled application to finish
    '
    iFinished = False
    lStartTime = GetTickCount
    Do
        lRet = WaitForSingleObject(stProc.hProcess, 0&)
        If lTimeout <> INFINITE Then
            If GetTickCount < lStartTime Then
                iFinished = (((&H7FFFFFFF - lStartTime) + GetTickCount) > lTimeout)
            Else
                iFinished = ((GetTickCount - lStartTime) > lTimeout)
            End If
        End If
        DoEvents
        Sleep 200
    Loop Until (lRet = 0) Or iFinished Or gbPollForTaskFinish
    
    '
    ' Check if the thing has timed out
    '
    If lRet <> 0 Then
        lRet = WAIT_TIMEOUT
    End If
    Call TerminateProcess(stProc.hProcess, 0&)
    Call CloseHandle(stProc.hProcess)
    

    PSGEN_PollForTask = lRet
    
End Function


Public Function PSGEN_GetTempFilename$(Optional ByVal sExtension$ = "")
'****************************************************************************
'
'     Pivotal Solutions Ltd © 1997
'
'****************************************************************************
'
'                     NAME:   Function PSGEN_GetTempFilename
'
'                             sExtension    - The extension requested
'
'                          ) As String
'
'             DEPENDENCIES:   NONE
'
'     MODIFICATION HISTORY:   Steve O'Hara    04 November 1997   First created for DocBlazer
'
'                  PURPOSE:   Creates a temporary unique filename
'
'****************************************************************************
'
'
Dim sReturn$

    '
    ' Initialise error vector
    ' Get the temp filename
    '
    On Error Resume Next
    sReturn = "ps" + Hex(GetTickCount) + IIf(sExtension = "", ".tmp", "." + sExtension)

    '
    ' Return value to caller
    '
    PSGEN_GetTempFilename = sReturn

End Function

    
    

Public Function PSGEN_GetTempPathFilename$(Optional ByVal sExtension$ = "")
'****************************************************************************
'
'     Pivotal Solutions Ltd © 1997
'
'****************************************************************************
'
'                     NAME:   Function PSGEN_GetTempPathFilename
'
'                             sExtension    - The extension requested
'
'                          ) As String
'
'             DEPENDENCIES:   NONE
'
'     MODIFICATION HISTORY:   Steve O'Hara    04 November 1997   First created for DocBlazer
'
'                  PURPOSE:   Creates a temporary unique filename in temp path
'
'****************************************************************************
'
'
Dim sReturn$, sFile$
Dim sTmp As String * 255
Dim sDir As String * 255

    '
    ' Get the temp filename and clean it up
    '
    On Error Resume Next
    sTmp = String(255, vbNullChar)
    sDir = String(255, vbNullChar)
    Call GetTempPath(Len(sDir), sDir)
    Call GetTempFileName(PSGEN_GetItem(1, vbNullChar, sDir), "ps", 0&, sTmp)
    sReturn = PSGEN_GetItem(1, vbNullChar, sTmp)

    '
    ' If an extension has been provided then kill this file and
    ' create a new one with the required extension
    '
    If sExtension <> "" Then
        sFile = sReturn
        sReturn = PSGEN_GetItem(1, ".tmp", sReturn) + "." + sExtension
        If PSGEN_FileExists(sReturn) Then Kill sReturn
        Name sFile As sReturn
    End If

    '
    ' Return value to caller
    '
    PSGEN_GetTempPathFilename = sReturn

End Function



Public Function PSGEN_GetTempDir$()
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME:   Function PSGEN_GetTempDir
'
'
'                          ) As String
'
'             DEPENDENCIES:   NONE
'
'     MODIFICATION HISTORY:   Steve O'Hara    04 November 1997   First created for DocBlazer
'
'                  PURPOSE:   Returns the system temporary directory
'
'****************************************************************************
'
'
Dim sReturn$
Dim sDir As String * 255

    '
    ' Initialise error vector
    '
    On Error Resume Next

    '
    ' Get the temp filename and clean it up
    '
    sDir = String(255, Chr(0))
    Call GetTempPath(Len(sDir), sDir)
    sReturn = PSGEN_GetItem(1, Chr(0), sDir)

    '
    ' Return value to caller
    '
    PSGEN_GetTempDir = sReturn

End Function

    
    


Function PSGEN_ConvertTime(ByVal vSource As Variant, ByVal iToNumber%) As Variant
'***************************************************************************
'
'       NAME:           Function PSGEN_ConvertTime  (
'
'                       vSource As Variant - Value of time to convert
'                       iToNumber%         - Boolean True=Convert to long, False=convert to string
'
'                       ) As Variant       - Converted value
'
'       DEPENDENCIES:   None
'
'       PURPOSE:        Converts the source into either a long if iToNumber
'                       is true, or converts to a string.  The format for
'                       time as a string is 'hh:mm:ss'.
'                       If the seconds is not specified then it is assumed
'                       to be zero.
'
'***************************************************************************
'
'
Dim vReturn As Variant
Dim sSource$

    '
    ' Determine the action
    '
    If iToNumber Then
        
        '
        ' Check that the string source is in the correct format
        '
        sSource = Trim$(vSource)
        If sSource Like "##:##" Then
            sSource = sSource + "00"
        ElseIf sSource Like "##" Then
            sSource = sSource + "0000"
        End If
        vReturn = Replace(sSource, ":", "")
    Else
        '
        ' Simply do some divisions to get the numbers
        '
        sSource = Format$(Val(vSource), "000000")
        vReturn = Left$(sSource, 2) + ":" + Mid$(sSource, 3, 2) + ":" + Right$(sSource, 2)
    End If
    
    PSGEN_ConvertTime = vReturn
    
End Function



Public Function PSGEN_FormLoaded%(frmLookFor As Form)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME:  Function PSGEN_FormLoaded
'
'                     frmLookFor   - Form to look for
'
'                         ) As Integer
'
'             DEPENDENCIES:  NONE
'
'     MODIFICATION HISTORY:  Steve O'Hara    26 March 1997   First created for RTE
'
'                  PURPOSE:  Determines if the form passsed is loaded in the project.
'
'****************************************************************************
'
'
Dim iReturn%
Dim frmForm As Form


    '
    ' Initialise error vector and loop through all the forms in the collection
    '
    On Error Resume Next
    iReturn = False
    For Each frmForm In Forms
        If frmForm Is frmLookFor Then
            iReturn = True
            Exit For
        End If
    Next frmForm

    '
    ' Return value to caller
    '
    PSGEN_FormLoaded = iReturn

End Function

Sub PSGEN_SetTopMost(ByVal hWnd&)
'*OBJS DESC START***********************************************************
'
'       NAME:           Sub PSGEN_SetTopMost (
'
'                       hWnd As Integer   - Window to make floating
'
'                       )
'
'       DEPENDENCIES:   SDK User DLL
'
'       PURPOSE:        This sub makes the specified window float on top
'                       of ll others.
'
'*OBJS DESC END*************************************************************
'
'
    Call SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)

End Sub


Public Sub PSGEN_UnderConstruction()
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME:  Sub PSGEN_UnderConstruction
'
'
'             DEPENDENCIES:  NONE
'
'     MODIFICATION HISTORY:  Steve O'Hara    26 February 1997   First created for RTE
'
'                  PURPOSE:  A simple routine to act as a common ennunciation
'                            of the fact that a function in the application
'                            may be under construction.
'
'****************************************************************************
'
'

    '
    ' Show a message box telling the user that the system is under construction
    '
    MsgBox "This facility is currently under construction and so is not available.", vbOKOnly + vbExclamation

End Sub

Public Sub PSGEN_SetFormsEnabled(ByVal iValue%)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME:  Sub PSGEN_SetFormsEnabled
'
'                     iValue%            - Flag indicating whether to enable or disable the forms
'
'             DEPENDENCIES:  NONE
'
'     MODIFICATION HISTORY:  Steve O'Hara    21 February 1997   First created for RTE
'
'                  PURPOSE:  Enable/Disable all currently loaded forms within
'                            the app
'
'****************************************************************************
'
'

Dim frmForm As Form

    '
    ' Set the correct mouse indicator
    '
    On Error Resume Next
    If Not App.UnattendedApp Then
        If iValue Then
            Screen.MousePointer = vbDefault
        Else
            Screen.MousePointer = vbHourglass
        End If
    
        '
        ' Now disable/enable all the forms
        '
        For Each frmForm In Forms
            Err = 0
'            If Not frmForm.MDIChild Then
                If (Err = 0) Then frmForm.Enabled = iValue
'            End If
        Next frmForm
    End If
    
End Sub


Function PSGEN_TrimQuoted$(ByVal sInput$)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME:  Function PSGEN_TrimQuoted
'
'                         sInput$      - The string to work on
'
'                         ) As String  - Modified string
'
'
'             DEPENDENCIES:  NONE
'
'     MODIFICATION HISTORY:  Steve O'Hara    11th March 1997   First created for RTE from Desktop
'
'                  PURPOSE:  This routine removes leading and trailing blanks
'                            within quoted strings
'
'****************************************************************************
'
'
Dim bOutput%
Dim bInQuote%
Dim sChar$
Dim sOutput$
Dim iCharPos%
    
    '
    ' First, remove leading and trailing blanks.
    '
    sInput = Trim$(sInput)
    sInput = Replace(sInput, "''", "~~")
    bInQuote = False
    sOutput = ""
    
    '
    ' Remove leading blanks
    '
    For iCharPos = 1 To Len(sInput)
        sChar = Mid$(sInput, iCharPos, 1)
        
        '
        ' Set the flag that indicates whether or not we are within quotes.
        '
        If sChar = "'" Then
            bInQuote = Not bInQuote
        End If
        If Not (bInQuote And sChar = " " And (Right$(sOutput, 1) = "'")) Then
            sOutput = sOutput + sChar
        End If
        
    Next iCharPos

    '
    ' Remove trailing blanks
    '
    sInput = sOutput
    sOutput = ""
    bInQuote = False
    For iCharPos = Len(sInput) To 1 Step -1
        sChar = Mid$(sInput, iCharPos, 1)
        
        '
        ' Set the flag that indicates whether or not we are within quotes.
        '
        If sChar = "'" Then
            bInQuote = Not bInQuote
        End If
        If Not (bInQuote And sChar = " " And (Left$(sOutput, 1) = "'")) Then
            sOutput = sChar + sOutput
        End If
        
    Next iCharPos

    sOutput = Replace(sOutput, "~~", "''")
    PSGEN_TrimQuoted = sOutput

End Function


Function PSGEN_Trim$(ByVal sInput$, Optional ByVal sSep$ = vbCrLf)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME:  Function PSGEN_Trim
'
'                         sInput$      - The string to work on
'                         sSEp$        - Character sequence to trim off
'
'                         ) As String  - Modified string
'
'
'             DEPENDENCIES:  NONE
'
'     MODIFICATION HISTORY:  Steve O'Hara    11th March 1997   First created for RTE from Desktop
'
'                  PURPOSE:  This routine removes leading and trailing sSep
'
'****************************************************************************
'
'
Dim sOutput$
    
    '
    ' First, remove leading and trailing blanks.
    '
    On Error Resume Next
    sOutput = Trim$(sInput)
    If sOutput <> "" And sSep <> "" Then
        While Left$(sOutput, Len(sSep)) = sSep And Err = 0
            If Err = 0 Then sOutput = Trim$(Mid$(sOutput, Len(sSep) + 1))
        Wend
        Err.Clear
        While Right$(sOutput, Len(sSep)) = sSep And Err = 0
            If Err = 0 Then sOutput = Trim$(Left$(sOutput, Len(sOutput) - Len(sSep)))
        Wend
    End If
    
    PSGEN_Trim = sOutput

End Function


Function PSGEN_TrimTrail$(ByVal sInput$, Optional ByVal sSep$ = vbCrLf)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME:  Function PSGEN_TrimTrail
'
'                         sInput$      - The string to work on
'                         sSEp$        - Character sequence to trim off
'
'                         ) As String  - Modified string
'
'
'             DEPENDENCIES:  NONE
'
'     MODIFICATION HISTORY:  Steve O'Hara    11th March 1997   First created for RTE from Desktop
'
'                  PURPOSE:  This routine removes trailing sSep
'
'****************************************************************************
'
'
Dim sOutput$
    
    '
    ' First, remove trailing blanks.
    '
    On Error Resume Next
    sOutput = RTrim$(sInput)
    If sOutput <> "" And sSep <> "" Then
        While Right$(sOutput, Len(sSep)) = sSep And Err = 0
            If Err = 0 Then sOutput = Trim$(Left$(sOutput, Len(sOutput) - Len(sSep)))
        Wend
    End If
    
    PSGEN_TrimTrail = sOutput

End Function


Function PSGEN_RemoveBrackets$(ByVal sInput$)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME:  Function PSGEN_RemoveBrackets
'
'                       sInput&          - The string to work on
'
'                          ) As String   - Modified string
'
'
'             DEPENDENCIES:  NONE
'
'     MODIFICATION HISTORY:  Steve O'Hara    12th March 1997   First created for RTE from Desktop
'
'                  PURPOSE:  This routine removes brackets that are not contained
'                            within single quotes.
'
'****************************************************************************
'
'
Dim bInQuote%
Dim sChar$
Dim sOutput$
Dim iCharPos%
    
    '
    '   First, remove leading and trailing blanks.
    '
    sInput = Trim$(sInput)
    
    '
    '   Convert multiple blanks to a single one.
    '
    For iCharPos = 1 To Len(sInput)
        sChar = Mid$(sInput, iCharPos, 1)
        
        '
        ' Set the flag that indicates whether or not we are within quotes.
        '
        If sChar = "'" Then bInQuote = Not bInQuote
        
        '
        ' Add the character to the output string.
        '
        If bInQuote Or ((sChar <> "(") And (sChar <> ")")) Then
            sOutput = sOutput + sChar
        End If
    
    Next iCharPos

    PSGEN_RemoveBrackets = sOutput

End Function



Function PSGEN_SqueezeNonQuoted$(ByVal sInput$)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME:  Function PSGEN_SqueezeNonQuoted
'
'                       sInput&          - The string to squeeze
'                          ) As String   - Modified string
'
'
'             DEPENDENCIES:  NONE
'
'     MODIFICATION HISTORY:  Steve O'Hara    11th March 1997   First created for RTE from Desktop
'
'                  PURPOSE:  This routine removes leading and trailing blanks
'                            and converts multiple contiguous internal blanks to
'                            a single blank. It retains multiple blanks within
'                            single quotes.
'
'****************************************************************************
'
'
Dim bInQuote%
Dim sChar$
Dim sOutput$
Dim iCharPos%
    
    '
    '   First, remove leading and trailing blanks.
    '
    sInput = Trim$(sInput)
    
    '
    '   Convert multiple blanks to a single one.
    '
    For iCharPos = 1 To Len(sInput)
        sChar = Mid$(sInput, iCharPos, 1)
        
        '
        ' Set the flag that indicates whether or not we are within quotes.
        '
        If sChar = "'" Then bInQuote = Not bInQuote
        
        '
        ' Add the character to the output string.
        '
        If bInQuote Or Not (sChar = " " And Mid$(sInput, iCharPos + 1, 1) = " ") Then
            sOutput = sOutput + sChar
        End If
    
    Next iCharPos

    PSGEN_SqueezeNonQuoted = sOutput

End Function




Sub PSGEN_SetWindowUpdate(ByVal hWnd&, ByVal iUpdate%)
    
    ' TITLE:    BDGN_SetWindowUpdate
    '
    ' PURPOSE:  To turn off/on updating for a window to minimize flicker when items
    '           in the window are being modified/added/removed etc.
    '
    ' USAGE:    BDGN_SetWindowUpdate hWnd, bReadOnly
    '
    '    where:
    '       hWnd       UPDATE  Window handle of control.
    '       bUpdate    READ    Flag indicating whether to turn off or on updating.
    '                          False = turn off; True = turn on
    
    Const WM_SETREDRAW = &HB
    Dim lStatus&
    
    lStatus = SendMessage(hWnd, WM_SETREDRAW, iUpdate, 0)

End Sub


Public Sub PSGEN_TilePicture(ByVal lWindow&, ByVal lPicture&, Optional ByVal lDC&)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME:   Sub GRUGEN_TilePicture
'
'                     lWindow&           - Destination window
'                     lPicture&          - Handle of picture
'
'             DEPENDENCIES:   NONE
'
'     MODIFICATION HISTORY:   Steve O'Hara    16 September 1997   First created for steve
'
'                  PURPOSE:   Tiles a picture across the whole client area of
'                             the specified window.
'
'****************************************************************************
'
'
Dim stBitMap As BITMAP
Dim lDCSrc&, lDCDest&
Dim lBitmapTmp&
Dim stRect As RECT
Dim lRows&, lCols&, lX&, lY&
Dim lRowCnt&, lColCnt&


    '
    ' Initialise error vector
    '
    On Error Resume Next

    '
    ' Get destination rectangle and device context
    '
    Call GetClientRect(lWindow, stRect)
    If lDC = 0 Then
        lDCDest = GetDC(lWindow)
    Else
        lDCDest = lDC
    End If
    
    '
    ' Create source device context and select the bitmap into it
    '
    lDCSrc = CreateCompatibleDC(lDCDest)
    lBitmapTmp = SelectObject(lDCSrc, lPicture)
    
    '
    ' Get the size info about the source bitmap
    '
    Call GetObjectA(lPicture, Len(stBitMap), stBitMap)
    lRows = stRect.Right \ stBitMap.bmWidth
    lCols = stRect.Bottom \ stBitMap.bmHeight
    
    '
    ' Copy the bitmap across the whole window
    '
    For lRowCnt = 0 To lRows
        lX = lRowCnt * stBitMap.bmWidth
        For lColCnt = 0 To lCols
            lY = lColCnt * stBitMap.bmHeight
            Call BitBlt(lDCDest, lX, lY, stBitMap.bmWidth, stBitMap.bmHeight, lDCSrc, 0, 0, SRCCOPY)
        Next lColCnt
    Next lRowCnt
    
    '
    ' Clean up after ourselves
    '
    Call SelectObject(lDCSrc, lBitmapTmp)
    Call DeleteDC(lDCSrc)
    Call ReleaseDC(lWindow, lDCDest)

End Sub

    


Function PSGEN_GetNoOfItems&(ByVal sSep$, ByVal sSource$)
'***************************************************************************
'
'       NAME:           Function PSGEN_GetNoOfItems  (
'
'                       sSep As String     - Item delimiters
'                       sSource As String  - String to search
'
'                       ) As Long       - No. of items in list
'
'       DEPENDENCIES:   None
'
'       PURPOSE:        Returns the item string delimited by the seperator
'                       characters.
'
'***************************************************************************
'
'
Dim lCnt&


    '
    ' Use the VB array functions
    '
    lCnt = UBound(Split(sSource, sSep, -1, vbTextCompare)) + 1
    Err.Clear

    '
    ' Return the modified value
    '
    PSGEN_GetNoOfItems = lCnt


End Function


Public Sub PSGEN_TranslateForm(sTLTFilename$, frmForm As Form)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME:  Sub PSGEN_TranslateForm
'
'                     sTLTFilename$        - Name of the TLT file
'                     frmForm As Form      - Form to translate
'
'             DEPENDENCIES:  NONE
'
'     MODIFICATION HISTORY:  Steve O'Hara    04 February 1997   First created for RTE
'
'                  PURPOSE:  Translates the form using the
'                            project TLT file and the tag
'                            property of the form for the section
'
'****************************************************************************
'
'

Dim ctlControl As Object
Dim sTmp$

    '
    ' Only do something if the tag property of the form contains a value
    '
    If Trim$(frmForm.Tag) <> "" Then
        
        '
        ' Loop round all the controls on the form looking for
        ' ones that will have a caption property
        '
        On Error Resume Next
        For Each ctlControl In frmForm
            
            '
            ' Do a dummy read to see if we get an error about the
            ' fact that a caption property is not supported
            '
            sTmp = ctlControl.Caption
            If Err = 0 Then
                ctlControl.Caption = PSGEN_GetTranslated(sTLTFilename, frmForm.Tag, ctlControl.Caption)
            End If
            Err = 0
            
            '
            ' Same for any tool tips
            '
            sTmp = ctlControl.ToolTipText
            If Err = 0 Then
                ctlControl.ToolTipText = PSGEN_GetTranslated(sTLTFilename, frmForm.Tag, ctlControl.ToolTipText)
            End If
            Err = 0
        Next ctlControl
    End If

End Sub

Private Function Z_GetINIStringSequence$(ByVal sSection$, ByVal sItem$, ByVal sDefault$, ByVal sFilename$)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME:  Function Z_GetINIStringSequence
'
'                     sSection$          - Section name in INI file
'                     sItem$             - Item within the section
'                     sDefault$          - Default value
'                     sFileName$         - The ini file to get the information from
'
'                         ) As String
'
'             DEPENDENCIES:  frmLibraries!lstLibraries
'
'     MODIFICATION HISTORY:  Mark Travis 16 June 1997   First created for RTE
'
'                  PURPOSE: Called from PSGEN_GetINIString when no entry was found.
'                           Some stored INI values exceed the 256 character limit. They
'                           are therefore stored as multiple entries in the ini
'                           file. The item entry is stored as a sequence of entries
'                           in which the original entry has a number added to the
'                           end so the entry name has the syntax 'entry_1'
'
'****************************************************************************
'
'
Const PSGEN_DEFAULT_STRING = ""

Dim sReturn$
Dim sValue As String * 210
Dim sEntry$
Dim iEntry%, iLength%

    On Error Resume Next
    sReturn = ""

    iEntry = 0
    Do
        sEntry = sItem + "_" + Format$(iEntry)
        iLength = GetPrivateProfileString(sSection, sEntry, PSGEN_DEFAULT_STRING, sValue, Len(sValue), sFilename)
        If (iLength <> 0) Then
            If iLength > 4 Then
                If Mid(sValue, iLength - 4, 5) = "'@@@@" Or Mid(sValue, iLength - 4, 5) = " @@@@" Then iLength = iLength - 4
                If Left(sValue, 5) = "@@@@'" Or Left(sValue, 5) = "@@@@ " Then
                    sValue = Right(sValue, Len(sValue) - 4)
                    iLength = iLength - 4
                End If
            End If
            sReturn = sReturn + Left$(sValue, iLength)
            iEntry = iEntry + 1
        End If
    Loop Until (iLength = 0)
    
    If (sReturn = "") Then
        '
        ' No entry was found so return the default value
        '
        sReturn = sDefault
    End If
    
    Z_GetINIStringSequence = sReturn

End Function


Private Function Z_PutINIStringSequence%(ByVal sSection$, ByVal sItem$, ByVal sValue$, ByVal sFilename$)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME:  Function Z_GetINIStringSequence
'
'                     sSection$          - Section name in INI file
'                     sItem$             - Item within the section
'                     sValue$            - Value to store
'                     sFileName$         - The inifile to store the information in
'
'                         ) As String
'
'             DEPENDENCIES:  frmLibraries!lstLibraries
'
'     MODIFICATION HISTORY:  Mark Travis 16 June 1997   First created for RTE
'
'                  PURPOSE: Called from PSGEN_PutINIString when the value to be stored
'                           exceeds the 256 character limit. The value is therefore
'                           stored as multiple entries in the ini file.
'                           The item entry is stored as a sequence of entries
'                           in which the original entry has a number added to the
'                           end so the entry name has the syntax 'entry_1'
'
'****************************************************************************
'
'
Const PSGEN_DEFAULT_STRING = ""
Const WRITE_WIDTH = 200

Dim sSubValue$, sEntry$
Dim iEntry%, iLength%, iTmp%
Dim sTmp As String * 255

    '
    ' Write the value to the ini file
    '
    On Error Resume Next
    iEntry = 0
    If sValue <> "" Then
        Do
            sEntry = sItem + "_" + Format$(iEntry)
            sSubValue = Mid$(sValue, iEntry * WRITE_WIDTH + 1, IIf((iEntry + 1) * WRITE_WIDTH > Len(sValue), Len(sValue) - (iEntry * WRITE_WIDTH), WRITE_WIDTH))
            
            '
            ' Add text to end and start if start or end with a space or single quote
            ' In these cases the characters were getting 'lost'
            '
            If Right(sSubValue, 1) = "'" Or Right(sSubValue, 1) = " " Then sSubValue = sSubValue + "@@@@"
            If Left(sSubValue, 1) = "'" Or Left(sSubValue, 1) = " " Then sSubValue = "@@@@" + sSubValue
            iTmp = WritePrivateProfileString(sSection, sEntry, sSubValue, sFilename)
            If iTmp <> 0 Then iEntry = iEntry + 1
        Loop Until ((iTmp = 0) Or (Len(sSubValue) < WRITE_WIDTH))
        
        '
        ' Ensure there isn't an entry for when the entry was less than 256 characters
        '
        If (GetPrivateProfileString(sSection, sItem, PSGEN_DEFAULT_STRING, sTmp, 255, sFilename) > 0) Then
            Call WritePrivateProfileString(sSection, sItem, 0&, sFilename)
        End If
    End If
    
    '
    ' Delete all other entries that remain from previously stored values
    '
    Do
        sEntry = sItem + "_" + Format$(iEntry)
        iLength = GetPrivateProfileString(sSection, sEntry, PSGEN_DEFAULT_STRING, sTmp, 255, sFilename)
        If iLength <> 0 Then Call WritePrivateProfileString(sSection, sEntry, 0&, sFilename)
        iEntry = iEntry + 1
    Loop Until iLength = 0
    
    Z_PutINIStringSequence = iEntry

End Function


Public Function PSGEN_FindFile$(ByVal sFileToFind$)
'***************************************************************************
'
'       NAME:           Sub PSGEN_FindFile (
'
'                       sFileToFind As String   - File to look for
'
'                       ) As String
'
'       DEPENDENCIES:   NONE
'
'       PURPOSE:        This function attempts to locate the file given
'                       by sFileToFind.  It first looks in the working
'                       directory, then the directory in which the application
'                       started and finally the Windows directory.  If it
'                       cannot be found in any of these, then "" is returned
'                       otherwise the full file spec. is returned.
'
'***************************************************************************
'

Dim sTmp$
Dim sFilename$
Dim sWindowsDir As String * 127
Dim iTmp%

    '
    ' First check if it is in the current directory
    '
    PSGEN_FindFile = ""
    On Error Resume Next
            
    '
    ' Try the current directory
    '
    If Right$(CurDir$, 1) = "\" Then sFilename = CurDir$ + sFileToFind Else sFilename = CurDir$ + "\" + sFileToFind
    sTmp = Dir$(sFilename)
    If (sTmp <> "") And (Err = 0) Then
        PSGEN_FindFile = sFilename
    Else
        
        '
        ' Now try the application directory
        '
        If Right$(App.path, 1) = "\" Then sFilename = App.path + sFileToFind Else sFilename = App.path + "\" + sFileToFind
        sTmp = Dir$(sFilename)
        If (sTmp <> "") And (Err = 0) Then
            PSGEN_FindFile = sFilename
        Else
            
            '
            ' Now try the windows directory
            '
            sWindowsDir = ""
            iTmp = GetWindowsDirectory(sWindowsDir, Len(sWindowsDir))
            sFilename = Trim$(sWindowsDir)
            sFilename = Left$(sFilename, Len(sFilename) - 1)
            If Right$(sFilename, 1) = "\" Then sFilename = sFilename + sFileToFind Else sFilename = sFilename + "\" + sFileToFind
            sTmp = Dir$(sFilename)
            If (sTmp <> "") And (Err = 0) Then
                PSGEN_FindFile = sFilename
            End If
        End If
    End If

End Function



Public Function PSGEN_SystemDirectory$()
'***************************************************************************
'
'       NAME:           Sub PSGEN_SystemDirectory (
'
'                       ) As String
'
'       DEPENDENCIES:   NONE
'
'       PURPOSE:        This function returns the system directory
'
'***************************************************************************
'

Dim iCnt%
Dim sWindowsDir As String * 255

    '
    ' Call the API to get the directory
    '
    On Error Resume Next
    sWindowsDir = ""
    iCnt = GetSystemDirectory(sWindowsDir, Len(sWindowsDir))
    PSGEN_SystemDirectory = Left$(sWindowsDir, iCnt)

End Function




Public Function PSGEN_GetTranslated$(ByVal sTLTFilename$, ByVal sSection$, ByVal sValue$)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME:  Function PSGEN_GetTranslated
'
'                     sTLTFilename$      - Name of the TLT file
'                     sSection$          - TLT Section for translation items
'                     sValue$            - Item to translate
'
'                         ) As String
'
'             DEPENDENCIES:  NONE
'
'     MODIFICATION HISTORY:  Steve O'Hara    04 February 1997   First created for RTE
'
'                  PURPOSE:  Returns the translated string
'                            from within the TLT section defined
'
'****************************************************************************
'
'
Dim sReturn As String * 255
Dim iTmp%

    '
    ' Check the value of the TLT file and use a default if empty
    '
    On Error Resume Next
    iTmp = GetPrivateProfileString(sSection, sValue, sValue, sReturn, Len(sReturn), sTLTFilename)
    If iTmp > 1 Then
        PSGEN_GetTranslated = Left$(sReturn, iTmp)
    Else
        PSGEN_GetTranslated = sValue
    End If
    
End Function

Public Sub PSGEN_PutINIString(ByVal sIniFileName$, ByVal sSection$, ByVal sItem$, ByVal sValue$)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME:  Function PSGEN_PutINIString
'
'                     sINIFilename$      - INI file to use
'                     sSection$          - Section name in INI file
'                     sItem$             - Item within the section
'                     sValue$            - Value of item
'
'             DEPENDENCIES:  frmLibraries!lstLibraries
'
'     MODIFICATION HISTORY:  Steve O'Hara    04 February 1997   First created for RTE
'
'                  PURPOSE:  Sets the item value in the specified section in the
'                            default INI file for the current library.
'                            If the Section name starts with '*' then the library
'                            is not used as a qulifier, the Section name is used
'                            explicitly.
'
'****************************************************************************
'
'
Dim iTmp%
Dim sSectionHeader$
Dim lMutex&

    '
    ' Check the value of the INI file and use a default if empty
    '
    On Error Resume Next
    lMutex = PSGEN_StartMutex(sIniFileName)
    sSectionHeader = Trim$(sSection)
    If Left$(sValue, 1) = "'" Then sValue = "~~" + Mid$(sValue, 2) + "~~"
    If (Len(sValue) <= 255) Then
        If sValue = "" Then
            If (sItem = "") Then
                iTmp = WritePrivateProfileString(sSectionHeader, 0&, 0&, sIniFileName)
            Else
                iTmp = WritePrivateProfileString(sSectionHeader, sItem, 0&, sIniFileName)
            End If
        Else
            If (sItem = "") Then
                iTmp = WritePrivateProfileString(sSectionHeader, 0&, 0&, sIniFileName)
            Else
                iTmp = WritePrivateProfileString(sSectionHeader, sItem, sValue, sIniFileName)
            End If
        End If
        Call Z_PutINIStringSequence(sSectionHeader, sItem, "", sIniFileName)
    Else
        iTmp = Z_PutINIStringSequence(sSectionHeader, sItem, sValue, sIniFileName)
    End If
    Call PSGEN_EndMutex(lMutex)
    
    '
    ' Check to see that the INI file write worked Ok, sometimes it will fail
    ' because it is set read-only.
    '
    If iTmp = 0 Then
        If App.StartMode = vbSModeStandalone Then
            MsgBox "Write to INI file (" + sIniFileName + ") failed", vbOKOnly + vbExclamation
        Else
            Call PSGEN_Log("Write to INI file (" + sIniFileName + ") failed", LogError)
        End If
    End If
    
End Sub


Public Function PSGEN_GetINIString$(ByVal sIniFileName$, ByVal sSection$, ByVal sItem$, ByVal sDefault$)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME:  Function PSGEN_GetINIString
'
'                     sINIFilename$      - INI file to use
'                     sSection$          - Section name in INI file
'                     sItem$             - Item within the section
'                     sDefault$          - Default value
'
'                         ) As String
'
'             DEPENDENCIES:  frmLibraries!lstLibraries
'
'     MODIFICATION HISTORY:  Steve O'Hara    04 February 1997   First created for RTE
'
'                  PURPOSE:  Retrieves the item value from the specified section in the
'                            default INI file for the current open library.  If the
'                            item is not present, then the default value is returned
'                            If the Section name starts with '*' then the library
'                            is not used as a qulifier, the Section name is used
'                            explicitly.
'
'****************************************************************************
'
'
Const PSGEN_DEFAULT_STRING = ""
Dim sValue As String * 32000
Dim sSectionHeader$, sReturn$, sTmp$
Dim iTmp%
Dim lMutex&

    '
    ' Check the value of the INI file and use a default if empty
    '
    On Error Resume Next
    lMutex = PSGEN_StartMutex(sIniFileName)
    sSectionHeader = Trim$(sSection)
    If sSectionHeader = "" Then
        iTmp = GetPrivateProfileString(vbNullString, vbNullString, PSGEN_DEFAULT_STRING, sValue, Len(sValue), sIniFileName)
    Else
        If sItem = "" Then
            iTmp = GetPrivateProfileString(sSectionHeader, vbNullString, PSGEN_DEFAULT_STRING, sValue, Len(sValue), sIniFileName)
        Else
            iTmp = GetPrivateProfileString(sSectionHeader, sItem, PSGEN_DEFAULT_STRING, sValue, Len(sValue), sIniFileName)
        End If
    End If
    
    If iTmp > 0 Then
        sReturn = Left$(sValue, iTmp)
    Else
        sReturn = Z_GetINIStringSequence(sSectionHeader, sItem, sDefault, sIniFileName)
    End If

    If Left$(sReturn, 2) = "~~" And Right$(sReturn, 2) = "~~" Then sReturn = "'" + Mid$(sReturn, 3, Len(sReturn) - 4)
    Call PSGEN_EndMutex(lMutex)

    PSGEN_GetINIString = sReturn
    
End Function

Public Function PSGEN_InsertTextIntoFile%(ByVal sText$, ByVal sIntoFile$, ByVal sHere$, sError$)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME:   Function PSGEN_InsertTextIntoFile
'
'                     sText$              - Text to insert
'                     sIntoFile$          - File to insert into
'                     sHere$              - Replacing this string
'                     sError$             - Errors encounterd
'
'                          ) As Integer
'
'             DEPENDENCIES:   NONE
'
'     MODIFICATION HISTORY:   Steve O'Hara    3rd March 1998     First created for DocBlazer
'
'                  PURPOSE:   Inserts sText into sIntoFile replacing the
'                             string sHere. Returns any errors encountered.
'
'****************************************************************************
'
'
Dim iReturn%, iOutput%
Dim iIntoFile%
Dim sOutput$, sLine$, sTmp$


    '
    ' Initialise error vector and file handle
    '
    On Error GoTo InsertFileError
    sError = ""
    iReturn = False
        
    '
    ' Create a temporary output file
    '
    iOutput = FreeFile
    sOutput = PSGEN_GetTempPathFilename()
    Open sOutput For Output Access Write As #iOutput
    iIntoFile = FreeFile
    Open sIntoFile For Input Access Read As #iIntoFile

    '
    ' Open all the files Ok so read everything up to the marker
    ' into the output file
    '
    While Not EOF(iIntoFile)
        Line Input #iIntoFile, sLine
        
        '
        ' Look to see if the marker is in the line
        '
        If InStr(1, sLine, sHere, vbTextCompare) > 0 Then
            
            '
            ' Now output the whole of the insertion file
            '
            Print #iOutput, PSGEN_GetItem(1, sHere, sLine);
            Print #iOutput, sText;
            Print #iOutput, PSGEN_GetItem(2, sHere, sLine)
        Else
            Print #iOutput, sLine
        End If
    Wend
    Close iOutput
    Close iIntoFile
    
    '
    ' Move the output file to overwrite the old one
    '
    If PSGEN_FileExists(sIntoFile) Then Kill sIntoFile
    Name sOutput As sIntoFile
    
    '
    ' Return value to caller
    '
    PSGEN_InsertTextIntoFile = True
    Exit Function
    
InsertFileError:
    sError = "Problem inserting the text into '" + sIntoFile + "' - " + Err.Description
    Close iOutput
    Close iIntoFile
    If PSGEN_FileExists(sOutput) Then Kill sOutput
    PSGEN_InsertTextIntoFile = False

End Function

Public Function PSGEN_InsertTextFile%(ByVal sFromFile$, ByVal sIntoFile$, ByVal sHere$, sError$)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME:   Function PSGEN_InsertTextFile
'
'                     sFromFile$          - File to insert
'                     sIntoFile$          - File to insert into
'                     sHere$              - Replacing this string
'                     sError$             - Errors encounterd
'
'                          ) As Integer
'
'             DEPENDENCIES:   NONE
'
'     MODIFICATION HISTORY:   Steve O'Hara    3rd March 1998     First created for DocBlazer
'
'                  PURPOSE:   Inserts sFromFile into sIntoFile replacing the
'                             string sHere. Returns any errors encountered.
'
'****************************************************************************
'
'
Dim iReturn%, iOutput%, iFirstLine%
Dim iFromFile%, iIntoFile%
Dim sOutput$, sLine$, sTmp$


    '
    ' Initialise error vector and file handle
    '
    On Error GoTo InsertFileError
    sError = ""
    iReturn = False
        
    '
    ' Create a temporary output file to act as the output
    '
    iOutput = FreeFile
    sOutput = PSGEN_GetTempPathFilename()
    Open sOutput For Output Access Write As #iOutput
    iIntoFile = FreeFile
    Open sIntoFile For Input Access Read As #iIntoFile

    '
    ' Open all the files Ok so read everything up to the marker
    ' into the output file
    '
    While Not EOF(iIntoFile)
        Line Input #iIntoFile, sLine
        
        '
        ' Look to see if the marker is in the line
        '
        If InStr(1, sLine, sHere, vbTextCompare) > 0 Then
            
            '
            ' Now output the whole of the insertion file
            '
            Print #iOutput, PSGEN_GetItem(1, sHere, sLine);
            iFromFile = FreeFile
            Open sFromFile For Input Access Read As #iFromFile
            iFirstLine = True
            While Not EOF(iFromFile)
                Line Input #iFromFile, sTmp
                If Not iFirstLine Then Print #iOutput, ""
                Print #iOutput, sTmp;
                iFirstLine = False
            Wend
            Close iFromFile
            Print #iOutput, PSGEN_GetItem(2, sHere, sLine)
        Else
            Print #iOutput, sLine
        End If
    Wend
    Close iOutput
    Close iIntoFile
    
    '
    ' Move the output file to overwrite the old one
    '
    If PSGEN_FileExists(sIntoFile) Then Kill sIntoFile
    Name sOutput As sIntoFile
    
    '
    ' Return value to caller
    '
    PSGEN_InsertTextFile = True
    Exit Function
    
InsertFileError:
    sError = "Problem inserting the file :'" + sFromFile + "' into '" + sIntoFile + "' - " + Err.Description
    Close iOutput
    Close iIntoFile
    Close iFromFile
    If PSGEN_FileExists(sOutput) Then Kill sOutput
    PSGEN_InsertTextFile = False

End Function


Public Function Z_FolderCallback&(ByVal lHandle&, ByVal lMsg&, ByVal lParam&, ByVal lData&)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME: Function PSGEN_FolderCallback
'
'                     lHandle&           - Dialog handle
'                     lMsg&              - Message ID
'                     lParam&            - Message parameter
'                     lData&             - Message data
'
'                          ) As Integer
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    27 March 2000   First created for PSDeployment
'
'                  PURPOSE: This is called from the shell folder selection
'                           dialog when it is initialised. It is Z_ but public
'                           because it is called from an external DLL and so has
'                           to be public but shouldn't be called by anyone else
'
'****************************************************************************
'
'
Dim lReturn&, lTmp&
Dim sTmp$

    '
    ' Determine what it is we're mean't to do
    '
    On Error Resume Next
    Select Case lMsg
        
        '
        ' Set the first directory
        '
        Case BFFM_INITIALIZED
            If Trim$(msBrowseInitDir) <> "" Then
                sTmp = StrConv(msBrowseInitDir, vbUnicode)
                Call SendMessage(lHandle, BFFM_SETSELECTION, 1, ByVal sTmp)
            End If

        '
        ' Track the selection
        '
        Case BFFM_SELCHANGED
            sTmp = String$(MAX_PATH, 0)
            Call SHGetPathFromIDList(ByVal lParam, ByVal sTmp)
            sTmp = StrConv(sTmp, vbUnicode)
            Call SendMessage(lHandle, BFFM_SETSTATUSTEXT, 0, ByVal sTmp)
            Call SendMessage(lHandle, WM_SETTEXT, 0, ByVal msBrowseTitle)
            
    End Select
    
    '
    ' Return value to caller
    '
    Z_FolderCallback = 0

End Function


Public Function PSGEN_AddressOf&(ByVal lAddress&)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME: Function PSGEN_AddressOf
'
'                     lAddress&          - Address of function
'
'                          ) As Long
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    27 March 2000   First created for PSDeployment
'
'                  PURPOSE: Returns the address of a function by returning
'                           the parameter passed.
'                           The function should be called with the AddressOf
'                           function
'                           e.g. PSGEN_AddressOf(addressof MyFunction)
'
'****************************************************************************
'
'

    '
    ' Initialise error vector
    '
    On Error Resume Next
    PSGEN_AddressOf = lAddress

End Function


Public Function PSGEN_GetControl(ByVal frmForm As Form, ByVal sDataField$) As Control
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME: Function PSGEN_GetControl
'
'                     frmForm As Form           - Form to look through
'                     sDataField$               - Field name
'
'                          ) As Control
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    02 March 2000   First created for ConScriptConsole
'
'                  PURPOSE: Returns the control with the matching data field
'
'****************************************************************************
'
'
Dim ctlReturn As Object
Dim objTmp As Object

    '
    ' Loop round all the controls on the form
    '
    On Error Resume Next
    For Each objTmp In frmForm.Controls
        If StrComp(objTmp.DataField, sDataField, vbTextCompare) = 0 Then
            If Err = 0 Then
                Set ctlReturn = objTmp
                Exit For
            End If
        End If
        Err.Clear
    Next objTmp

    '
    ' Return value to caller
    '
    Set PSGEN_GetControl = ctlReturn

End Function



Function PSGEN_GetItem$(ByVal iItem%, ByVal sSep$, ByVal sSource$)
'*OBJS DESC START***********************************************************
'
'       NAME:           Function PSGEN_GetItem  (
'
'                       iItem As Integer   - Item number to get
'                       sSep As String     - Item delimiters
'                       sSource As String  - String to search
'
'                       ) As String        - Item value
'
'       DEPENDENCIES:   None
'
'       PURPOSE:        Returns the item string delimited by the seperator
'                       characters.
'
'*OBJS DESC END*************************************************************
'
'
Dim sOutput$
Dim iStart%, iFound%, iCnt%


    '
    ' Loop until we have got to the end of the string or the item number
    '
    On Error Resume Next
    sOutput = ""
    iStart = 1
    iCnt = 0
    Do
        iFound = InStr(iStart, sSource, sSep)
        If iFound <> 0 Then
            
            '
            ' Have we found the item number we're interested in
            '
            iCnt = iCnt + 1
            If iCnt = iItem Then
                sOutput = Mid$(sSource, iStart, iFound - iStart)
                iFound = 0
            Else
                iStart = iFound + Len(sSep)
            End If
        Else

            '
            ' Nothing found so either return nothing or the whole thing
            ' if this is the first item we're looking for
            '
            If (iCnt = 0) And (iItem = 1) Then
                sOutput = sSource
            ElseIf iCnt = iItem - 1 Then
                sOutput = Mid$(sSource, iStart, Len(sSource) - iStart + 1)
            Else
                sOutput = ""
            End If
        End If

    Loop Until iFound = 0

    '
    ' Return the modified value
    '
    PSGEN_GetItem = sOutput

End Function


Public Function PSGEN_GetFileAttributes$(ByVal sFile$, ByVal bVerbose As Boolean)
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME: Function PSGEN_GetFileAttributes
'
'                     sFile$             - File to convert
'                     bVerbose           - True for full listing
'
'                          ) As String
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    12 June 2000   First created for BikeSwap
'
'                  PURPOSE: Returns information about the given image file.
'                           This function uses the ImageMagick set of
'                           utilities and so expects the neccersary DLLs and
'                           EXEs to be in the path.
'                           When this is used on Win9x the system doesn't return
'                           the single line non-verbose value.
'
'****************************************************************************
'
'
Dim sReturn$, sFeedback$, sCommand$
Dim sError$, sBatch$, sCmd$, sPath$
    
    '
    ' Build the required command line
    '
    On Error Resume Next
    sPath = Environ(PSGEN_IMAGEMAGICK)
    If sPath = "" Then
        Err.Raise vbObjectError + PSGEN_ERROR_BASE, "PSGEN_GetFileAttributes", "Cannot find the path to image conversion library - missing env. variable " + PSGEN_IMAGEMAGICK
    Else
        sFeedback = PSGEN_GetTempPathFilename
        Kill sFeedback
        sBatch = PSGEN_GetTempPathFilename("bat")
        Kill sBatch
        If PSGEN_IsSystemNT Then
            Call PSGEN_WriteTextFile("cmd.exe /c """ + sPath + "\identify""" + IIf(bVerbose, " -verbose ", " ") + sFile + " >" + sFeedback, sBatch, sError)
            sCmd = "cmd.exe /c " + sBatch
        Else
            Call PSGEN_WriteTextFile("""" + sPath + "\identify""" + IIf(bVerbose, " -verbose ", " ") + sFile + " >" + sFeedback, sBatch, sError)
            sCmd = "command.com /c " + sBatch
        End If
    
        '
        ' Run the command and gather the output
        '
        If PSGEN_WaitForTask(sCmd, "Getting information for " + sFile, True, 20000) <> 0 Then
            Kill sReturn
            Call PSGEN_ReadTextFile(sFeedback, sCommand, sError)
            If sCommand = "" Then
                Err.Raise vbObjectError + PSGEN_ERROR_BASE, "PSGEN_GetFileAttributes", "Conversion taking too long"
            Else
                Err.Raise vbObjectError + PSGEN_ERROR_BASE, "PSGEN_GetFileAttributes", sCommand
            End If
        Else
            '
            ' Check that there was some feedback
            '
            Call PSGEN_ReadTextFile(sFeedback, sReturn, sError)
            If sReturn = "" Then
                Err.Raise vbObjectError + PSGEN_ERROR_BASE, "PSGEN_GetFileAttributes", "No feedback produced"
            End If
        End If
        Kill sFeedback
        Kill sBatch
    End If

    '
    ' Return value to caller
    '
    PSGEN_GetFileAttributes = sReturn

End Function


Public Function PSGEN_GetLoginUsername$()
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME: Function PSGEN_GetLoginUsername
'
'
'                          ) As String
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    06 April 2000   First created for PivotalDesktop
'
'                  PURPOSE: Returns the username of the currently logged in
'                           user from the machine
'
'****************************************************************************
'
'
Const MAX_LEN = 255
Dim sReturn$
Dim lLength&

    '
    ' Initialise the return variable
    '
    On Error Resume Next
    sReturn = String(MAX_LEN, vbNullChar)
    lLength = MAX_LEN
    
    '
    ' Get the username and cut it down to the actual value
    '
    Call GetUserName(sReturn, lLength)
    sReturn = PSGEN_GetItem(1, vbNullChar, sReturn)

    '
    ' Return value to caller
    '
    PSGEN_GetLoginUsername = sReturn

End Function



Public Function PSGEN_GetComputerName$()
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME: Function PSGEN_GetComputerName
'
'
'                          ) As String
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    06 April 2000   First created for PivotalDesktop
'
'                  PURPOSE: Returns the name of the machine
'
'****************************************************************************
'
'
Const MAX_LEN = 255
Dim sReturn$
Dim lLength&

    '
    ' Initialise the return variable
    '
    On Error Resume Next
    sReturn = String(MAX_LEN, vbNullChar)
    lLength = MAX_LEN
    
    '
    ' Get the username and cut it down to the actual value
    '
    Call GetComputerName(sReturn, lLength)
    sReturn = PSGEN_GetItem(1, vbNullChar, sReturn)

    '
    ' Return value to caller
    '
    PSGEN_GetComputerName = sReturn

End Function




Public Function PSGEN_IsSystemNT() As Boolean
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME: Function PSGEN_IsSystemNT
'
'
'                          ) As Boolean
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    16 December 1999   First created for DocBlazer
'
'                  PURPOSE: Returns true if the operating system is NT
'
'****************************************************************************
'
'
Dim bReturn As Boolean
Dim stInfo As OSVERSIONINFO


    '
    ' Initialise error vector and the info structure
    '
    On Error Resume Next
    stInfo.dwOSVersionInfoSize = Len(stInfo)
    Call GetVersionEx(stInfo)
    bReturn = (stInfo.dwPlatformId = VER_PLATFORM_WIN32_NT)

    '
    ' Return value to caller
    '
    PSGEN_IsSystemNT = bReturn

End Function


Public Function PSGEN_CreateFilePath(ByVal sPath$, Optional sError$) As Boolean
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME: Function PSGEN_CreateFilePath
'
'                     sPath$             - Folder path to create
'
'                          ) As Boolean
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    27 March 2000   First created for PSDeployment
'
'                  PURPOSE: Creates the path specified
'                           Returns True if it created the folders Ok
'
'****************************************************************************
'
'
Dim bReturn As Boolean
Dim sDrive$, sBuildPath$
Dim asPaths$()
Dim iCnt%

    '
    ' Determine the drive and pair down the folder
    '
    On Error Resume Next
    If InStr(sPath, ":") = 0 Then
        sDrive = PSGEN_GetItem(1, ":", CurDir)
    Else
        sDrive = PSGEN_GetItem(1, ":", sPath)
        sPath = PSGEN_GetItem(2, ":", sPath)
    End If
    
    '
    ' Recurse through each folder trying to build it
    '
    If sPath = "" Or sPath = "\" Then
        bReturn = (Dir(sDrive + ":", vbDirectory) <> "")
    Else
        asPaths = Split(sPath, "\")
        sBuildPath = sDrive + ":"
        For iCnt = 0 To UBound(asPaths)
            
            '
            ' Build up the path checking if it exists already
            '
            sBuildPath = sBuildPath + IIf(iCnt = 1, "", "\") + asPaths(iCnt)
            If Dir(sBuildPath, vbDirectory) = "" Then
                MkDir sBuildPath
                If Err <> 0 Then
                    sError = "Problem creating file path '" + sBuildPath + "'"
                    Exit For
                End If
            End If
        Next iCnt
        bReturn = (iCnt > UBound(asPaths))
    End If
    
    '
    ' Return value to caller
    '
    PSGEN_CreateFilePath = bReturn

End Function

Public Function PSGEN_GetAPIColor&(ByVal lColor#, Optional ByVal lPal)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME: Function PSGEN_GetAPIColor
'
'                     lColor&            - VB color value
'                     lPal&              - Pallette
'
'                          ) As Long
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    07 May 1999   First created for Willow
'
'                  PURPOSE: Returns a correct API useable color value
'
'****************************************************************************
'
'
Dim lReturn&


    '
    ' Initialise error vector
    '
    On Error Resume Next
    If lColor > 2147483647 Then lColor = &H80000000 + Abs(2147483647 - lColor + 1)
    If lColor < 0 Then
        lReturn = GetSysColor(lColor And &H7FFFFFFF)
    ElseIf Not IsMissing(lPal) Then
        Call OleTranslateColor(lColor, lPal, lReturn)
    Else
        lReturn = lColor
    End If

    '
    ' Return value to caller
    '
    PSGEN_GetAPIColor = lReturn

End Function



Public Function PSGEN_SelectOpenFile$(ByVal lOwner&, ByVal sFilter$, ByVal lFlags&, ByVal sTitle$, Optional lHookProc)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME: Function PSGEN_SelectOpenFile
'
'                     lOwner&            - Owner window handle
'                     sFilter$           - File types
'                     lFlags&            - Selection flags
'                     sTitle$            - Title for the dialog
'                     lHookProc          - Address of a hook procedure
'
'                          ) As String
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    24 January 1999   First created for XList 32
'
'                  PURPOSE: Returns the selected file
'
'****************************************************************************
'
'
Dim stOpenFile As OPENFILENAME
Dim lReturn&
Dim sReturn$


    '
    ' Initialise error vector
    '
    On Error Resume Next
    stOpenFile.lStructSize = Len(stOpenFile)
    stOpenFile.hwndOwner = lOwner
    stOpenFile.hInstance = App.hInstance
    stOpenFile.lpstrFilter = Replace(sFilter, "|", vbNullChar) + vbNullChar + vbNullChar
    stOpenFile.nFilterIndex = 1
    stOpenFile.lpstrFile = String(257, 0)
    stOpenFile.nMaxFile = Len(stOpenFile.lpstrFile) - 1
    stOpenFile.lpstrFileTitle = stOpenFile.lpstrFile
    stOpenFile.nMaxFileTitle = stOpenFile.nMaxFile
    stOpenFile.lpstrInitialDir = CurDir
    stOpenFile.lpstrTitle = sTitle
    stOpenFile.Flags = lFlags Or OFN_HIDEREADONLY
    If Not IsMissing(lHookProc) Then
        stOpenFile.lpfnHook = lHookProc
        stOpenFile.Flags = stOpenFile.Flags Or OFN_ENABLEHOOK Or OFN_EXPLORER
    End If
    lReturn = GetOpenFileName(stOpenFile)
    If lReturn = 0 Then
        sReturn = ""
    Else
        sReturn = PSGEN_GetItem(1, vbNullChar, Trim$(stOpenFile.lpstrFile))
    End If

    '
    ' Return value to caller
    '
    PSGEN_SelectOpenFile = sReturn

End Function


Public Function PSGEN_SelectSaveFile$(ByVal lOwner&, ByVal sFilter$, ByVal lFlags&, ByVal sTitle$)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME: Function PSGEN_SelectSaveFile
'
'                     lOwner&            - Owner window handle
'                     sFilter$           - File types
'                     lFlags&            - Selection flags
'                     sTitle$            - Title for the dialog
'
'                          ) As String
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    24 January 1999   First created for XList 32
'
'                  PURPOSE: Returns the selected file
'
'****************************************************************************
'
'
Dim stOpenFile As OPENFILENAME
Dim lReturn&
Dim sReturn$, sExt$


    '
    ' Initialise error vector
    '
    On Error Resume Next
    stOpenFile.lStructSize = Len(stOpenFile)
    stOpenFile.hwndOwner = lOwner
    stOpenFile.hInstance = App.hInstance
    stOpenFile.lpstrFilter = sFilter
    stOpenFile.nFilterIndex = 1
    stOpenFile.lpstrFile = String(257, 0)
    stOpenFile.nMaxFile = Len(stOpenFile.lpstrFile) - 1
    stOpenFile.lpstrFileTitle = stOpenFile.lpstrFile
    stOpenFile.nMaxFileTitle = stOpenFile.nMaxFile
    stOpenFile.lpstrInitialDir = CurDir
    stOpenFile.lpstrTitle = sTitle
    stOpenFile.Flags = lFlags Or OFN_HIDEREADONLY
    lReturn = GetSaveFileName(stOpenFile)
    If lReturn = 0 Then
        sReturn = ""
    Else
        sReturn = PSGEN_GetItem(1, vbNullChar, Trim$(stOpenFile.lpstrFile))
        
        '
        ' If there is no extension, then get one from the fropdown
        '
        If InStr(sReturn, ".") = 0 Or PSGEN_GetItem(PSGEN_GetNoOfItems(".", sReturn), ".", sReturn) = "" Then
            sExt = PSGEN_GetItem(stOpenFile.nFilterIndex * 2, vbNullChar, sFilter)
            sExt = Trim$(Replace(PSGEN_GetItem(1, ";", sExt), "*", ""))
            sReturn = sReturn + sExt
        End If
    End If

    '
    ' Return value to caller
    '
    PSGEN_SelectSaveFile = sReturn

End Function


Public Function PSGEN_ChooseColor&(ByVal lOwner&, ByVal lStartColor&)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME: Function PSGEN_ChooseColor
'
'                     lOwner&            - Handle of the owner window
'                     lStartColor&       - Starting color
'
'                          ) As Long
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    22 November 1998   First created for XList 32
'
'                  PURPOSE: Chooses a color from the common dialog
'
'****************************************************************************
'
'
Dim lReturn&
Dim abCustomColors() As Byte
Dim iCnt As Integer
Dim stColor As ChooseColor


    '
    ' Initialise error vector and the custom colors
    '
    stColor.rgbResult = lStartColor
    On Error Resume Next
    ReDim abCustomColors(0 To (16 * 4) - 1) As Byte
    For iCnt = LBound(abCustomColors) To UBound(abCustomColors)
        abCustomColors(iCnt) = 0
    Next iCnt

    '
    ' Now setup the structure with the flags
    '
    stColor.lStructSize = Len(stColor)
    stColor.hwndOwner = lOwner
    stColor.hInstance = 0
    stColor.lpCustColors = StrConv(abCustomColors, vbUnicode)
    stColor.Flags = 0
    
    '
    ' Show the dialog and then retrieve the selected and custom colors
    '
    Call ChooseColorAPI(stColor)
    lReturn = stColor.rgbResult
    abCustomColors = StrConv(stColor.lpCustColors, vbFromUnicode)
    
    '
    ' Return value to caller
    '
    PSGEN_ChooseColor = lReturn

End Function


Public Function PSGEN_GetKeysPressed$()
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME: Function PSGEN_GetKeysPressed
'
'                          ) As String
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    29 August 2000   First created for ScreenGrabber
'
'                  PURPOSE: Returns a string representing the keys currently
'                           being held down
'
'****************************************************************************
'
'
Dim abState(255) As Byte
Dim iCnt%
Dim sKey$, sKeys$


    '
    ' Get the key states
    '
    On Error Resume Next
    sKeys = ""
    Call GetKeyboardState(abState(0))
    For iCnt = 0 To 255
        sKey = ""
        If abState(iCnt) >= 128 Then
            If iCnt = VK_CONTROL Then
                sKey = "CTRL"
            
            ElseIf iCnt = VK_SHIFT Then
                sKey = "SHIFT"
            
            ElseIf iCnt = VK_ALTERNATE Then
                sKey = "ALT"
            
            ElseIf iCnt >= VK_F1 And iCnt <= VK_F12 Then
                sKey = "F" + Format$(iCnt - VK_F1 + 1)
            
            ElseIf InStr(1, "`1234567890-=¬!£$%^&*()_+[]{};'#:@~,./<>?\|qwertyuiopasdfghjklzxcvbnm", Chr(iCnt), vbTextCompare) > 0 Then
                sKey = UCase(Chr(iCnt))
            End If
         End If
         If sKey <> "" Then sKeys = sKeys + IIf(sKeys = "", "", " + ") + sKey
    Next iCnt

    '
    ' Return value to caller
    '
    PSGEN_GetKeysPressed = sKeys

End Function

Public Sub PSGEN_DrawRoundedShadow(ByVal frmWork As Form)
'****************************************************************************
'
'     NEWLEAF LTD (C) 2000
'
'****************************************************************************
'
'                     NAME: Sub PSGEN_DrawRoundedShadow
'
'                     frmWork As Form           - Form to round
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    08 January 2000   First created for nlTeamSelector
'
'                  PURPOSE: Draws the shadow effect round rectangles
'
'****************************************************************************
'
'
Const PI_VAL = 3.142

    '
    ' Top Left
    '
    On Error Resume Next
    frmWork.DrawWidth = 2
    frmWork.Circle (10, 10), 9, vbWhite, PI_VAL / 2, PI_VAL * 1.05
    frmWork.DrawWidth = 1
    frmWork.Circle (9, 9), 8, vb3DLight, PI_VAL / 2, PI_VAL

    '
    ' Top right
    '
    frmWork.Circle (frmWork.ScaleWidth - 12, 10), 10, vb3DShadow, 0, PI_VAL * 0.6
    frmWork.Circle (frmWork.ScaleWidth - 13, 10), 10, vb3DShadow, 0, PI_VAL * 0.6

    '
    ' Bottom left
    '
    frmWork.Circle (9, frmWork.ScaleHeight - 11), 9, vb3DLight, PI_VAL, PI_VAL * 1.6
    frmWork.Circle (10, frmWork.ScaleHeight - 12), 10, vb3DLight, PI_VAL, PI_VAL * 1.6
    frmWork.Circle (11, frmWork.ScaleHeight - 14), 12, vb3DShadow, PI_VAL, PI_VAL * 1.6

    '
    ' Bottom right
    '
    frmWork.Circle (frmWork.ScaleWidth - 12, frmWork.ScaleHeight - 12), 9, vb3DShadow, PI_VAL * 1.5, PI_VAL * 1.999
    frmWork.DrawWidth = 2
    frmWork.Circle (frmWork.ScaleWidth - 13, frmWork.ScaleHeight - 12), 11, vb3DShadow, PI_VAL * 1.5, PI_VAL * 1.999
    frmWork.Circle (frmWork.ScaleWidth - 14, frmWork.ScaleHeight - 13), 12, vb3DShadow, PI_VAL * 1.5, PI_VAL * 1.999
    frmWork.DrawWidth = 1

    '
    ' Top Line
    '
    frmWork.DrawWidth = 1
    frmWork.Line (11, 0)-(frmWork.ScaleWidth - 10, 0), vb3DLight
    frmWork.Line (10, 1)-(frmWork.ScaleWidth - 12, 1), vbWhite
    
    '
    ' Bottom Line
    '
    frmWork.Line (11, frmWork.ScaleHeight - 2)-(frmWork.ScaleWidth - 10, frmWork.ScaleHeight - 2), vb3DShadow
    frmWork.Line (9, frmWork.ScaleHeight - 3)-(frmWork.ScaleWidth - 10, frmWork.ScaleHeight - 3), vb3DShadow
    
    '
    ' Left side
    '
    frmWork.Line (0, 11)-(0, frmWork.ScaleHeight - 9), vb3DLight
    frmWork.Line (1, 9)-(1, frmWork.ScaleHeight - 11), vbWhite
    
    '
    ' Right side
    '
    frmWork.Line (frmWork.ScaleWidth - 2, 10)-(frmWork.ScaleWidth - 2, frmWork.ScaleHeight - 10), vb3DShadow
    frmWork.Line (frmWork.ScaleWidth - 3, 10)-(frmWork.ScaleWidth - 3, frmWork.ScaleHeight - 10), vb3DShadow


End Sub


Public Function PSGEN_MontageImage$(ByVal sFile$, ByVal lX&, ByVal lY&, ByVal sAddFile$, Optional ByVal sExtra$ = "")
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME: Function PSGEN_MontageImage
'
'                     sFile$             - File to convert
'                     lX&                - X position
'                     lY&                - Y position
'                     sAddFile$          - File to add
'                     sExtra$            - Extra options to add
'
'                          ) As String
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    12 June 2000   First created for BikeSwap
'
'                  PURPOSE: Montages sAddFile onto sFile at the
'                           location specified.
'                           This functions uses the ImageMagick set of
'                           utilities and so expects the neccersary DLLs and
'                           EXEs to be in the path.
'                           Returns a temporary file if successful otherwise
'                           rasies an error.
'
'****************************************************************************
'
'
Dim sReturn$, sFeedback$, sCommand$, sPath$
Dim sError$, sBatch$, sGeometery$, sCmd$
    
    
    '
    ' Build the required command line
    '
    On Error Resume Next
    sPath = Environ(PSGEN_IMAGEMAGICK)
    If sPath = "" Then
        Err.Raise vbObjectError + PSGEN_ERROR_BASE, "PSGEN_ConvertImage", "Cannot find the path to image conversion library - missing env. variable " + PSGEN_IMAGEMAGICK
    Else
        sGeometery = """image " + Format$(lX) + "," + Format$(lY) + " '" + sAddFile + "'"""
        sFeedback = PSGEN_GetTempPathFilename
        Kill sFeedback
        sBatch = PSGEN_GetTempPathFilename("bat")
        Kill sBatch
        sReturn = PSGEN_GetTempPathFilename(PSGEN_GetItem(2, ".", sFile))
        Kill sReturn
        If PSGEN_IsSystemNT Then
            Call PSGEN_WriteTextFile("cmd.exe /c """ + sPath + "\convert"" -draw " + sGeometery + IIf(sExtra = "", " ", " " + sExtra + " ") + sFile + " " + sReturn + " >" + sFeedback, sBatch, sError)
            sCmd = "cmd.exe /c " + sBatch
        Else
            Call PSGEN_WriteTextFile("""" + sPath + "\convert"" -draw " + sGeometery + IIf(sExtra = "", " ", " " + sExtra + " ") + sFile + " " + sReturn + " >" + sFeedback, sBatch, sError)
            sCmd = "command.com /c " + sBatch
        End If
    
        '
        ' Run the command and gather the output
        '
        If PSGEN_WaitForTask(sCmd, "Montaging " + sFile + " to " + sReturn, True, 20000) <> 0 Then
            Kill sReturn
            Call PSGEN_ReadTextFile(sFeedback, sCommand, sError)
            If sCommand = "" Then
                Err.Raise vbObjectError + PSGEN_ERROR_BASE, "PSGEN_MontageImage", "Conversion taking too long"
            Else
                Err.Raise vbObjectError + PSGEN_ERROR_BASE, "PSGEN_MontageImage", sCommand
            End If
        Else
            '
            ' Check that there was no feedback
            '
            Call PSGEN_ReadTextFile(sFeedback, sCommand, sError)
            If sCommand <> "" Then
                Kill sReturn
                Err.Raise vbObjectError + PSGEN_ERROR_BASE, "PSGEN_MontageImage", sCommand
            End If
        End If
        Kill sFeedback
        Kill sBatch
    End If

    '
    ' Return value to caller
    '
    PSGEN_MontageImage = sReturn

End Function

Public Sub PSGEN_DialANumber(ByVal sNumber$, ByVal iPortNumber%)
'****************************************************************************
'
'     Pivotal LTD (C) 2000
'
'****************************************************************************
'
'                     NAME: Sub PSGEN_DialANumber
'
'                     sNumber$           - Phone number
'                     iPortNumber%       - Comms port to use
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    16 March 2000   First created for GURU
'
'                  PURPOSE: Dials the phone assuming that a modem is attached
'                           to the comms port specified
'
'****************************************************************************
'
'
' The number of seconds to wait for the modem to dial before
' .. resetting the modem. If the phone hangs up prematurely
' .. try increasing this value by small increments.
'
Const WAITSECONDS = 10
Dim bModemCommand() As Byte
Dim sModemCommand$, sTmp$
Dim lOpenPort&, lRetBytes&, lStartTime&
Dim iCnt%


    '
    ' Open the communications port for read/write (&HC0000000).
    ' Must specify existing file (3).
    '
    On Error Resume Next
    lOpenPort = CreateFile("COM" + Format$(iPortNumber), &HC0000000, 0, 0, 3, 0, 0)
    If lOpenPort = -1 Then
        MsgBox "Unable to open communication port COM" + Format$(iPortNumber), vbOKOnly + vbExclamation
    Else
    
        '
        ' Wait 2 seconds for the port to initialise
        '
        lStartTime = Timer
        While Timer < lStartTime + 1
           DoEvents
        Wend
            
        '
        ' Send the telephone number to the modem
        '
        For iCnt = 1 To Len(sNumber)
            Select Case Mid$(sNumber, iCnt, 1)
                Case "0" To "9"
                    sTmp = sTmp + Mid$(sNumber, iCnt, 1) + " "
            End Select
        Next iCnt
        sModemCommand = "ATDT" & sTmp & vbCrLf
        bModemCommand = StrConv(sModemCommand, vbFromUnicode)
        If WriteFile(lOpenPort, bModemCommand(0), Len(sModemCommand), lRetBytes, 0) = 0 Then
            MsgBox "Unable to dial number " & sNumber, vbOKOnly + vbExclamation
        Else
        
            '
            ' Flush the buffer to make sure it actually wrote
            '
            Call FlushFileBuffers(lOpenPort)
            
            '
            ' Wait WAITSECONDS seconds for the phone to dial
            '
            lStartTime = Timer
            While Timer < lStartTime + WAITSECONDS
               DoEvents
            Wend
            
            '
            ' Reset the modem and take it off line
            '
            sModemCommand = "ATH0" & sNumber & vbCrLf
            bModemCommand = StrConv(sModemCommand, vbFromUnicode)
            Call WriteFile(lOpenPort, bModemCommand(0), Len(sModemCommand), lRetBytes, 0)
            
            '
            ' Flush the buffer again and close the communications port
            '
            Call FlushFileBuffers(lOpenPort)
            Call CloseHandle(lOpenPort)
        End If
    End If

End Sub

Public Function PSGEN_IsSameText(ByVal sString1$, ByVal sString2$, Optional ByVal iCompare As VbCompareMethod = vbTextCompare) As Boolean
'****************************************************************************
'
'     Pivotal LTD (C) 2000
'
'****************************************************************************
'
'                     NAME: Function PSGEN_IsSameText
'
'                     sString1           - String to compare
'                     sString2           - String to compare
'                     iCompare           - Type of comparison
'
'                     ) As Boolean
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    16 March 2000   First created for GURU
'
'                  PURPOSE: Compares the two strings and returns true if they
'                           are the same
'
'****************************************************************************
'
  
  On Error Resume Next
  If LenB(sString1) = LenB(sString2) Then
    If iCompare = vbBinaryCompare Then
      PSGEN_IsSameText = (InStrB(1, sString1, sString2, iCompare) <> 0)
    Else
      PSGEN_IsSameText = (StrComp(sString1, sString2, iCompare) = 0)
    End If
  End If

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' CreateBitmapPicture
'    - Creates a bitmap type Picture object from a bitmap and palette
'
' hBmp
'    - Handle to a bitmap
'
' hPal
'    - Handle to a Palette
'    - Can be null if the bitmap doesn't use a palette
'
' Returns
'    - Returns a Picture object containing the bitmap
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Public Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture

        Dim r As Long
         Dim Pic As PICBMP
         ' IPicture requires a reference to "Standard OLE Types"
         Dim IPic As IPicture
         Dim IID_IDispatch As GUID

         ' Fill in with IDispatch Interface ID
         With IID_IDispatch
            .Data1 = &H20400
            .Data4(0) = &HC0
            .Data4(7) = &H46
         End With

         ' Fill Pic with necessary parts
         With Pic
            .size = Len(Pic)          ' Length of structure
            .Type = vbPicTypeBitmap   ' Type of Picture (bitmap)
            .hBmp = hBmp              ' Handle to bitmap
            .hPal = hPal              ' Handle to palette (may be null)
         End With

         ' Create Picture object
         r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)

         ' Return the new Picture object
         Set CreateBitmapPicture = IPic
      End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' CaptureWindow
'    - Captures any portion of a window
'
' hWndSrc
'    - Handle to the window to be captured
'
' Client
'    - If True CaptureWindow captures from the client area of the
'      window
'    - If False CaptureWindow captures from the entire window
'
' LeftSrc, TopSrc, WidthSrc, HeightSrc
'    - Specify the portion of the window to capture
'    - Dimensions need to be specified in pixels
'
' Returns
'    - Returns a Picture object containing a bitmap of the specified
'      portion of the window that was captured
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''
'
   Public Function CaptureWindow(ByVal hWndSrc As Long, _
            ByVal Client As Boolean, ByVal LeftSrc As Long, _
            ByVal TopSrc As Long, ByVal WidthSrc As Long, _
            ByVal HeightSrc As Long) As Picture

            Dim hDCMemory As Long
            Dim hBmp As Long
            Dim hBmpPrev As Long
            Dim r As Long
            Dim hDCSrc As Long
            Dim hPal As Long
            Dim hPalPrev As Long
            Dim RasterCapsScrn As Long
            Dim HasPaletteScrn As Long
            Dim PaletteSizeScrn As Long

         Dim LogPal As LOGPALETTE

         ' Depending on the value of Client get the proper device context
         If Client Then
            hDCSrc = GetDC(hWndSrc) ' Get device context for client area
         Else
            hDCSrc = GetWindowDC(hWndSrc) ' Get device context for entire
                                          ' window
         End If

         ' Create a memory device context for the copy process
         hDCMemory = CreateCompatibleDC(hDCSrc)
         ' Create a bitmap and place it in the memory DC
         hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
         hBmpPrev = SelectObject(hDCMemory, hBmp)

         ' Get screen properties
         RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster
                                                            'capabilities
         HasPaletteScrn = RasterCapsScrn And RC_PALETTE       ' Palette
                                                              'support
         PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of
                                                              ' palette

         ' If the screen has a palette make a copy and realize it
         If HasPaletteScrn And (PaletteSizeScrn = 256) Then
            ' Create a copy of the system palette
            LogPal.palVersion = &H300
            LogPal.palNumEntries = 256
            r = GetSystemPaletteEntries(hDCSrc, 0, 256, _
                LogPal.palPalEntry(0))
            hPal = CreatePalette(LogPal)
            ' Select the new palette into the memory DC and realize it
            hPalPrev = SelectPalette(hDCMemory, hPal, 0)
            r = RealizePalette(hDCMemory)
         End If

         ' Copy the on-screen image into the memory DC
         r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, _
            LeftSrc, TopSrc, vbSrcCopy)

      ' Remove the new copy of the  on-screen image
         hBmp = SelectObject(hDCMemory, hBmpPrev)

         ' If the screen has a palette get back the palette that was
         ' selected in previously
         If HasPaletteScrn And (PaletteSizeScrn = 256) Then
            hPal = SelectPalette(hDCMemory, hPalPrev, 0)
         End If

         ' Release the device context resources back to the system
         r = DeleteDC(hDCMemory)
         r = ReleaseDC(hWndSrc, hDCSrc)

         ' Call CreateBitmapPicture to create a picture object from the
         ' bitmap and palette handles.  Then return the resulting picture
         ' object.
         Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
      End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' CaptureScreen
'    - Captures the entire screen
'
' Returns
'    - Returns a Picture object containing a bitmap of the screen
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Public Function CaptureScreen() As Picture
      Dim hWndScreen As Long

         ' Get a handle to the desktop window
         hWndScreen = GetDesktopWindow()

         ' Call CaptureWindow to capture the entire desktop give the handle
         ' and return the resulting Picture object

         Set CaptureScreen = CaptureWindow(hWndScreen, False, 0, 0, _
            Screen.Width \ Screen.TwipsPerPixelX, _
            Screen.Height \ Screen.TwipsPerPixelY)
      End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' CaptureForm
'    - Captures an entire form including title bar and border
'
' frmSrc
'    - The Form object to capture
'
' Returns
'    - Returns a Picture object containing a bitmap of the entire
'      form
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Public Function CaptureForm(frmSrc As Form) As Picture
   ' Call CaptureWindow to capture the entire form given it's window
   ' handle and then return the resulting Picture object
   Set CaptureForm = CaptureWindow(frmSrc.hWnd, False, 0, 0, _
      frmSrc.ScaleX(frmSrc.Width, vbTwips, vbPixels), _
      frmSrc.ScaleY(frmSrc.Height, vbTwips, vbPixels))
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' CaptureClient
'    - Captures the client area of a form
'
' frmSrc
'    - The Form object to capture
'
' Returns
'    - Returns a Picture object containing a bitmap of the form's
' client area
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Public Function CaptureClient(frmSrc As Form) As Picture
   ' Call CaptureWindow to capture the client area of the form given
   ' it's window handle and return the resulting Picture object
   Set CaptureClient = CaptureWindow(frmSrc.hWnd, True, 0, 0, _
      frmSrc.ScaleX(frmSrc.ScaleWidth, frmSrc.ScaleMode, vbPixels), _
      frmSrc.ScaleY(frmSrc.ScaleHeight, frmSrc.ScaleMode, vbPixels))
End Function


Public Function PSGEN_SortArray(ByVal vSource As Variant) As Variant
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2001
'
'****************************************************************************
'
'                     NAME: Function PSGEN_SortArray
'
'                     vSource As variant           - Array to sort
'
'                          ) As Long
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    22 November 2001   First created for PivotalExtendedControls
'
'                  PURPOSE: Returns a variant array which is a sorted version
'                           of the variant array passed
'                           Expects the array to be 2 dimensions, first
'                           dimension the sort and the second the item data
'
'****************************************************************************
'
'
Dim vReturn As Variant
Dim vTmpVal1 As Variant
Dim sTmpVal0$
Dim lCnt&, lGapSize&, lCurPos&, lTmp&
Dim lFirstRow&, lLastRow&, lNumRows&, lSecondDim&
Dim bNumeric As Boolean
Dim bIsObject As Boolean
    
    
    '
    ' Determine if the sort should be numeric
    '
    vReturn = vSource
    lFirstRow = LBound(vReturn)
    lLastRow = UBound(vReturn)
    bNumeric = True
    bIsObject = IsObject(vSource(lCnt, 1))
    For lCnt = lFirstRow To lLastRow
        If Not IsNumeric(vSource(lCnt, 0)) Then
            bNumeric = False
            Exit For
        End If
    Next lCnt
    
    lSecondDim = UBound(vReturn, 2)
    ReDim vTmpVal1(lSecondDim)
    
    '
    ' Determine the optimum sort start point
    '
    lNumRows = lLastRow - lFirstRow + 1
    Do
      lGapSize = lGapSize * 3 + 1
    Loop Until lGapSize > lNumRows
    
    '
    ' Keep going until the gap is closed
    '
    Do
        lGapSize = lGapSize \ 3
      
        '
        ' Loop round each of the elements flipping their contents
        '
        For lCnt = (lGapSize + lFirstRow) To lLastRow
            lCurPos = lCnt
            sTmpVal0 = vReturn(lCnt, 0)
            For lTmp = 1 To lSecondDim
                If bIsObject Then
                    Set vTmpVal1(lTmp) = vReturn(lCnt, lTmp)
                Else
                    vTmpVal1(lTmp) = vReturn(lCnt, lTmp)
                End If
            Next lTmp
            
            If bNumeric Then
                '
                ' Keep flipping until we find a value that is greater
                '
                Do While Val(vReturn(lCurPos - lGapSize, 0)) > Val(sTmpVal0)
                    vReturn(lCurPos, 0) = vReturn(lCurPos - lGapSize, 0)
                    If bIsObject Then
                        Set vReturn(lCurPos, 1) = vReturn(lCurPos - lGapSize, 1)
                    Else
                        vReturn(lCurPos, 1) = vReturn(lCurPos - lGapSize, 1)
                    End If
                    lCurPos = lCurPos - lGapSize
                    If (lCurPos - lGapSize) < lFirstRow Then Exit Do
                Loop
            Else
                '
                ' Keep flipping until we find a value that is greater
                '
                Do While StrComp(vReturn(lCurPos - lGapSize, 0), sTmpVal0, vbTextCompare) > 0
                    vReturn(lCurPos, 0) = vReturn(lCurPos - lGapSize, 0)
                    For lTmp = 1 To lSecondDim
                        If bIsObject Then
                            Set vReturn(lCurPos, lTmp) = vReturn(lCurPos - lGapSize, lTmp)
                        Else
                            vReturn(lCurPos, lTmp) = vReturn(lCurPos - lGapSize, lTmp)
                        End If
                    Next lTmp
                    lCurPos = lCurPos - lGapSize
                    If (lCurPos - lGapSize) < lFirstRow Then Exit Do
                Loop
            End If
            vReturn(lCurPos, 0) = sTmpVal0
            For lTmp = 1 To lSecondDim
                If bIsObject Then
                    Set vReturn(lCurPos, lTmp) = vTmpVal1(lTmp)
                Else
                    vReturn(lCurPos, lTmp) = vTmpVal1(lTmp)
                End If
            Next lTmp
        Next lCnt
    Loop Until lGapSize = 1
    
    '
    ' Return the sorted array
    '
    PSGEN_SortArray = vReturn

End Function


Public Function PSGEN_SortArraySimple(ByVal vSource As Variant) As Variant
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2001
'
'****************************************************************************
'
'                     NAME: Function PSGEN_SortArraySimple
'
'                     vSource As variant           - Array to sort
'
'                          ) As Long
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    22 November 2001   First created for PivotalExtendedControls
'
'                  PURPOSE: Returns a variant array which is a sorted version
'                           of the variant array passed
'                           Expects the array to be single dimension
'
'****************************************************************************
'
'
Dim vReturn As Variant
Dim vTmpVal As Variant
Dim lCnt&, lGapSize&, lCurPos&
Dim lFirstRow&, lLastRow&, lNumRows&
Dim bNumeric As Boolean
    
    
    '
    ' Determine if the sort should be numeric
    '
    vReturn = vSource
    lFirstRow = LBound(vReturn)
    lLastRow = UBound(vReturn)
    bNumeric = True
    For lCnt = lFirstRow To lLastRow
        If Not IsNumeric(vSource(lCnt)) Then
            bNumeric = False
            Exit For
        End If
    Next lCnt
    
    '
    ' Determine the optimum sort start point
    '
    lNumRows = lLastRow - lFirstRow + 1
    Do
      lGapSize = lGapSize * 3 + 1
    Loop Until lGapSize > lNumRows
    
    '
    ' Keep going until the gap is closed
    '
    Do
        lGapSize = lGapSize \ 3
      
        '
        ' Loop round each of the elements flipping their contents
        '
        For lCnt = (lGapSize + lFirstRow) To lLastRow
            lCurPos = lCnt
            vTmpVal = vReturn(lCnt)
            
            If bNumeric Then
                '
                ' Keep flipping until we find a value that is greater
                '
                Do While Val(vReturn(lCurPos - lGapSize)) > Val(vTmpVal)
                    vReturn(lCurPos) = vReturn(lCurPos - lGapSize)
                    lCurPos = lCurPos - lGapSize
                    If (lCurPos - lGapSize) < lFirstRow Then Exit Do
                Loop
            Else
                '
                ' Keep flipping until we find a value that is greater
                '
                Do While StrComp(vReturn(lCurPos - lGapSize), vTmpVal, vbTextCompare) > 0
                    vReturn(lCurPos) = vReturn(lCurPos - lGapSize)
                    lCurPos = lCurPos - lGapSize
                    If (lCurPos - lGapSize) < lFirstRow Then Exit Do
                Loop
            End If
            vReturn(lCurPos) = vTmpVal
        Next lCnt
    Loop Until lGapSize = 1 Or lLastRow < 0
    
    '
    ' Return the sorted array
    '
    PSGEN_SortArraySimple = vReturn

End Function


Public Sub PSGEN_DrawTransparentPicture(dest As PictureBox, ByVal srcBmp&, ByVal DestX&, ByVal DestY&, ByVal TransColor&)
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2001
'
'****************************************************************************
'
'                     NAME: Sub PSGEN_DrawTransparentPicture
'
'                        lDest        - ldc of the device context to paint the picture on
'                        picSource    - Picture to paint
'                         lLeft        - X coordinate of the upper left corner of the area that the
'                                         picture is to be painted on. (in pixels)
'                        lTop         - Y coordinate of the upper left corner of the area that the
'                                         picture is to be painted on. (in pixels)
'                        lWidth       - Width of picture area to paint in pixels
'                        lHeight      - Height of picture area to paint in pixels
'                        lBackColor   - Is the backcolor of the ldc that the image will be painted on
'                        lMaskColor   - Color to mask, must be a valid HCOLORREF
'                        lPal         - Must be a valid HPALETTE
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    08 February 1999   First created for GURU
'
'                  PURPOSE: Draws a bitmap or icon to a ldc, applying a disabled or embossed
'                           look to the picture.  If the DrawState API is available it will
'                           be used else, the effect will be mimicked.  If the passed picture
'                           is a bitmap a mask color can be provided to make the areas of the
'                           picture that have that color transparent
'
'****************************************************************************
'
'

Const PIXEL = 3
Dim destScale As Integer
Dim srcDC As Long  'source bitmap (color)
Dim saveDC As Long 'backup copy of source bitmap
Dim maskDC As Long 'mask bitmap (monochrome)
Dim invDC As Long  'inverse of mask bitmap (monochrome)
Dim resultDC As Long 'combination of source bitmap & background
Dim bmp As BITMAP 'description of the source bitmap
Dim hResultBmp As Long 'Bitmap combination of source & background
Dim hSaveBmp As Long 'Bitmap stores backup copy of source bitmap
Dim hMaskBmp As Long 'Bitmap stores mask (monochrome)
Dim hInvBmp As Long  'Bitmap holds inverse of mask (monochrome)
Dim hPrevBmp As Long 'Bitmap holds previous bitmap selected in DC
Dim hSrcPrevBmp As Long  'Holds previous bitmap in source DC
Dim hSavePrevBmp As Long 'Holds previous bitmap in saved DC
Dim hDestPrevBmp As Long 'Holds previous bitmap in destination DC
Dim hMaskPrevBmp As Long 'Holds previous bitmap in the mask DC
Dim hInvPrevBmp As Long 'Holds previous bitmap in inverted mask DC
Dim OrigColor As Long 'Holds original background color from source DC
Dim Success As Long 'Stores result of call to Windows API

        
    '
    ' Initialise the DCs
    '
    destScale = dest.ScaleMode 'Store ScaleMode to restore later
    dest.ScaleMode = PIXEL 'Set ScaleMode to pixels for Windows GDI
    'Retrieve bitmap to get width (bmp.bmWidth) & height (bmp.bmHeight)

    Call GetObjectA(srcBmp, Len(bmp), bmp)
    srcDC = CreateCompatibleDC(dest.hDC)    'Create DC to hold stage
    saveDC = CreateCompatibleDC(dest.hDC)   'Create DC to hold stage
    maskDC = CreateCompatibleDC(dest.hDC)   'Create DC to hold stage
    invDC = CreateCompatibleDC(dest.hDC)    'Create DC to hold stage
    resultDC = CreateCompatibleDC(dest.hDC) 'Create DC to hold stage

    'Create monochrome bitmaps for the mask-related bitmaps:
    hMaskBmp = CreateBitmap(bmp.bmWidth, bmp.bmHeight, 1, 1, ByVal 0&)
    hInvBmp = CreateBitmap(bmp.bmWidth, bmp.bmHeight, 1, 1, ByVal 0&)
    'Create color bitmaps for final result & stored copy of source
    hResultBmp = CreateCompatibleBitmap(dest.hDC, bmp.bmWidth, bmp.bmHeight)
    hSaveBmp = CreateCompatibleBitmap(dest.hDC, bmp.bmWidth, bmp.bmHeight)
    hSrcPrevBmp = SelectObject(srcDC, srcBmp)     'Select bitmap in DC
    hSavePrevBmp = SelectObject(saveDC, hSaveBmp) 'Select bitmap in DC
    hMaskPrevBmp = SelectObject(maskDC, hMaskBmp) 'Select bitmap in DC
    hInvPrevBmp = SelectObject(invDC, hInvBmp)    'Select bitmap in DC
    hDestPrevBmp = SelectObject(resultDC, hResultBmp) 'Select bitmap
    Call BitBlt(saveDC, 0, 0, bmp.bmWidth, bmp.bmHeight, srcDC, 0, 0, SRCCOPY) 'Make backup of source bitmap to restore later
    'Create mask: set background color of source to transparent color.
    OrigColor = SetBkColor(srcDC, TransColor)
    Call BitBlt(maskDC, 0, 0, bmp.bmWidth, bmp.bmHeight, srcDC, 0, 0, SRCCOPY)
    TransColor = SetBkColor(srcDC, OrigColor)
    'Create inverse of mask to AND w/ source & combine w/ background.
    Call BitBlt(invDC, 0, 0, bmp.bmWidth, bmp.bmHeight, maskDC, 0, 0, NOTSRCCOPY)
    'Copy background bitmap to result & create final transparent bitmap
    Call BitBlt(resultDC, 0, 0, bmp.bmWidth, bmp.bmHeight, dest.hDC, DestX, DestY, SRCCOPY)
    'AND mask bitmap w/ result DC to punch hole in the background by
    'painting black area for non-transparent portion of source bitmap.
    Call BitBlt(resultDC, 0, 0, bmp.bmWidth, bmp.bmHeight, maskDC, 0, 0, SRCAND)
    'AND inverse mask w/ source bitmap to turn off bits associated
    'with transparent area of source bitmap by making it black.
    Call BitBlt(srcDC, 0, 0, bmp.bmWidth, bmp.bmHeight, invDC, 0, 0, SRCAND)
    'XOR result w/ source bitmap to make background show through.
    Call BitBlt(resultDC, 0, 0, bmp.bmWidth, bmp.bmHeight, srcDC, 0, 0, SRCPAINT)
    Call BitBlt(dest.hDC, DestX, DestY, bmp.bmWidth, bmp.bmHeight, resultDC, 0, 0, SRCCOPY)           'Display transparent bitmap on backgrnd
    Call BitBlt(srcDC, 0, 0, bmp.bmWidth, bmp.bmHeight, saveDC, 0, 0, SRCCOPY)           'Restore backup of bitmap.
    hPrevBmp = SelectObject(srcDC, hSrcPrevBmp) 'Select orig object
    hPrevBmp = SelectObject(saveDC, hSavePrevBmp) 'Select orig object
    hPrevBmp = SelectObject(resultDC, hDestPrevBmp) 'Select orig object
    hPrevBmp = SelectObject(maskDC, hMaskPrevBmp) 'Select orig object
    hPrevBmp = SelectObject(invDC, hInvPrevBmp) 'Select orig object
    Call DeleteObject(hSaveBmp)   'Deallocate system resources.
    Call DeleteObject(hMaskBmp)   'Deallocate system resources.
    Call DeleteObject(hInvBmp)    'Deallocate system resources.
    Call DeleteObject(hResultBmp) 'Deallocate system resources.
    Call DeleteDC(srcDC)          'Deallocate system resources.
    Call DeleteDC(saveDC)         'Deallocate system resources.
    Call DeleteDC(invDC)          'Deallocate system resources.
    Call DeleteDC(maskDC)         'Deallocate system resources.
    Call DeleteDC(resultDC)       'Deallocate system resources.
    dest.ScaleMode = destScale 'Restore ScaleMode of destination.

End Sub




Public Function PSGEN_Win32TimetoVbDate(ByVal rTime As Currency) As Date
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2001
'
'****************************************************************************
'
'                     NAME: Sub PSGEN_Win32TimetoVbDate
'
'                        rTime#       - Converts Win32 file time (WIN32_FIND_DATA) to VB
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    08 February 1999   First created for GURU
'
'                  PURPOSE: Converts the file time and dates you get from Win32
'                           into a standard VB date
'
'****************************************************************************
'
'

' Difference between day zero for VB dates and Win32 dates
' (or #12-30-1899# - #01-01-1601#)

Const DAY_ZERO_BIAS As Double = 109205#

' Abs(CDbl(#01-01-1601#))
' 10000000 nanoseconds * 60 seconds * 60 minutes * 24 hours / 10000
' comes to 86400000 (the 10000 adjusts for fixed point in Currency)

Const MILLESECONDS_PER_DAY As Double = 10000000# * 60# * 60# * 24# / 10000#

Dim rAdjustedTime As Currency

    
    '
    ' Call API to convert from UTC time to local time
    '
    On Error Resume Next
    If FileTimeToLocalFileTime(rTime, rAdjustedTime) Then
        
        ' Local time is nanoseconds since 01-01-1601
        ' In Currency that comes out as milliseconds
        ' Divide by milliseconds per day to get days since 1601
        ' Subtract days from 1601 to 1899 to get VB Date equivalent
        PSGEN_Win32TimetoVbDate = CDate((rAdjustedTime / MILLESECONDS_PER_DAY) - DAY_ZERO_BIAS)
    End If

End Function

Public Function PSGEN_Log(ByVal sMessage$, Optional ByVal iLogType As LogEventTypes = LogEventTypes.LogInformation, Optional ByVal sSource$ = "") As Boolean
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2001
'
'****************************************************************************
'
'                     NAME: Sub PSGEN_Log
'
'                        iLogType%    - Type of event to register
'                        sMessage$    - Message to register
'                        sSource$     - Optional source name
'
'                        ) as Boolean  - Returns true if successful
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    08 February 2003   First created for ASPHelper
'
'                  PURPOSE: Writes a message to the NT Event log.
'                           If sSource is empty, then the App.Title value is used
'
'****************************************************************************
'
'

Dim lEventID&
Dim hEventLog&


    '
    ' Open the event log on this machine
    '
    If sSource = "" Then sSource = App.Title
    hEventLog = RegisterEventSource(vbNullString, sSource)

    '
    ' ReportEvent returns 0 if failed, any other number indicates success
    '
    If ReportEvent(hEventLog, iLogType, 0, 0, 0, 1, Len(sMessage), sMessage, 0) = 0 Then
        PSGEN_Log = False
    Else
        PSGEN_Log = True
    End If

    '
    ' Free the resources
    '
    Call DeregisterEventSource(hEventLog)
    
    Debug.Print Format(Now(), "dd/MM/yy hh:mm:ss") + " " + sMessage

End Function

Public Function Z_DialogCallback&(ByVal lHandle&, ByVal lMsg&, ByVal lParam&, ByVal lData&)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME: Function Z_DialogCallback
'
'                     lHandle&           - Dialog handle
'                     lMsg&              - Message ID
'                     lParam&            - Message parameter
'                     lData&             - Message data
'
'                          ) As Integer
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    27 March 2000   First created for PSDeployment
'
'                  PURPOSE: This is called from the shell folder selection
'                           dialog when it is initialised. It is Z_ but public
'                           because it is called from an external DLL and so has
'                           to be public but shouldn't be called by anyone else
'
'****************************************************************************
'
'
Const WM_TIMER = &H113

Dim lReturn&, lTmp&
Dim stRect As RECT

    '
    ' Determine what it is we're mean't to do
    '
    On Error Resume Next
    Select Case lMsg
        
        '
        ' Used by the Open or Save As dialog
        ' Centre the window and bring it to the front
        '
        Case WM_INITDIALOG
            lTmp = GetParent(lHandle)
            Call GetWindowRect(lTmp, stRect)
            Call SetWindowPos(lTmp, HWND_TOPMOST, ((Screen.Width / Screen.TwipsPerPixelX) - (stRect.Right - stRect.Left)) / 2, ((Screen.Height / Screen.TwipsPerPixelY) - (stRect.Bottom - stRect.Top)) / 2, stRect.Right - stRect.Left, stRect.Bottom - stRect.Top, 0)

    End Select
    
    '
    ' Return value to caller
    '
    Z_DialogCallback = 0

End Function

Public Function PSGEN_EncodeToBase64String(DecryptedText As String) As String
  
  Dim c1 As Integer, c2 As Integer, c3 As Integer
  Dim w1 As Integer
  Dim w2 As Integer
  Dim w3 As Integer
  Dim w4 As Integer
  Dim n As Long
  Dim m As Long
  Dim retry As String
  Dim nLength As Long
  Dim arString() As Byte
  Dim arResult() As Byte
  ReDim arResult((Len(DecryptedText) * 8) \ 3 + 6)
  On Error Resume Next
  arString() = DecryptedText
  nLength = Len(DecryptedText)
   For n = 1 To nLength Step 3
    c1 = arString((2 * n) - 2)
    c2 = arString(2 * n)
    If 2 * (n + 1) < UBound(arString) Then
        c3 = arString(2 * (n + 1))
    Else
        c3 = 0
    End If
    arResult(m) = Asc(Z_MimeEncode(c1 \ 4))
    m = m + 2
    arResult(m) = Asc(Z_MimeEncode((c1 And 3) * 16 + c2 \ 16))
    m = m + 2
    If nLength >= n + 1 Then
        arResult(m) = Asc(Z_MimeEncode((c2 And 15) * 4 + c3 \ 64))
        m = m + 2
    Else
        arResult(m) = 0
        m = m + 2
    End If
    If nLength >= n + 2 Then
        arResult(m) = Asc(Z_MimeEncode(c3 And 63))
        m = m + 2
    Else
        arResult(m) = 0
        m = m + 2
    End If
    
 Next
 
 PSGEN_EncodeToBase64String = arResult

End Function



Private Function Z_MimeEncode(w As Integer) As String

Const BASE64CHR = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

    If w >= 0 Then Z_MimeEncode = Mid$(BASE64CHR, w + 1, 1) Else Z_MimeEncode = ""

End Function


Public Function PSIMAGE_SaveFileToGIF(ByVal sInFile$, ByVal sGifFile$, Optional hDC&, Optional bUseTrans As Boolean, Optional ByVal lTransColor&) As Boolean
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2004
'
'****************************************************************************
'
'                     NAME: Function PSIMAGE_SaveFileToGIF
'
'                     sInFile$           - Input file
'                     sGifFile$          - GIF File to save to
'                     hDC&               - Optional context to use for colors etc
'                     bUseTrans          - Optional instruction to create transparent GIF
'                     lTransColor&       - Optional color to make transparent
'
'                          ) As Boolean
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    19 March 2004   First created for MediaServices
'
'                  PURPOSE: Saves a filename to GIF
'
'****************************************************************************
'
'
Dim bReturn As Boolean
Dim objPicture As New StdPicture


    '
    ' Load the source picture
    '
    On Error Resume Next
    Set objPicture = LoadPicture(sInFile)
    If Err = 0 Then bReturn = PSGEN_SaveGIF(objPicture, sGifFile, hDC, bUseTrans, lTransColor)

    '
    ' Return value to caller
    '
    PSIMAGE_SaveFileToGIF = bReturn

End Function



Public Function PSGEN_SaveGIF(ByVal objPic As StdPicture, ByVal sFilename$, Optional hDC&, Optional bUseTrans As Boolean, Optional ByVal lTransColor&) As Boolean
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2004
'
'****************************************************************************
'
'                     NAME: Function PSGEN_SaveGIF
'
'                     objPic             - Picture object to use
'                     sGifFile$          - GIF File to save to
'                     hDC&               - Optional context to use for colors etc
'                     bUseTrans          - Optional instruction to create transparent GIF
'                     lTransColor&       - Optional color to make transparent
'
'                          ) As Boolean
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    19 March 2004   First created for MediaServices
'
'                  PURPOSE: Saves a filename to GIF
'
'****************************************************************************
'
'
   
Dim stScreen As GifScreenDescriptor
Dim stImage As GifImageDescriptor
Dim stBitmapInfo As BITMAPINFO256
Dim stBitMap As BITMAP
Dim lDCScn&, lOldObj&, lSrchDc&
Dim lDib256&, lDC256&, lOldObj256&
Dim abBuf() As Byte
Dim bData As Byte
Dim bTransIndex As Byte
Dim lRow&, lCol&, lColor&
Dim bFound As Boolean
Dim iCode%, iCount%, iOutFile%, iCodeCount%
Dim iBitPosition%
Dim sPrefix$, sByte$
Dim objTempPic As StdPicture
Dim objColorTable As New Collection
Dim lPicWidth&, lPicHeight&
Dim astGifPalette(0 To 255) As RGBTRIPLE
Dim abDataBuffer(255) As Byte


    '
    ' Get image size and allocate buffer memory
    '
    Call GetObjectAPI(objPic, Len(stBitMap), stBitMap)
    lPicWidth = stBitMap.bmWidth
    lPicHeight = stBitMap.bmHeight
    ReDim buf(CLng(((lPicWidth + 3) \ 4) * 4), lPicHeight) As Byte

    '
    ' Prepare DC for paintings
    '
    lDCScn = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
    lDC256 = CreateCompatibleDC(lDCScn)
    If hDC = 0 Then
        lSrchDc = CreateCompatibleDC(lDCScn)
        lOldObj = SelectObject(lSrchDc, objPic)
    Else
        lSrchDc = hDC
    End If
    DeleteDC lDCScn

    '
    ' Since GIF works only with 256 colors, reduce color depth to 256
    ' This sample use simpliest HalfTone palette to reduce color depth
    ' If you want advanced color manipulation with web-safe palettes or
    ' optimal palette with the specified number of colors using octree
    ' quantisation, visit http://vbaccelerator.com/codelib/gfx/octree.htm
    '

    If stBitMap.bmBitsPixel <> 8 Then lDib256 = Z_CreateDib256(lDC256, stBitmapInfo, lPicWidth, lPicHeight)
    If lDib256 <> 0 Then
        lOldObj256 = SelectObject(lDC256, lDib256)
        Call BitBlt(lDC256, 0, 0, lPicWidth, lPicHeight, lSrchDc, 0, 0, vbSrcCopy)
        For lRow = 0 To lPicHeight - 1
            Call GetDIBits(lDC256, lDib256, lRow, 1, buf(0, lPicHeight - lRow), stBitmapInfo, 0)
        Next
    Else
        With stBitmapInfo.bmiHeader
        .biSize = Len(stBitmapInfo.bmiHeader)
        .biWidth = lPicWidth
        .biHeight = lPicHeight
        .biPlanes = 1
        .biBitCount = 8
        .biCompression = BI_RGB
        End With
        For lRow = 0 To lPicHeight - 1
            Call GetDIBits(lSrchDc, objPic, lRow, 1, buf(0, lPicHeight - lRow), stBitmapInfo, 0)
        Next
    End If

    '
    ' Fill gif file info
    '
    For lRow = 0 To 255
        astGifPalette(lRow).rgbBlue = stBitmapInfo.bmiColors(lRow).rgbBlue
        astGifPalette(lRow).rgbGreen = stBitmapInfo.bmiColors(lRow).rgbGreen
        astGifPalette(lRow).rgbRed = stBitmapInfo.bmiColors(lRow).rgbRed
        If Not bFound Then
            lColor = RGB(astGifPalette(lRow).rgbRed, astGifPalette(lRow).rgbGreen, astGifPalette(lRow).rgbBlue)
            If lColor = lTransColor Then
               bTransIndex = lRow: bFound = True
            End If
        End If
    Next
   
    stScreen.background_color_index = 0
    stScreen.Flags = &HF7 '256-color gif with global color map
    stScreen.pixel_aspect_ratio = 0
    
    stImage.Format = &H7 'GlobalNonInterlaced
    stImage.Height = lPicHeight
    stImage.Width = lPicWidth
    
    '
    ' Get rid of the existing file and create a new one
    '
    If PSGEN_FileExists(sFilename) Then Kill sFilename
    iOutFile = FreeFile
    Open sFilename For Binary As iOutFile

    '
    ' Write GIF header and header info
    '
    If bUseTrans = True Then
        Put #iOutFile, , GIF89a
    Else
        Put #iOutFile, , GIF87A
    End If
    Put #iOutFile, , stScreen
    Put #iOutFile, , astGifPalette
    If bUseTrans = True Then
        Put #iOutFile, , CTRLINTRO
        Put #iOutFile, , CTRLLABEL
        Dim cb As CONTROLBLOCK
        cb.Blocksize = 4 'Always 4
        cb.Flags = 9 'Packed = 00001001 (If Bit 0 = 1: Use Transparency)
        cb.Delay = 0
        cb.TransParent_Color = bTransIndex
        cb.Terminator = 0 'Always 0
        Put #iOutFile, , cb
    End If
    Put #iOutFile, , IMAGESEPARATOR
    Put #iOutFile, , stImage
    bData = CODESIZE - 1
    Put #iOutFile, , bData
    abDataBuffer(0) = 0
    iBitPosition = CHAR_BIT

    '
    ' Process pixels data using LZW/GIF compression
    '
    For lRow = 1 To lPicHeight
        Set objColorTable = New Collection
        Call Z_OutputBitsToGif(iOutFile, CLEARCODE, CODESIZE, iBitPosition, abDataBuffer)
        sPrefix = ""
        iCode = buf(0, lRow)
        On Error Resume Next
        For lCol = 1 To lPicWidth - 1
            sByte = Right$("00" & buf(lCol, lRow), 3)
            sPrefix = sPrefix & sByte
            iCode = objColorTable(sPrefix)
            If Err <> 0 Then 'Prefix wasn't in collection - save it and output code
                iCount = objColorTable.Count
                If iCount = MAX_CODE Then
                    Set objColorTable = New Collection
                    Call Z_OutputBitsToGif(iOutFile, CLEARCODE, CODESIZE, iBitPosition, abDataBuffer)
                End If
                objColorTable.Add iCount + FIRSTCODE, sPrefix
                Call Z_OutputBitsToGif(iOutFile, iCode, CODESIZE, iBitPosition, abDataBuffer)
                sPrefix = sByte
                iCode = buf(lCol, lRow)
                Err.Clear
            End If
        Next
        Call Z_OutputBitsToGif(iOutFile, iCode, CODESIZE, iBitPosition, abDataBuffer)
    Next
    
    
    Call Z_OutputCode(iOutFile, ENDCODE, iCodeCount, iBitPosition, objColorTable, abDataBuffer)
    For lRow = 0 To abDataBuffer(0)
        Put #iOutFile, , abDataBuffer(lRow)
    Next
    bData = 0
    Put #iOutFile, , bData
    Put #iOutFile, , GIFTERMINATOR
    Close iOutFile
   
    '
    ' Clear up all the temporary objects
    '
    Erase buf
    If hDC = 0 Then
        SelectObject lSrchDc, lOldObj
        DeleteDC lSrchDc
    End If
    SelectObject lDC256, lOldObj256
    DeleteObject lDib256
    DeleteDC lDC256
   
    '
    ' Return status to caller
    '
    PSGEN_SaveGIF = True

End Function

Private Sub Z_OutputBitsToGif(ByVal iOutFile%, ByVal iValue%, ByVal iCount%, ByRef iBitPosition%, abDataBuffer() As Byte)
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2004
'
'****************************************************************************
'
'                     NAME: Sub Z_OutputBitsToGif
'
'                     iOutFile%          - Open file handle
'                     iValue%            - Pixel value to output
'                     iCount%            - Number of pixels to output
'                     iBitPosition%      - Position to place value
'                     abDataBuffer       - Buffer to fill
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    19 March 2004   First created for MediaServices
'
'                  PURPOSE: Outputs values to the GIF
'
'****************************************************************************
'
'

Dim iCnt%, iBit%
   
    Do While iCnt < iCount
        If iBitPosition = CHAR_BIT Then
            If abDataBuffer(0) = 255 Then
                Put #iOutFile, , abDataBuffer
                abDataBuffer(0) = 1
            Else
                abDataBuffer(0) = abDataBuffer(0) + 1
            End If
            abDataBuffer(abDataBuffer(0)) = 0
            iBitPosition = 0
        End If
        iBit = Sgn(Z_Power2(iCnt) And iValue)
        If iBit > 0 Then abDataBuffer(abDataBuffer(0)) = Z_Power2(iBitPosition) Or abDataBuffer(abDataBuffer(0))
        iCnt = iCnt + 1: iBitPosition = iBitPosition + 1
    Loop

End Sub

Private Sub Z_OutputCode(ByVal iOutFile%, ByVal iCode%, ByRef iCodeCount%, ByRef iBitPosition%, ByRef objColorTable As Collection, ByRef abDataBuffer() As Byte)
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2004
'
'****************************************************************************
'
'                     NAME: Sub Z_OutputCode
'
'                     iOutFile%          - Open file handle
'                     iCode%             - Code to output
'                     iBitPosition%      - Position to place value
'                     objcolorTable      - Collection of colors in the palette
'                     abDataBuffer       - Buffer to fill
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    19 March 2004   First created for MediaServices
'
'                  PURPOSE: Outputs color code to the GIF
'
'****************************************************************************
'
'
      
   '
   ' Output the file code
   '
   On Error Resume Next
   iCodeCount = iCodeCount + 1
   If iCodeCount > LASTCODE Then
      iCodeCount = FIRSTCODE
      Call Z_OutputBitsToGif(iOutFile, CLEARCODE, CODESIZE, iBitPosition, abDataBuffer)
      Set objColorTable = New Collection
    End If
    Call Z_OutputBitsToGif(iOutFile, iCode, CODESIZE, iBitPosition, abDataBuffer)

End Sub


Private Function Z_Power2&(ByVal iPower%)
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2004
'
'****************************************************************************
'
'                     NAME: Function Z_Power2
'
'                     iPower%          - 2 to the power of
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    19 March 2004   First created for MediaServices
'
'                  PURPOSE: Returns 2 to the power of
'
'****************************************************************************
'
'

Dim aPower2(31) As Long
 
    aPower2(0) = &H1&
    aPower2(1) = &H2&
    aPower2(2) = &H4&
    aPower2(3) = &H8&
    aPower2(4) = &H10&
    aPower2(5) = &H20&
    aPower2(6) = &H40&
    aPower2(7) = &H80&
    aPower2(8) = &H100&
    aPower2(9) = &H200&
    aPower2(10) = &H400&
    aPower2(11) = &H800&
    aPower2(12) = &H1000&
    aPower2(13) = &H2000&
    aPower2(14) = &H4000&
    aPower2(15) = &H8000&
    aPower2(16) = &H10000
    aPower2(17) = &H20000
    aPower2(18) = &H40000
    aPower2(19) = &H80000
    aPower2(20) = &H100000
    aPower2(21) = &H200000
    aPower2(22) = &H400000
    aPower2(23) = &H800000
    aPower2(24) = &H1000000
    aPower2(25) = &H2000000
    aPower2(26) = &H4000000
    aPower2(27) = &H8000000
    aPower2(28) = &H10000000
    aPower2(29) = &H20000000
    aPower2(30) = &H40000000
    aPower2(31) = &H80000000
    
    Z_Power2 = aPower2(iPower)

End Function


Private Function Z_CreateDib256&(ByVal hDC&, stBitmapInfo As BITMAPINFO256, ByVal lPicWidth&, ByVal lPicHeight&)
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2004
'
'****************************************************************************
'
'                     NAME: Sub Z_CreateDib256
'
'                     hDC%               - Context to use for colors
'                     stBitmapInfo       - Bitmap information
'                     lPicWidth&         - Width of output GIF
'                     lPicHeight&         - Height of output GIF
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    19 March 2004   First created for MediaServices
'
'                  PURPOSE: Returns a 256 color DIB
'
'****************************************************************************
'
'

Dim lScanSize&, lPtr&, lIndex&
Dim r&, g&, b&
Dim rA&, gA&, bA&
   
    With stBitmapInfo.bmiHeader
        .biSize = Len(stBitmapInfo.bmiHeader)
        .biWidth = lPicWidth
        .biHeight = lPicHeight
        .biPlanes = 1
        .biBitCount = 8
        .biCompression = BI_RGB
        lScanSize = (lPicWidth + lPicWidth Mod 4)
        .biSizeImage = lScanSize * lPicHeight
    End With
    
    '
    ' Halftone 256 colour palette
    '
    For b = 0 To &H100 Step &H40
       If b = &H100 Then
          bA = b - 1
       Else
          bA = b
       End If
       For g = 0 To &H100 Step &H40
          If g = &H100 Then
             gA = g - 1
          Else
             gA = g
          End If
          For r = 0 To &H100 Step &H40
             If r = &H100 Then
                rA = r - 1
             Else
                rA = r
             End If
             With stBitmapInfo.bmiColors(lIndex)
                .rgbRed = rA: .rgbGreen = gA: .rgbBlue = bA
             End With
             lIndex = lIndex + 1
          Next r
       Next g
    Next b
    
    Z_CreateDib256 = CreateDIBSection256(hDC, stBitmapInfo, DIB_RGB_COLORS, lPtr, 0, 0)

End Function



Public Function PSGEN_NormaliseString$(ByVal sValue$)
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2004
'
'****************************************************************************
'
'                     NAME: Function PSGEN_NormaliseString
'
'                     sValue$           to Value to normalise
'
'                     return  to string with all foreign characters removed
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    19 March 2004   First created for MediaServices
'
'                  PURPOSE: Changes all the accented characters with similar ASCII versions
'
'****************************************************************************
'
'

Dim sReturn$
Dim lCnt&

    '
    ' Loop through all the characters
    '
    sReturn = sValue
    For lCnt = 1 To Len(sValue)
        Select Case Asc(Mid(sValue, lCnt, 1))
        
            Case 0 To 127
            
            Case 192 To 198
                Mid(sReturn, lCnt, 1) = "A"
            
            Case 224 To 230
                Mid(sReturn, lCnt, 1) = "a"
                
            Case 208
                Mid(sReturn, lCnt, 1) = "D"
                
            Case 209
                Mid(sReturn, lCnt, 1) = "N"
                
            Case 210 To 214, 216
                Mid(sReturn, lCnt, 1) = "O"
                
            Case 217 To 220, 181
                Mid(sReturn, lCnt, 1) = "U"
                
            Case 221
                Mid(sReturn, lCnt, 1) = "Y"
                
            Case 222
                Mid(sReturn, lCnt, 1) = "P"
                
            Case 223
                Mid(sReturn, lCnt, 1) = "B"
                
            Case 199
                Mid(sReturn, lCnt, 1) = "C"
                
            Case 200 To 203
                Mid(sReturn, lCnt, 1) = "E"
                
            Case 204 To 207
                Mid(sReturn, lCnt, 1) = "I"
                
            Case 231
                Mid(sReturn, lCnt, 1) = "c"
                
            Case 232 To 235
                Mid(sReturn, lCnt, 1) = "e"
                
            Case 236 To 239
                Mid(sReturn, lCnt, 1) = "i"
                
            Case 242 To 246, 240, 248
                Mid(sReturn, lCnt, 1) = "o"
                
            Case 249, 252
                Mid(sReturn, lCnt, 1) = "u"
                
            Case 253, 255
                Mid(sReturn, lCnt, 1) = "y"
                
            Case 254
                Mid(sReturn, lCnt, 1) = "p"
                
            Case Else
                Mid(sReturn, lCnt, 1) = " "
                
        End Select
    Next lCnt
    
    PSGEN_NormaliseString = Trim(sReturn)

End Function

Public Sub PSGEN_DrawTransparentRectangle(ByVal lDC&, ByVal lLeft&, ByVal lTop&, ByVal lRight&, ByVal lBottom&, ByVal lBorderColor As OLE_COLOR, ByVal lFillColor As OLE_COLOR, ByVal iTransparency%)
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2006
'
'****************************************************************************
'
'                     NAME: Sub PSGEN_DrawTransparentRectangle
'
'                     lDC&                      - Device context to draw on
'                     lLeft&                    - Left coordinate
'                     lTop&                     - Top coordinate
'                     lRight&                   - Right coordinate
'                     lBottom&                  - Bottom coordinate
'                     lBorderColor As OLE_COLOR - Color to use as highlight
'                     lFillColor As OLE_COLOR   - Color to use as highlight
'                     iTransparency%            - Level of transparency to se (0-255)
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    09 May 2006   First created for Project1
'
'                  PURPOSE: Draws a transparent rectangle at the pixel
'                           coordinates using the border color, fill color and level of
'                           transparency provided
'
'****************************************************************************
'
'
Dim lSrcDC&, lSrcBmp&, lBrush&, lBlend&
Dim stRect As RECT
Dim stBlend As BLEND_PROPS


    '
    ' Create a DC to read from
    '
    On Error Resume Next
    lSrcDC = CreateCompatibleDC(lDC)
    lSrcBmp = CreateCompatibleBitmap(lDC, lRight - lLeft, lBottom - lTop)
    Call SelectObject(lSrcDC, lSrcBmp)
    
    '
    ' Fill the source with the colour
    '
    stRect.Left = 0
    stRect.Top = 0
    stRect.Right = lRight - lLeft
    stRect.Bottom = lBottom - lTop
    lBrush = CreateSolidBrush(lFillColor)
    Call FillRect(lSrcDC, stRect, lBrush)
    Call DeleteObject(lBrush)

    '
    ' Now set the blending
    '
    stBlend.tBlendAmount = iTransparency
    Call CopyMemory(lBlend, stBlend, 4)

    '
    ' Blend the two DCs
    '
    Call AlphaBlend(lDC, lLeft, lTop, lRight - lLeft, lBottom - lTop, lSrcDC, 0, 0, lRight - lLeft, lBottom - lTop, lBlend)
    stRect.Left = lLeft
    stRect.Top = lTop
    stRect.Right = lRight
    stRect.Bottom = lBottom
    lBrush = CreateSolidBrush(lBorderColor)
    Call FrameRect(lDC, stRect, lBrush)
    Call DeleteObject(lBrush)
    '
    ' Clean up
    '
    Call DeleteDC(lSrcDC)
    Call DeleteObject(lSrcBmp)

End Sub

Public Function PSGEN_ProcessExists%(ByVal sNameLike$)

Dim lProcessSnapshot&
Dim uProcess As PROCESSENTRY32
Dim iExists%
Dim sName$

    If LenB(sNameLike) <> 0 Then
        lProcessSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
        uProcess.dwSize = LenB(uProcess)

        If lProcessSnapshot > 0 Then
            If Process32First(lProcessSnapshot, uProcess) <> 0 Then
                Do
                    sName = Split(uProcess.szExeFile, vbNullChar, 2)(0)
                    If LCase(sName) Like LCase(sNameLike) Then iExists = iExists + 1
                Loop Until Process32Next(lProcessSnapshot, uProcess) = 0
            End If
            
            'Close snapshot handle
            CloseHandle lProcessSnapshot
        End If
        
    End If
    
    'Return information
    PSGEN_ProcessExists = iExists

End Function


Public Function PSGEN_KillExe(myName As String) As Boolean
  
  Const TH32CS_SNAPPROCESS As Long = 2&
  Const PROCESS_ALL_ACCESS = 0
  Dim uProcess As PROCESSENTRY32
  Dim rProcessFound As Long
  Dim hSnapshot As Long
  Dim szExename As String
  Dim exitCode As Long
  Dim myProcess As Long
  Dim AppKill As Boolean
  Dim appCount As Integer
  Dim i As Integer
  On Local Error GoTo Finish
  appCount = 0
  
  uProcess.dwSize = Len(uProcess)
  hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
  rProcessFound = Process32First(hSnapshot, uProcess)
  Do While rProcessFound
      i = InStr(1, uProcess.szExeFile, Chr(0))
      szExename = LCase$(Left$(uProcess.szExeFile, i - 1))
      If Right$(szExename, Len(myName)) = LCase$(myName) Then
          PSGEN_KillExe = True
          appCount = appCount + 1
          myProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
           If PSGEN_KillProcess(uProcess.th32ProcessID, 0) Then
               'For debug.... Remove this
               'MsgBox "Instance no. " & appCount & " of " & szExename & " was terminated!"
           End If

      End If
      rProcessFound = Process32Next(hSnapshot, uProcess)
  Loop
  Call CloseHandle(hSnapshot)
  Exit Function
Finish:
   'MsgBox "Error!"
End Function

'Terminate any application and return an exit code to Windows.
Function PSGEN_KillProcess(ByVal hProcessID As Long, Optional ByVal exitCode As Long) As Boolean
   Dim hToken As Long
   Dim hProcess As Long
   Dim tp As TOKEN_PRIVILEGES
   

   If GetVersion() >= 0 Then

       If OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken) = 0 Then
           GoTo CleanUp
       End If

       If LookupPrivilegeValue("", "SeDebugPrivilege", tp.LuidUDT) = 0 Then
           GoTo CleanUp
       End If

       tp.PrivilegeCount = 1
       tp.Attributes = SE_PRIVILEGE_ENABLED

       If AdjustTokenPrivileges(hToken, False, tp, 0, ByVal 0&, ByVal 0&) = 0 Then
           GoTo CleanUp
       End If
   End If

   hProcess = OpenProcess(PROCESS_ALL_ACCESS, 0, hProcessID)
   If hProcess Then

       PSGEN_KillProcess = (TerminateProcess(hProcess, exitCode) <> 0)
       ' close the process handle
       CloseHandle hProcess
   End If
   
   If GetVersion() >= 0 Then
       ' under NT restore original privileges
       tp.Attributes = 0
       AdjustTokenPrivileges hToken, False, tp, 0, ByVal 0&, ByVal 0&
       
CleanUp:
       If hToken Then CloseHandle hToken
   End If
   
End Function

Function PSGEN_IsCommaLocale() As Boolean

    PSGEN_IsCommaLocale = InStr(Format(0, "0.0"), ",") > 0

End Function

Function PSGEN_GetLocaleValue#(ByVal sValue$)

    On Error Resume Next
    If PSGEN_IsCommaLocale Then sValue = Replace(sValue, ".", ",")
    PSGEN_GetLocaleValue = CDbl(sValue)

End Function

