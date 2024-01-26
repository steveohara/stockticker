Attribute VB_Name = "Support"
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2004
'
'****************************************************************************
'
' LANGUAGE:             Microsoft Visual Basic V6.00
'
' MODULE NAME:          Pivotal_Main
'
' MODULE TYPE:          BASIC Form
'
' FILE NAME:            PSMAIN.FRM
'
' MODIFICATION HISTORY: Steve O'Hara    30 April 2004   First created for ScaffoldTicker
'
' PURPOSE:              Main ticker interface
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
    ' Version number form the build system
    '
    Public Const VERSION_NAME = "Pivotal Stock Ticker (pivotalstockticker.exe)"
    Public Const VERSION_NUMBER = "3.5.5"
    Public Const VERSION_TIMESTAMP = "26-Jan-2024 13:58"

    '
    ' Registry entries
    '
    Dim mobjReg As New cRegistry

    Public Const REG_SETTINGS = "Settings"
    Public Const REG_LAST_GOOD_RATES = "Last Good Rates"
    Public Const REG_LAST_GOOD_VALUES = "Last Good Values"
    Public Const REG_PROXY = "Proxy"
    Public Const REG_FREQUENCY = "Frequency"
    Public Const REG_FREQUENCY_DEF = 30
    Public Const REG_DOCK_TYPE = "Dock Type"
    Public Const REG_DOCK_AUTOHIDE = "Dock Auto Hide"
    Public Const REG_UPGRADE_SERVER = "Upgrade Server"
    Public Const REG_UPGRADE_SERVER_DEF = "stockticker.pivotal-solutions.co.uk"
    Public Const REG_UPGRADE_DIR = "Upgrade Folder"
    Public Const REG_UPGRADE_DIR_DEF = "/upgrades"
    
    Public Const REG_URL = "URL"
    Public Const REG_LAUNCH_URL = "Launch URL"
    Public Const REG_CHART_URL = "Chart Url"
    Public Const REG_CHART_URL_ALT = "Chart Url Alternative"
    
    Public Const REG_URL_DEF = "http://download.finance.yahoo.com"
    Public Const REG_LAUNCH_URL_DEF = "https://finance.yahoo.com/quote"
    Public Const REG_CHART_URL_DEF = "http://chart.finance.yahoo.com"
    Public Const REG_CHART_URL_ALT_DEF = "http://bigcharts.marketwatch.com/charts/big.chart"
    
    Public Const REG_SYMBOLS = "Symbols"
    Public Const REG_SYMBOL = "Symbol"
    Public Const REG_ALIAS = "Alias"
    Public Const REG_PRICE = "Price"
    Public Const REG_CURRENCY = "Currency"
    Public Const REG_CURRENCY_SYMBOL = "Currency Symbol"
    Public Const REG_SHARES = "Shares"
    Public Const REG_SHOW_PRICE = "Show Price"
    Public Const REG_SHOW_CHANGE = "Show Change"
    Public Const REG_SHOW_CHANGE_PERCENT = "Show Change Percent"
    Public Const REG_SHOW_CHANGE_INDICATOR = "Show Change Indicator"
    Public Const REG_SHOW_DAY_CHANGE = "Show Day Change"
    Public Const REG_SHOW_DAY_CHANGE_PERCENT = "Show Day Change Percent"
    Public Const REG_SHOW_DAY_CHANGE_INDICATOR = "Show Day Change Indicator"
    Public Const REG_SHOW_PROFIT_LOSS = "Show Profit and Loss"
    Public Const REG_EXCLUDE_FROM_SUMMARY = "Exclude From Summary"
    Public Const REG_DISABLED = "Disabled"
    
    Public Const REG_LOW_ALARM_ENABLED = "Low Alarm Enabled"
    Public Const REG_LOW_ALARM_VALUE = "Low Alarm Value"
    Public Const REG_LOW_ALARM_AS_PERCENT = "Low Alarm As Percent"
    Public Const REG_LOW_ALARM_SOUND = "Low Alarm Sound Enabled"
    
    Public Const REG_HIGH_ALARM_ENABLED = "High Alarm Enabled"
    Public Const REG_HIGH_ALARM_VALUE = "High Alarm Value"
    Public Const REG_HIGH_ALARM_AS_PERCENT = "High Alarm As Percent"
    Public Const REG_HIGH_ALARM_SOUND = "High Alarm Sound Enabled"

    Public Const REG_SHOW_SUMMARY_PROFIT_LOSS = "Show Total Profit and Loss"
    Public Const REG_SHOW_SUMMARY_PROFIT_LOSS_PERCENT = "Show Total Profit and Loss as Percent"
    Public Const REG_SHOW_SUMMARY_TOTAL_COST = "Show Total Cost"
    Public Const REG_SHOW_SUMMARY_TOTAL_VALUE = "Show Total Value"
    Public Const REG_SHOW_SUMMARY_DAILY_CHANGE = "Show Daily Change"
    
    Public Const REG_SHOW_SUMMARY_COST_BASE = "Show Cost Base"
    Public Const REG_SHOW_SUMMARY_PRICE = "Show Price"
    Public Const REG_SHOW_SUMMARY_PERCENT = "Show Percent"
    Public Const REG_SHOW_SUMMARY_SUMMARISE = "Summarise"
    Public Const REG_SUMMARY_CURRENCY = "Currency"
    Public Const REG_SUMMARY_CURRENCY_SYMBOL = "Currency Symbol"
    Public Const REG_SUMMARY_TOTAL = "Total Investment"
    Public Const REG_SUMMARY_MARGIN = "Margin"

    Public Const REG_BACK_COLOUR = "Background Colour"
    Public Const REG_TEXT_COLOUR = "Text Colour"
    Public Const REG_UP_COLOUR = "Up Colour"
    Public Const REG_DOWN_COLOUR = "Down Colour"
    Public Const REG_UP_ARROW_COLOUR = "Up Arrow Colour"
    Public Const REG_DOWN_ARROW_COLOUR = "Down Arrow Colour"
    Public Const REG_FONT = "Font"
    Public Const REG_BOLD = "Bold"
    Public Const REG_ALWAYS_ON_TOP = "AlwaysOnTop"
    Public Const REG_IEX_KEY = "IEX key"
    Public Const REG_ALPHA_VANTAGE_KEY = "AlphaVantage key"
    Public Const REG_MARKET_STACK_KEY = "MarketStack key"
    Public Const REG_TWELVE_DATA_KEY = "TwelveData key"
    Public Const REG_FREE_CURRENCY_KEY = "FreeCurrency key"
    Public Const REG_ITALIC = "Italic"
    
    Public Const REG_DAY_SUMMARY_SORT_COLUMN = "Day Summary Sort Column"
    Public Const REG_DAY_SUMMARY_SORT_ORDER = "Day Summary Sort Order"

    Public Const REG_SUMMARY_SORT_COLUMN = "Summary Sort Column"
    Public Const REG_SUMMARY_SORT_ORDER = "Summary Sort Order"

    Public Const REG_HIGH_ALARM_WAVE = "High Alarm Wave File"
    Public Const REG_LOW_ALARM_WAVE = "Low Alarm Wave File"

    Public glPrevWndProc As Long


Public Sub CentreForm(ByVal frmCentre As Form)
Attribute CentreForm.VB_Description = "Centres the form on the appropriate monitor"
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2008
'
'****************************************************************************
'
'                     NAME: Sub CentreForm
'
'                     frmCentre As Form         - Form to centre
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    05 September 2008   First created for StockTicker
'
'                  PURPOSE: Centres the form on the appropriate monitor
'
'****************************************************************************
'
'
Const SM_CXVIRTUALSCREEN = 78

Dim lScreenWidth&, lLeft&, lMonitorWidth&
Dim stRect As RECT


    '
    ' Centre the scrrens on the appropriate monitor
    '
    On Error Resume Next
    Call GetWindowRect(frmMain.hWnd, stRect)
    lScreenWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN)
    lMonitorWidth = Screen.Width / Screen.TwipsPerPixelX
    If lScreenWidth > lMonitorWidth And stRect.Left > lMonitorWidth Then
        Do
            lLeft = lLeft + lMonitorWidth
        Loop Until lLeft + lMonitorWidth > stRect.Left
        frmCentre.Move lLeft * Screen.TwipsPerPixelX + ((Screen.Width - frmCentre.Width) / 2), (Screen.Height - frmCentre.Height) / 3
    Else
        frmCentre.Move (Screen.Width - frmCentre.Width) / 2, (Screen.Height - frmCentre.Height) / 3
    End If

End Sub



Public Function ReadSymbolsFromRegistry()

Dim asSymbols As Variant
Dim sSymbolKey As Variant
Dim objSymbol As cSymbol
Dim objReturn As Collection
Dim iCnt%

    '
    ' Get the symbol keys from the registry
    '
    On Error Resume Next
    Set objReturn = New Collection
    asSymbols = mobjReg.GetSubkeys(App.Title, REG_SYMBOLS)
    If Not IsEmpty(asSymbols) Then
        For iCnt = 0 To UBound(asSymbols)
            asSymbols(iCnt) = Trim(UCase(mobjReg.GetSetting(App.Title, REG_SYMBOLS + "\" + asSymbols(iCnt), REG_SYMBOL))) + "|" + asSymbols(iCnt)
        Next
        asSymbols = PSGEN_SortArraySimple(asSymbols)
        
        '
        ' Loop round all the symbols
        '
        For Each sSymbolKey In asSymbols
            Set objSymbol = New cSymbol
            objSymbol.Init Split(sSymbolKey, "|")(1)
            objReturn.Add objSymbol, objSymbol.RegKey
        Next
    End If
    
    Set ReadSymbolsFromRegistry = objReturn
    
End Function


Public Sub WriteSymbolsToRegistry(objSymbols As Collection)

Dim objSymbol As cSymbol

    '
    ' Clear the registry
    '
    On Error Resume Next
    mobjReg.DeleteSetting App.Title, REG_SYMBOLS, vbNullString

    '
    ' Loop round all the symbols
    '
    For Each objSymbol In objSymbols
        objSymbol.Save
    Next
    
End Sub


Public Function FormatCurrencyValue$(ByVal sSymbol$, ByVal rValue#)

    If InStr(1, "abcdefghijklmnopqrstuvwxyz", sSymbol, vbTextCompare) > 0 Then
        FormatCurrencyValue = Format(rValue, "#,0.00") + sSymbol
    Else
        FormatCurrencyValue = Format(rValue, sSymbol + "#,0.00")
    End If

End Function

Public Function FormatCurrencyValueWithSymbol$(ByVal sSymbol$, ByVal sCurrency$, ByVal rValue#)

    If InStr(1, "abcdefghijklmnopqrstuvwxyz", sSymbol, vbTextCompare) > 0 Then
        If PSGEN_IsSameText(sCurrency, "gbp") Or _
           PSGEN_IsSameText(sCurrency, "gip") Or _
           PSGEN_IsSameText(sCurrency, "fkp") Or _
           PSGEN_IsSameText(sCurrency, "egp") Then
            FormatCurrencyValueWithSymbol = Format(rValue / 100, "£#,0.00")
        
        ElseIf PSGEN_IsSameText(sCurrency, "usd") Or _
               PSGEN_IsSameText(sCurrency, "bnd") Or _
               PSGEN_IsSameText(sCurrency, "hkd") Or _
               PSGEN_IsSameText(sCurrency, "xcd") Or _
               PSGEN_IsSameText(sCurrency, "jmd") Or _
               PSGEN_IsSameText(sCurrency, "lbd") Or _
               PSGEN_IsSameText(sCurrency, "nad") Or _
               PSGEN_IsSameText(sCurrency, "nzd") Or _
               PSGEN_IsSameText(sCurrency, "spd") Or _
               PSGEN_IsSameText(sCurrency, "sgd") Or _
               PSGEN_IsSameText(sCurrency, "twd") Or _
               PSGEN_IsSameText(sCurrency, "ttd") Or _
               PSGEN_IsSameText(sCurrency, "zwd") Or _
               PSGEN_IsSameText(sCurrency, "cad") Then
            FormatCurrencyValueWithSymbol = Format(rValue / 100, "$#,0.00")
        
        ElseIf PSGEN_IsSameText(sCurrency, "eur") Then
            FormatCurrencyValueWithSymbol = Format(rValue / 100, "€#,0.00")
        
        Else
            FormatCurrencyValueWithSymbol = Format(rValue, "#,0.00") + sSymbol
        End If
    Else
        FormatCurrencyValueWithSymbol = Format(rValue, sSymbol + "#,0.00")
    End If

End Function



Public Function PSMAIN_CheckForUpgrade(Optional ByVal bForceDownload As Boolean) As Boolean
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2002
'
'****************************************************************************
'
'                     NAME: Function PSMAIN_CheckForUpgrade
'
'                          bForceDownload as boolean   - True if the products are upgraded without asking the user
'
'                          ) As Boolean
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    25 June 2002   First created for Constructor
'
'                  PURPOSE: Checks the web site to see if there is an upgrade
'                           available and optionally upgrades the system if there is
'
'****************************************************************************
'
'
Dim bReturn As Boolean
Dim iMouse%
Dim sServer$, sDir$, sList$, sDownload$, sFilename$
Dim lLatestMajor&, lLatestMinor&, lLatestRev&
Dim lMajor&, lMinor&, lRev&
Dim asList$()
Dim asTmp$()
Dim iCnt%

    '
    ' Get the server details
    '
    On Error Resume Next
    iMouse = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    sServer = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_UPGRADE_SERVER, REG_UPGRADE_SERVER_DEF)
    sDir = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_UPGRADE_DIR, REG_UPGRADE_DIR_DEF)
    sDir = IIf(InStr(sDir, "/") = 1, "", "/") + Replace(sDir, " ", "%20")
    
    '
    ' Check that the server is reachable
    '
    If Not PSINET_Ping(sServer, 1000) Then
        MsgBox "Cannot reach the upgrade server at this time", vbCritical + vbOKOnly
    Else
    
        '
        ' OK, the server is accessible so let's get the list of available upgrades
        '
        sServer = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_UPGRADE_SERVER, REG_UPGRADE_SERVER_DEF)
        If Not PSINET_GetHTTPFile("http://" + sServer + sDir, sList) Then
            MsgBox "Problem reading the list of upgrades - " + Err.Description, vbCritical + vbOKOnly
        Else
        
            '
            ' Determine the list of files from the names on the server
            ' They will follow this pattern; name_xxxx_yyyy_zzz where xxxx is the major version,
            ' yyyy is the minor and zzzz is the revision
            '
            asList = Split(sList, "<A HREF=""" + sDir, compare:=vbTextCompare)
            
            '
            ' Now loop through all the possible updates to see if there is one greater than we already are
            '
            If UBound(asList) > 0 Then
                lLatestMajor = App.Major
                lLatestMinor = App.Minor
                lLatestRev = App.Revision
                For iCnt = 1 To UBound(asList)
                    asTmp = Split(Split(asList(iCnt), ".exe")(0), "_")
                    If UBound(asTmp) > 2 Then
                        lMajor = CLng(asTmp(1))
                        lMinor = CLng(asTmp(2))
                        lRev = CLng(asTmp(3))
                        
                        '
                        ' Check is this is the latest version
                        '
                        If (lMajor > lLatestMajor) Or _
                           (lMajor = lLatestMajor And lMinor > lLatestMinor) Or _
                           (lMajor = lLatestMajor And lMinor = lLatestMinor And lRev > lLatestRev) Then
                           lLatestMajor = lMajor
                           lLatestMinor = lMinor
                           lLatestRev = lRev
                           sDownload = asList(iCnt)
                        End If
                    End If
                Next
                
                '
                ' Have we got a later version to download
                '
                If sDownload = "" Then
                    MsgBox "The current version you are running (v" + Format(lLatestMajor) + "." + Format(lLatestMinor) + "." + Format(lRev) + ") is the latest version available" + vbCrLf + vbCrLf + "No need to upgrade", vbInformation + vbOKOnly
                Else
                    If MsgBox("A later version of the application is available for download (v" + Format(lLatestMajor) + "." + Format(lLatestMinor) + "." + Format(lRev) + ")" + vbCrLf + vbCrLf + "Download and install it ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                    
                        '
                        ' Download the file
                        '
                        sDownload = "http://" + sServer + sDir + "/" + Split(Split(sDownload, "<")(0), ">")(1)
                        sFilename = App.path + "\" + App.EXEName + "_download.exe"
                        If Not Z_GetHTTPFileToFile(sDownload, sFilename) Then
                            MsgBox "Problem downloading the upgrade - " + Err.Description, vbCritical + vbOKOnly
                        Else
                        
                            '
                            ' Now swap the executable
                            '
                            MsgBox App.Title + " now needs to restart to use the new version", vbInformation + vbOKOnly
                            Shell sFilename + " " + Command, vbHide
                            End
                            
                        End If
                        
                    End If
                End If
            Else
                MsgBox "No updates available", vbInformation + vbOKOnly
            End If
        End If
    End If

    '
    ' Return value to caller
    '
    Screen.MousePointer = iMouse
    PSMAIN_CheckForUpgrade = bReturn

End Function

Private Function Z_GetHTTPFileToFile(ByVal sURL$, ByVal sFilename$, Optional ByVal lFlags& = (INTERNET_FLAG_RELOAD + INTERNET_FLAG_PRAGMA_NOCACHE), Optional vCookies, Optional vHeaders) As Boolean
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
Const BUF_LENGTH = 32000
Dim bReturn As Boolean
Dim abData() As Byte
Dim sBuffer As String * BUF_LENGTH
Dim lFile&, lLength&, lSession&, lTmp&, lFileLength&
Dim bFinished As Boolean
Dim iCnt%, iFile%

    '
    ' Set the cookies
    '
    On Error Resume Next
    Err.Clear
    If Not IsMissing(vCookies) Then
        For iCnt = 0 To UBound(vCookies)
            If Err <> 0 Then Exit For
            Call InternetSetCookie(sURL, PSGEN_GetItem(1, "=", vCookies(iCnt)), PSGEN_GetItem(2, "=", vCookies(iCnt)))
        Next iCnt
    End If
    
    '
    ' Connect to the session
    '
    lSession = InternetOpen(App.Title, INTERNET_OPEN_TYPE_DIRECT, "", "", 0)
    
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
        Err.Raise vbObjectError + ERROR_OFFSET, ERROR_SOURCE, PSINET_TranslateErrorCode(Err.LastDllError)
    Else
        
        '
        ' Check that we got the file we wanted
        '
        lLength = Len(sBuffer)
        Call HttpQueryInfo(lFile, HTTP_QUERY_STATUS_CODE, ByVal sBuffer, lLength, lTmp)
        If Left(sBuffer, 1) <> "2" Then
            Err.Raise vbObjectError + ERROR_OFFSET, ERROR_SOURCE, "STATUS:" + PSINET_GetHttpCodeMessage(Val(Left(sBuffer, lLength)))
        Else
            '
            ' Get the file length
            '
            lLength = Len(sBuffer)
            Call HttpQueryInfo(lFile, HTTP_QUERY_CONTENT_LENGTH, ByVal sBuffer, lLength, lTmp)
            lFileLength = CLng(Left(sBuffer, lLength))
            
            '
            ' Download the file in chunks
            '
            iFile = FreeFile
            If PSGEN_FileExists(sFilename) Then Kill sFilename
            Open sFilename For Binary Access Write As #iFile
            While Not bFinished
                lLength = BUF_LENGTH
                ReDim abData(lLength - 1) As Byte
                Call InternetReadBinaryFile(lFile, abData(0), lLength, lLength)
                If lLength > 0 Then
                    If lLength < BUF_LENGTH Then ReDim Preserve abData(lLength - 1)
                    Put #iFile, , abData
                Else
                    bFinished = True
                End If
                lTmp = lTmp + lLength
            Wend
            Close #iFile
            bReturn = True
        End If
        Call InternetCloseHandle(lFile)
    End If
    Call InternetCloseHandle(lSession)
    DoEvents

    '
    ' Return value to caller
    '
    Z_GetHTTPFileToFile = bReturn

End Function


Public Sub TimerProc(ByVal hWnd As Long, ByVal lngMsg As Long, ByVal lngID As Long, ByVal lngTime As Long)
      
    frmMain.TimerEvent
      
End Sub
    
    
Public Function SubClassedList(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim tItem As DRAWITEMSTRUCT
    Dim sBuff As String * 255
    Dim sItem As String
    Dim lBack As Long
    Dim sTmp As String
    
    On Error Resume Next
    If Msg = WM_DRAWITEM Then
        'Redraw the listbox
        'This function only passes the Address of the DrawItem Structure, so we need to
        'use the CopyMemory API to Get a Copy into the Variable we setup:
        Call CopyMemory(tItem, ByVal lParam, Len(tItem))
        'Make sure we're dealing with a Listbox
        If tItem.CtlType = ODT_LISTBOX Then
            
            'Get the Item Text
            sTmp = tItem.ItemData
            Call SendMessage(tItem.hwndItem, LB_GETTEXT, tItem.itemID, ByVal sBuff)
            sItem = Left(sBuff, InStr(sBuff, Chr(0)) - 1)
            If (tItem.itemState And ODS_FOCUS) Then
                'Item has Focus, Highlight it, I'm using the Default Focus
                'Colors for this example.
                lBack = CreateSolidBrush(GetSysColor(COLOR_HIGHLIGHT))
                Call FillRect(tItem.hDC, tItem.rcItem, lBack)
                Call SetBkColor(tItem.hDC, GetSysColor(COLOR_HIGHLIGHT))
                Call SetTextColor(tItem.hDC, GetSysColor(COLOR_HIGHLIGHTTEXT))
                TextOut tItem.hDC, tItem.rcItem.Left, tItem.rcItem.Top, ByVal sItem, Len(sItem)
                DrawFocusRect tItem.hDC, tItem.rcItem
            
            ElseIf Not frmSymbols.mobjSymbols.Item(sTmp).Disabled Then
            
                'Item Doesn't Have Focus, Draw it's Colored Background
                'Create a Brush using the Color we stored in ItemData
                lBack = CreateSolidBrush(vbRed)
                'Paint the Item Area
                Call FillRect(tItem.hDC, tItem.rcItem, lBack)
                'Set the Text Colors
                Call SetBkColor(tItem.hDC, vbRed)
                Call SetTextColor(tItem.hDC, vbWhite)
                'Display the Item Text
                TextOut tItem.hDC, tItem.rcItem.Left, tItem.rcItem.Top, ByVal sItem, Len(sItem)
            
            Else
            
                'Item Doesn't Have Focus, Draw it's Colored Background
                'Create a Brush using the Color we stored in ItemData
                lBack = CreateSolidBrush(tItem.ItemData)
                'Paint the Item Area
                Call FillRect(tItem.hDC, tItem.rcItem, vbWhite)
                'Set the Text Colors
                Call SetBkColor(tItem.hDC, vbWhite)
                Call SetTextColor(tItem.hDC, IIf(tItem.ItemData = vbBlack, vbWhite, vbBlack))
                'Display the Item Text
                TextOut tItem.hDC, tItem.rcItem.Left, tItem.rcItem.Top, ByVal sItem, Len(sItem)
            End If
            Call DeleteObject(lBack)
            'Don't Need to Pass a Value on as we've just handled the Message ourselves
            SubClassedList = 0
            Exit Function
                    
        End If
            
    End If
    SubClassedList = CallWindowProc(glPrevWndProc, hWnd, Msg, wParam, lParam)
End Function


Public Function ParseCSV(ByRef Expression As String, asValues() As String) As Long
' by Donald, donald@xbeat.net, 20020603, rev 20020701
  
  Const lAscSpace     As Long = 32   ' Asc(" ")
  Const lAscQuote     As Long = 34   ' Asc("""")
  Const lAscSeparator As Long = 44   ' Asc(","), comma
  
  Const lValueNone    As Long = 0 ' states of the parser
  Const lValuePlain   As Long = 1
  Const lValueQuoted  As Long = 2
  
  ' BUFFERREDIM is ideally exactly the number of values in Expression (minus 1)
  ' so: if you know what to expect, fine-tune here
  Const BUFFERREDIM   As Long = 64
  Dim ubValues        As Long
  Dim cntValues       As Long
  
  Dim abExpression() As Byte
  Dim lCharCode As Long
  Dim posStart As Long
  Dim posEnd As Long
  Dim cntTrim As Long
  Dim lState As Long
  Dim i As Long
  
  If LenB(Expression) > 0 Then
    
    abExpression = Expression         ' to byte array
    ubValues = -1 + BUFFERREDIM
    ReDim Preserve asValues(ubValues) ' init array (Preserve is faster)
    
    For i = 0 To UBound(abExpression) Step 2
      
      ' 1. unicode char has 16 bits, but 32 bit Longs process faster
      ' 2. add lower and upper byte: ignoring the upper byte can lead to misinterpretations
      lCharCode = abExpression(i) Or (&H100 * abExpression(i + 1))
      
      Select Case lCharCode
      
      Case lAscSpace
        If lState = lValuePlain Then
          ' at non-quoted value: trim 2 unicode bytes for each space
          cntTrim = cntTrim + 2
        End If
      
      Case lAscSeparator
        If lState = lValueNone Then
          ' ends zero-length value
          If cntValues > ubValues Then
            ubValues = ubValues + BUFFERREDIM
            ReDim Preserve asValues(ubValues)
          End If
          asValues(cntValues) = ""
          cntValues = cntValues + 1
          posStart = i + 2
        ElseIf lState = lValuePlain Then
          ' ends non-quoted value
          lState = lValueNone
          posEnd = i - cntTrim
          If cntValues > ubValues Then
            ubValues = ubValues + BUFFERREDIM
            ReDim Preserve asValues(ubValues)
          End If
          asValues(cntValues) = MidB$(Expression, posStart + 1, posEnd - posStart)
          cntValues = cntValues + 1
          posStart = i + 2
          cntTrim = 0
        End If
      
      Case lAscQuote
        If lState = lValueNone Then
          ' starts quoted value
          lState = lValueQuoted
          ' trims the opening quote
          posStart = i + 2
        ElseIf lState = lValueQuoted Then
          ' ends quoted value, or is a quote within
          lState = lValuePlain
          ' trims the closing quote
          cntTrim = 2
        End If
      
      Case Else
        If lState = lValueNone Then
          ' starts non-quoted value
          lState = lValuePlain
          posStart = i
        End If
        ' reset trimming
        cntTrim = 0
      
      End Select
    
    Next
    
    ' remainder
    posEnd = i - cntTrim
    If cntValues <> ubValues Then
      ReDim Preserve asValues(cntValues)
    End If
    asValues(cntValues) = MidB$(Expression, posStart + 1, posEnd - posStart)
    ParseCSV = cntValues + 1
  
  Else
    ' (Expression = "")
    ' return single-element array containing a zero-length string
    ReDim asValues(0)
    ParseCSV = 1
  
  End If

End Function

Public Function getJsonValue(ByRef sSrc As String, sName As String) As String

Dim tmp$

    tmp = ""
    If InStr(sSrc, sName) > 0 Then
        tmp = Split(Split(sSrc, sName)(1), ":")(1)
        If InStr(tmp, ",") > 0 Then
            tmp = Split(tmp, ",")(0)
        Else
            tmp = Split(tmp, "}")(0)
        End If
        tmp = Trim(LCase(tmp))
        If tmp = "null" Then
            tmp = ""
        End If
    End If
    getJsonValue = tmp

End Function

Private Sub Main()

Dim asArgs$()
Dim bHandled As Boolean
Dim sExeFile$
    
    '
    ' Need to check if we are being run with a command
    '
    asArgs = Split(Command$, " ")
    bHandled = False
    If UBound(asArgs) > -1 Then
        If asArgs(0) = "-version" Or asArgs(0) = "-v" Then
            MsgBox VERSION_NAME + vbCrLf + "Version:" + VERSION_NUMBER + " Built:" + VERSION_TIMESTAMP
            bHandled = True
        End If
    End If
    
    '
    ' Show the main dialog if we are not just running a command line
    '
    If Not bHandled Then
    
        '
        ' Check to see if this is the upgrade
        '
        If App.EXEName Like "*_download" Then
            sExeFile = Split(App.EXEName, "_download")(0) + ".exe"
            If PSGEN_ProcessExists(sExeFile) > 0 Then
                Call PSGEN_KillExe(sExeFile)
                Call Sleep(500)
            End If
            FileCopy App.path + "\" + App.EXEName + ".exe", App.path + "\" + sExeFile
            Shell App.path + "\" + sExeFile + Command, vbHide
            Call Sleep(500)
            End
        Else
            
            '
            ' Check to see if the app is already running
            '
            If PSGEN_ProcessExists(App.EXEName + ".exe") > 1 Then
                End
            Else
                frmMain.Show
            End If
        End If
    End If

End Sub

Sub DrawUpDownArrow(ByVal objSymbol As cSymbol, objSurface As Object, ByVal iScaleHeight%, ByVal iLeft%, ByVal iTop%)

Dim lTextColor&, lUpArrowColor&, lDownArrowColor&
Dim i%

    lTextColor = CLng(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_TEXT_COLOUR, Format(vbWhite)))
    lUpArrowColor = CLng(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_UP_ARROW_COLOUR, Format(vbGreen)))
    lDownArrowColor = CLng(mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_DOWN_ARROW_COLOUR, Format(vbRed)))
    
    If objSymbol.CurrentPrice < objSymbol.DayStart Then
        For i = 0 To iScaleHeight \ 4
           objSurface.Line (iLeft + (iScaleHeight \ 4) - i, iScaleHeight - i - 3)-(iLeft + (iScaleHeight \ 4) + i, iScaleHeight - i - 3), lDownArrowColor
        Next i
        objSurface.DrawWidth = 2
        objSurface.Line (iLeft + (iScaleHeight \ 4), iTop + 3)-(iLeft + (iScaleHeight \ 4), 3 * iScaleHeight \ 4), lDownArrowColor
    
    ElseIf objSymbol.CurrentPrice > objSymbol.DayStart Then
        For i = 0 To iScaleHeight \ 4
           objSurface.Line (iLeft + (iScaleHeight \ 4) - i, i + 3)-(iLeft + (iScaleHeight \ 4) + i, i + 3), lUpArrowColor
        Next i
        objSurface.DrawWidth = 2
        objSurface.Line (iLeft + (iScaleHeight \ 4), iTop + iScaleHeight \ 4)-(iLeft + (iScaleHeight \ 4), iScaleHeight - 4), lUpArrowColor
    Else
        For i = 0 To iScaleHeight \ 4
           objSurface.Line (iLeft + (iScaleHeight \ 4) - i, iScaleHeight - i - 3)-(iLeft + (iScaleHeight \ 4) + i, iScaleHeight - i - 3)
        Next i
        For i = 0 To iScaleHeight \ 4
           objSurface.Line (iLeft + (iScaleHeight \ 4) - i, i + 3)-(iLeft + (iScaleHeight \ 4) + i, i + 3), lTextColor
        Next i
        objSurface.DrawWidth = 2
        objSurface.Line (iLeft + (iScaleHeight \ 4), iTop + iScaleHeight \ 4)-(iLeft + (iScaleHeight \ 4), 3 * iScaleHeight \ 4), lTextColor
    End If

    objSurface.DrawWidth = 1
    objSurface.CurrentY = iTop
    objSurface.CurrentX = objSurface.CurrentX + 4

End Sub

Function ConvertCurrency#(ByVal objSymbol As Object, ByVal rValue#)

Dim sCurrencyDest$, sCSV$, sURL$
Dim rRate#

    On Error Resume Next
    ConvertCurrency = rValue
    rRate = frmMain.mobjExchangeRates.Item(objSymbol.CurrencyName)
    If rRate > 0 Then
        ConvertCurrency = rValue * rRate
        If InStr(1, "abcdefghijklmnopqrstuvwxyz", objSymbol.CurrencySymbol, vbTextCompare) > 0 Then ConvertCurrency = ConvertCurrency / 100
    End If

End Function


