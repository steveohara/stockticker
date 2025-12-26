Attribute VB_Name = "Data"
'
' Copyright (c) 2024, Pivotal Solutions and/or its affiliates. All rights reserved.
' Pivotal Solutions PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
'
' Provides general purpose functions and procedures.
'
Option Explicit

    Dim mobjReg As New cRegistry


Public Function PSDATA_GetExchangeRates(objCurrentSymbols As Collection, objExchangeRates As Collection) As Boolean
'
' objCurrentSymbols- Collection of symbols to get exchange rates for
' objExchangeRates - Collection of current exchange rates
' return Boolean   - True if successful
'
' Reads the exchange rates from IEX using the supplied key and saves
' them in the registry
'
Dim objSymbol As cSymbol
Dim objExchangeLookup As New Collection
Dim sSummaryCurrencyName$

    ' Check we have a base currency
    On Error Resume Next
    PSGEN_Log "Getting exchange rates", LogEventTypes.LogInformation, EventIdTypes.ExchangeRates
    sSummaryCurrencyName = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_SUMMARY_CURRENCY, "GBP")

    ' Make sure we have a sensible rate for the summary
    objExchangeRates.Remove sSummaryCurrencyName
    objExchangeRates.Add 1#, sSummaryCurrencyName
    
    ' Get all the exchange rates for all the symbols that are not disabled
    For Each objSymbol In objCurrentSymbols
        If Not objSymbol.Disabled Then
            If objSymbol.CurrencyName <> "" And Not PSGEN_IsSameText(objSymbol.CurrencyName, sSummaryCurrencyName) Then
                objExchangeLookup.Add objSymbol.CurrencyName, objSymbol.CurrencyName
            End If
        End If
    Next
    
    ' Now get all the exchange rates
    If objExchangeLookup.Count > 0 Then
        Call Z_GetExchangeRatesFromIEX(objExchangeLookup, sSummaryCurrencyName, objExchangeRates)
        Call Z_GetExchangeRatesFromFreeCurrency(objExchangeLookup, sSummaryCurrencyName, objExchangeRates)
        Call Z_GetExchangeRatesFromER(objExchangeLookup, sSummaryCurrencyName, objExchangeRates)
    End If
    
    PSGEN_Log "Got all exchange rates " & IIf(objExchangeLookup.Count = 0, "successfully", "unsuccessfully"), IIf(objExchangeLookup.Count = 0, LogEventTypes.LogInformation, LogEventTypes.LogWarning), EventIdTypes.ExchangeRates
    PSDATA_GetExchangeRates = (objExchangeLookup.Count = 0)

End Function

Private Function Z_GetExchangeRatesFromIEX(objExchangeLookup As Collection, ByVal sSummaryCurrencyName, objExchangeRates As Collection) As Boolean
'
' objExchangeLookup- Collection of currencies to lookup
' sSummaryCurrency - Currency symbol to be used for the summary
' objExchangeRates - Collection of current exchange rates
' return Boolean   - True if successful
'
' Reads the exchange rates from IEX using the supplied key and saves
' them in the registry
'
Dim bag As JsonBag
Dim sCurrency As Variant
Dim sData$, sCurrencies$, sAPIKey, sProxy$
Dim rRate#

    On Error Resume Next
    Z_GetExchangeRatesFromIEX = False

    ' Check to see if there's anything to do
    If objExchangeLookup.Count > 0 Then
    
        ' Check we have a key
        sAPIKey = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_IEX_KEY)
        If sAPIKey <> "" Then
            
            sProxy = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_PROXY)
            sCurrencies = ""
            For Each sCurrency In objExchangeLookup
                DoEvents
                If Not PSGEN_IsSameText(sCurrency, sSummaryCurrencyName) Then
                    sCurrencies = sCurrencies + IIf(sCurrencies = "", "", ",") + sCurrency
                End If
            Next
            
            ' Get all the rates in one go
            PSGEN_Log "Getting exchange rates from https://api.iex.cloud", LogEventTypes.LogInformation, EventIdTypes.ExchangeRates
            Call PSINET_GetHTTPFile("https://api.iex.cloud/v1/fx/convert?symbols=" + sCurrencies + "&token=" + sAPIKey, sData, sProxyName:=sProxy, lConnectionTimeout:=1000, lReadTimeout:=1000, iRetries:=3)
            If Trim(sData) <> "" Then
                
                ' We have something to process
                PSGEN_Log "Got exchange rates from https://api.iex.cloud successfully", LogEventTypes.LogInformation, EventIdTypes.ExchangeRates
                Set bag = New JsonBag
                bag.JSON = sData
                
                ' For each exchange rate
                For Each bag In bag
                    rRate = CDbl(bag.Item("rate"))
                    
                    ' Check that it is real
                    If rRate > 0 Then
                        rRate = 1 / rRate
                        sCurrency = Replace(bag.Item("symbol"), sSummaryCurrencyName, "")
                        Call mobjReg.SaveSetting(App.Title, REG_LAST_GOOD_RATES, sCurrency, rRate)
                        
                        ' Remove it from the lookup list so that we don't get it again
                        objExchangeLookup.Remove sCurrency
                        Z_GetExchangeRatesFromIEX = True
                            
                        ' Update the mast list we have been sent
                        objExchangeRates.Remove sCurrency
                        objExchangeRates.Add rRate, sCurrency
                    End If
                Next
            Else
                PSGEN_Log "Failed to get exchange rates from https://api.iex.cloud - " + Err.Description, LogEventTypes.LogError, EventIdTypes.ExchangeRates
            End If
        End If
    End If

End Function

Private Function Z_GetExchangeRatesFromFreeCurrency(objExchangeLookup As Collection, ByVal sSummaryCurrencyName, objExchangeRates As Collection) As Boolean
'
' objExchangeLookup- Collection of currencies to lookup
' sSummaryCurrency - Currency symbol to be used for the summary
' objExchangeRates - Collection of current exchange rates
' return Boolean   - True if successful
'
' Reads the exchange rates from Free Currency using the supplied key and saves
' them in the registry
'
Dim bag As JsonBag
Dim sCurrency As Variant
Dim sData$, sCurrencies$, sAPIKey, sProxy$
Dim rRate#

    On Error Resume Next
    Z_GetExchangeRatesFromFreeCurrency = False

    ' Check to see if there's anything to do
    If objExchangeLookup.Count > 0 Then
    
        ' Check we have a key
        sAPIKey = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_FREE_CURRENCY_KEY)
        If sAPIKey <> "" Then
            
            ' Get a list of the currencies
            sCurrencies = ""
            For Each sCurrency In objExchangeLookup
                DoEvents
                If Not PSGEN_IsSameText(sCurrency, sSummaryCurrencyName) Then
                    sCurrencies = sCurrencies + IIf(sCurrencies = "", "", ",") + sCurrency
                End If
            Next
            
            PSGEN_Log "Getting exchange rates from https://api.freecurrencyapi.com", LogEventTypes.LogInformation, EventIdTypes.ExchangeRates
            Call PSINET_GetHTTPFile("https://api.freecurrencyapi.com/v1/latest?apikey=" + sAPIKey + "&base_currency=" + sSummaryCurrencyName + "&currencies=" + sCurrencies, sData, sProxyName:=sProxy, lConnectionTimeout:=1000, lReadTimeout:=1000, iRetries:=3)
            If Trim(sData) <> "" Then
                
                PSGEN_Log "Got exchange rates from https://api.freecurrencyapi.com successfully", LogEventTypes.LogInformation, EventIdTypes.ExchangeRates
                Set bag = New JsonBag
                bag.JSON = sData
                
                ' For each currency
                For Each sCurrency In objExchangeLookup
                    If Not PSGEN_IsSameText(sCurrency, sSummaryCurrencyName) Then
                        rRate = CDbl(bag.Item("data").Item(sCurrency))
                        
                        ' Check the rate isn't nonsense
                        If rRate > 0 Then
                            rRate = 1 / rRate
                            Call mobjReg.SaveSetting(App.Title, REG_LAST_GOOD_RATES, sCurrency, rRate)
                            objExchangeLookup.Remove sCurrency
                            Z_GetExchangeRatesFromFreeCurrency = True
                            
                            ' Update the mast list we have been sent
                            objExchangeRates.Remove sCurrency
                            objExchangeRates.Add rRate, sCurrency
                        End If
                    End If
                Next
            Else
                PSGEN_Log "Failed to get exchange rates from https://api.freecurrencyapi.com - " + Err.Description, LogEventTypes.LogError, EventIdTypes.ExchangeRates
            End If
        End If
    End If
End Function

Private Function Z_GetExchangeRatesFromER(objExchangeLookup As Collection, ByVal sSummaryCurrencyName, objExchangeRates As Collection) As Boolean
'
' objExchangeLookup- Collection of currencies to lookup
' sSummaryCurrency - Currency symbol to be used for the summary
' objExchangeRates - Collection of current exchange rates
' return Boolean   - True if successful
'
' Reads the exchange rates from ER using the supplied key and saves
' them in the registry
'
Dim bag As JsonBag
Dim sCurrency As Variant
Dim sData$, sCurrencies$, sAPIKey, sProxy$
Dim rRate#

    On Error Resume Next
    Z_GetExchangeRatesFromER = False

    ' Check to see if there's anything to do
    If objExchangeLookup.Count > 0 Then
    
        ' Check we have a key
        sAPIKey = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_FREE_CURRENCY_KEY)
        If sAPIKey <> "" Then
            
            PSGEN_Log "Getting exchange rates from https://open.er-api.com", LogEventTypes.LogInformation, EventIdTypes.ExchangeRates
            Call PSINET_GetHTTPFile("https://open.er-api.com/v6/latest/" + sSummaryCurrencyName, sData, sProxyName:=sProxy, lConnectionTimeout:=1000, lReadTimeout:=1000, iRetries:=3)
            If Trim(sData) <> "" Then
                
                PSGEN_Log "Got exchange rates from https://open.er-api.com successfully", LogEventTypes.LogInformation, EventIdTypes.ExchangeRates
                Set bag = New JsonBag
                bag.JSON = sData
                
                ' For each currency
                For Each sCurrency In objExchangeLookup
                    If Not PSGEN_IsSameText(sCurrency, sSummaryCurrencyName) Then
                        rRate = Val(Trim(Split(Split(sData, """" + sCurrency + """:")(1), ",")(0)))
                        
                        ' Check the rate isn't rubbish
                        If rRate > 0 Then
                            rRate = 1 / rRate
                            Call mobjReg.SaveSetting(App.Title, REG_LAST_GOOD_RATES, sCurrency, rRate)
                            objExchangeLookup.Remove sCurrency
                            Z_GetExchangeRatesFromER = True
                            
                            ' Update the master list we have been sent
                            objExchangeRates.Remove sCurrency
                            objExchangeRates.Add rRate, sCurrency
                        End If
                    End If
                Next
            Else
                PSGEN_Log "Failed to get exchange rates from https://open.er-api.com - " + Err.Description, LogEventTypes.LogError, EventIdTypes.ExchangeRates
            End If
        End If
    End If
End Function

'
' objSymsToLookup    - Collection of ticket symbols to lookup
' objSymbolsWithData - Collection to poulate with cStock prices
'
' Reads the current stock prioces from the exchange and populates the collection with those values
'
Public Sub PSDATA_GetPrices(objSymsToLookup As Collection, objSymbolsWithData As Collection)

    On Error Resume Next

    ' Now get a list of all the symbols from IEX if we have a key
    Call Z_GetPricesFromIEX(objSymsToLookup, objSymbolsWithData)
    Call Z_GetPricesFromTwelveData(objSymsToLookup, objSymbolsWithData)
    Call Z_GetPricesFromAlphaVantage(objSymsToLookup, objSymbolsWithData)
    Call Z_GetPricesFromMarketStack(objSymsToLookup, objSymbolsWithData)
    Call Z_GetPricesFromFinhub(objSymsToLookup, objSymbolsWithData)
    Call Z_GetPricesFromTiingo(objSymsToLookup, objSymbolsWithData)
    
    ' Dont't require a key
    Call Z_GetPricesFromYahoo(objSymsToLookup, objSymbolsWithData)
    Call Z_GetPricesFromReuters(objSymsToLookup, objSymbolsWithData)

End Sub

'
' objSymsToLookup    - Collection of ticket symbols to lookup
' objSymbolsWithData - Collection to poulate with cStock prices
'
' Reads the current stock prioces from the exchange and populates the collection with those values
'
Private Sub Z_GetPricesFromIEX(objSymsToLookup As Collection, objSymbolsWithData As Collection)

Dim bag As JsonBag
Dim sSymbol As Variant
Dim sData$, sAPIKey, sProxy$, sName$
Dim rCurrentPrice#, rDayLow#, rDayHigh#, rDayOpen#, rDayClose#
Dim objStock As cStock


    ' Check to see if there's anything to do
    On Error Resume Next
    If objSymsToLookup.Count > 0 Then
    
        ' Check we have a key
        sAPIKey = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_IEX_KEY)
        If sAPIKey <> "" Then
            
            sProxy = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_PROXY)
            For Each sSymbol In objSymsToLookup
                rDayOpen = 0
                rDayHigh = 0
                rDayLow = 0
                rCurrentPrice = 0
        
                DoEvents
                sName = sSymbol
                If Not sSymbol Like "*.L" Then
                    PSGEN_Log "Getting stock price from IEX for " + sSymbol, LogEventTypes.LogInformation, EventIdTypes.StockPrices
                    Call PSINET_GetHTTPFile("https://cloud.iexapis.com/stable/stock/" + Replace(sName, "^", ".") + "/quote?token=" + sAPIKey, sData, sProxyName:=sProxy, lConnectionTimeout:=1000, lReadTimeout:=1000, iRetries:=3)
                    
                    ' Put the stock values into the lookup
                    If Trim(sData) <> "" Then
                        PSGEN_Log "Got stock price from IEX successfully for " + sSymbol, LogEventTypes.LogInformation, EventIdTypes.StockPrices
                        Set bag = New JsonBag
                        bag.JSON = sData
                        rDayOpen = CDbl(bag.Item("previousClose"))
                        rDayHigh = CDbl(bag.Item("high"))
                        rDayLow = CDbl(bag.Item("low"))
                        rCurrentPrice = CDbl(bag.Item("iexRealtimePrice"))
                        If rCurrentPrice = 0 Or bag.Item("iexRealtimePrice") = "" Then
                            rCurrentPrice = CDbl(bag.Item("latestPrice"))
                        End If
                        If rCurrentPrice <> 0 Then
                            Set objStock = New cStock
                            objStock.Code = sSymbol
                            objStock.CurrentPrice = rCurrentPrice
                            objStock.DayStart = rDayLow
                            objStock.DayHigh = rDayHigh
                            objStock.DayChange = rCurrentPrice - rDayOpen
                            objStock.Source = "IEX"
                            objSymbolsWithData.Add objStock, sSymbol
                            objSymsToLookup.Remove sSymbol
                        Else
                           PSGEN_Log "Zero value returned from IEX for " + sSymbol, LogEventTypes.LogWarning, EventIdTypes.StockPrices
                        End If
                    Else
                        PSGEN_Log "Failed to stock price from IEX for " + sSymbol + " - " + Err.Description, LogEventTypes.LogInformation, EventIdTypes.StockPrices
                    End If
                End If
            Next
        End If
    End If

End Sub

'
' objSymsToLookup    - Collection of ticket symbols to lookup
' objSymbolsWithData - Collection to poulate with cStock prices
'
' Reads the current stock prioces from the exchange and populates the collection with those values
'
Private Sub Z_GetPricesFromTwelveData(objSymsToLookup As Collection, objSymbolsWithData As Collection)

Dim sSymbol As Variant
Dim asSymVals$()
Dim sData$, sAPIKey, sProxy$
Dim rCurrentPrice#, rDayLow#, rDayHigh#, rDayOpen#, rDayClose#
Dim objStock As cStock


    ' Check to see if there's anything to do
    On Error Resume Next
    If objSymsToLookup.Count > 0 Then
    
        ' Check we have a key
        sAPIKey = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_TWELVE_DATA_KEY)
        If sAPIKey <> "" Then
            
            sProxy = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_PROXY)
            For Each sSymbol In objSymsToLookup
                rDayOpen = 0
                rDayHigh = 0
                rDayLow = 0
                rCurrentPrice = 0
        
                DoEvents
                PSGEN_Log "Getting stock price from Twelve Data for " + sSymbol, LogEventTypes.LogInformation, EventIdTypes.StockPrices
                Call PSINET_GetHTTPFile("https://api.twelvedata.com/quote?format=csv&apikey=" + sAPIKey + "&symbol=" + Replace(sSymbol, ".L", ""), sData, sProxyName:=sProxy, lConnectionTimeout:=1000, lReadTimeout:=1000, iRetries:=2)
    
                ' Put the stock values into the lookup
                If Trim(sData) <> "" Then
                    PSGEN_Log "Got stock price from Twelve Data successfully for " + sSymbol, LogEventTypes.LogInformation, EventIdTypes.StockPrices
                    sData = Replace(Split(sData, vbLf)(1), ";", ",")
                    Call ParseCSV(sData, asSymVals)
                    If UBound(asSymVals) > 8 Then
                        rDayOpen = CDbl(asSymVals(7))
                        rDayHigh = CDbl(asSymVals(8))
                        rDayLow = CDbl(asSymVals(9))
                        rDayClose = CDbl(asSymVals(10))
                        
                        Call PSINET_GetHTTPFile("https://api.twelvedata.com/price?format=csv&apikey=" + sAPIKey + "&symbol=" + Replace(sSymbol, ".L", ""), sData, sProxyName:=sProxy, lConnectionTimeout:=1000, lReadTimeout:=1000, iRetries:=2)
                        If Trim(sData) <> "" Then
                            sData = Split(sData, vbLf)(1)
                            rCurrentPrice = CDbl(sData)
            
                            Set objStock = New cStock
                            objStock.Code = sSymbol
                            objStock.CurrentPrice = rCurrentPrice
                            objStock.DayStart = rDayLow
                            objStock.DayHigh = rDayHigh
                            objStock.DayChange = rCurrentPrice - rDayClose
                            objStock.Source = "TwelveData"
                            objSymbolsWithData.Add objStock, sSymbol
                            objSymsToLookup.Remove sSymbol
                        End If
                    Else
                        PSGEN_Log "Cannot parse data from Twelve Data for " + sSymbol + " " + sData, LogEventTypes.LogWarning, EventIdTypes.StockPrices
                    End If
                Else
                    PSGEN_Log "Failed to get stock prices from Twelve Data for " + sSymbol + " - " + Err.Description, LogEventTypes.LogError, EventIdTypes.StockPrices
                End If
            Next
        End If
    End If

End Sub

'
' objSymsToLookup    - Collection of ticket symbols to lookup
' objSymbolsWithData - Collection to poulate with cStock prices
'
' Reads the current stock prioces from the exchange and populates the collection with those values
'
Private Sub Z_GetPricesFromAlphaVantage(objSymsToLookup As Collection, objSymbolsWithData As Collection)

Dim sSymbol As Variant
Dim asSymVals$()
Dim sCSV$, sAPIKey, sProxy$
Dim rCurrentPrice#, rDayLow#, rDayHigh#, rDayOpen#, rDayClose#
Dim objStock As cStock


    ' Check to see if there's anything to do
    On Error Resume Next
    If objSymsToLookup.Count > 0 Then
    
        ' Check we have a key
        sAPIKey = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_ALPHA_VANTAGE_KEY)
        If sAPIKey <> "" Then
            
            sProxy = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_PROXY)
            For Each sSymbol In objSymsToLookup
                rDayOpen = 0
                rDayHigh = 0
                rDayLow = 0
                rCurrentPrice = 0
        
                DoEvents
                PSGEN_Log "Getting stock price from Alpha Vantage for " + sSymbol, LogEventTypes.LogInformation, EventIdTypes.StockPrices
                Call PSINET_GetHTTPFile("https://www.alphavantage.co/query?function=TIME_SERIES_INTRADAY&interval=1min&apikey=" + sAPIKey + "&datatype=csv&symbol=" + Replace(sSymbol, "^", "."), sCSV, sProxyName:=sProxy, lConnectionTimeout:=1000, lReadTimeout:=1000, iRetries:=2)
    
                ' Put the stock values into the lookup
                If Trim(sCSV) <> "" Then
                    PSGEN_Log "Got stock price from Alpha Vantage successfully for " + sSymbol, LogEventTypes.LogInformation, EventIdTypes.StockPrices
                    sCSV = Split(sCSV, vbLf)(1)
                    Call ParseCSV(sCSV, asSymVals)
                    If UBound(asSymVals) = 5 Then
                        rDayOpen = CDbl(asSymVals(4))
                        rDayHigh = CDbl(asSymVals(2))
                        rDayLow = CDbl(asSymVals(3))
                        rCurrentPrice = CDbl(asSymVals(2))
        
                        Set objStock = New cStock
                        objStock.Code = sSymbol
                        objStock.CurrentPrice = rCurrentPrice
                        objStock.DayStart = rDayLow
                        objStock.DayHigh = rDayHigh
                        objStock.DayChange = rCurrentPrice - rDayOpen
                        objStock.Source = "AlphaVantage"
                        objSymbolsWithData.Add objStock, sSymbol
                        objSymsToLookup.Remove sSymbol
                    Else
                        PSGEN_Log "Cannot parse data from Alpha Vantage for " + sSymbol + " " + sCSV, LogEventTypes.LogWarning, EventIdTypes.StockPrices
                    End If
                Else
                    PSGEN_Log "Failed to get stock price from Alpha Vantage for " + sSymbol + " - " + Err.Description, LogEventTypes.LogInformation, EventIdTypes.StockPrices
                End If
            Next
        End If
    End If

End Sub

'
' objSymsToLookup    - Collection of ticket symbols to lookup
' objSymbolsWithData - Collection to poulate with cStock prices
'
' Reads the current stock prioces from the exchange and populates the collection with those values
'
Private Sub Z_GetPricesFromMarketStack(objSymsToLookup As Collection, objSymbolsWithData As Collection)

Dim sSymbol As Variant
Dim sCSV$, sAPIKey, sProxy$
Dim rCurrentPrice#, rDayLow#, rDayHigh#, rDayOpen#, rDayClose#
Dim objStock As cStock


    ' Check to see if there's anything to do
    On Error Resume Next
    If objSymsToLookup.Count > 0 Then
    
        ' Check we have a key
        sAPIKey = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_MARKET_STACK_KEY)
        If sAPIKey <> "" Then
            
            sProxy = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_PROXY)
            For Each sSymbol In objSymsToLookup
                rDayOpen = 0
                rDayHigh = 0
                rDayLow = 0
                rCurrentPrice = 0
        
                DoEvents
                PSGEN_Log "Getting stock price from Market Stack for " + sSymbol, LogEventTypes.LogInformation, EventIdTypes.StockPrices
                Call PSINET_GetHTTPFile("http://api.marketstack.com/v1/intraday/latest?access_key=" + sAPIKey + "&symbols=" + Replace(sSymbol, "^", "."), sCSV, sProxyName:=sProxy, lConnectionTimeout:=1000, lReadTimeout:=1000, iRetries:=2)
    
                ' Put the stock values into the lookup
                If Trim(sCSV) <> "" Then
                    PSGEN_Log "Got stock price from Market Stack successfully for " + sSymbol, LogEventTypes.LogInformation, EventIdTypes.StockPrices
                    rDayClose = CDbl(Split(Split(sCSV, """close"":", 2)(1), ",", 2)(0))
                    rDayHigh = CDbl(Split(Split(sCSV, """high"":", 2)(1), ",", 2)(0))
                    rDayLow = CDbl(Split(Split(sCSV, """low"":", 2)(1), ",", 2)(0))
                    rCurrentPrice = CDbl(Split(Split(sCSV, """last"":", 2)(1), ",", 2)(0))
    
                    Set objStock = New cStock
                    objStock.Code = sSymbol
                    objStock.CurrentPrice = rCurrentPrice
                    objStock.DayStart = rDayLow
                    objStock.DayHigh = rDayHigh
                    objStock.DayChange = rCurrentPrice - rDayClose
                    objStock.Source = "MarketStack"
                    objSymbolsWithData.Add objStock, sSymbol
                    objSymsToLookup.Remove sSymbol
                Else
                    PSGEN_Log "Failed to get stock price from Market Stack for " + sSymbol + " - " + Err.Description, LogEventTypes.LogError, EventIdTypes.StockPrices
                End If
            Next
        End If
    End If

End Sub

'
' objSymsToLookup    - Collection of ticket symbols to lookup
' objSymbolsWithData - Collection to poulate with cStock prices
'
' Reads the current stock prioces from the exchange and populates the collection with those values
'
Private Sub Z_GetPricesFromFinhub(objSymsToLookup As Collection, objSymbolsWithData As Collection)

Dim bag As JsonBag
Dim sSymbol As Variant
Dim sData$, sAPIKey, sProxy$, sName$
Dim rCurrentPrice#, rDayLow#, rDayHigh#, rDayOpen#, rDayClose#
Dim objStock As cStock


    ' Check to see if there's anything to do
    On Error Resume Next
    If objSymsToLookup.Count > 0 Then
    
        ' Check we have a key
        sAPIKey = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_FINHUB_KEY)
        If sAPIKey <> "" Then
            
            sProxy = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_PROXY)
            For Each sSymbol In objSymsToLookup
                rDayOpen = 0
                rDayHigh = 0
                rDayLow = 0
                rCurrentPrice = 0
        
                DoEvents
                sName = sSymbol
                If Not sSymbol Like "*.L" Then
                    PSGEN_Log "Getting stock price from Finnhub for " + sSymbol, LogEventTypes.LogInformation, EventIdTypes.StockPrices
                    Call PSINET_GetHTTPFile("https://finnhub.io/api/v1/quote?symbol=" + Replace(sSymbol, "^", ".") + "&token=" + sAPIKey, sData, sProxyName:=sProxy, lConnectionTimeout:=1000, lReadTimeout:=1000, iRetries:=2)
        
                    ' Put the stock values into the lookup
                    If Trim(sData) <> "" Then
                        PSGEN_Log "Got stock price from Finnhub successfully for " + sSymbol, LogEventTypes.LogInformation, EventIdTypes.StockPrices
                        Set bag = New JsonBag
                        bag.JSON = sData
                        rDayClose = CDbl(bag.Item("pc"))
                        rDayOpen = CDbl(bag.Item("o"))
                        rDayHigh = CDbl(bag.Item("h"))
                        rDayLow = CDbl(bag.Item("l"))
                        rCurrentPrice = CDbl(bag.Item("c"))
                        
                        Set objStock = New cStock
                        objStock.Code = sSymbol
                        objStock.CurrentPrice = rCurrentPrice
                        objStock.DayStart = rDayLow
                        objStock.DayHigh = rDayHigh
                        objStock.DayChange = rCurrentPrice - rDayClose
                        objStock.Source = "Finhub"
                        objSymbolsWithData.Add objStock, sSymbol
                        objSymsToLookup.Remove sSymbol
                    Else
                        PSGEN_Log "Failed to get stock price from Finnhub for " + sSymbol + " - " + Err.Description, LogEventTypes.LogError, EventIdTypes.StockPrices
                    End If
                End If
            Next
        End If
    End If

End Sub

'
' objSymsToLookup    - Collection of ticket symbols to lookup
' objSymbolsWithData - Collection to poulate with cStock prices
'
' Reads the current stock prioces from the exchange and populates the collection with those values
'
Private Sub Z_GetPricesFromTiingo(objSymsToLookup As Collection, objSymbolsWithData As Collection)

Dim bag As JsonBag
Dim sSymbol As Variant
Dim sData$, sAPIKey, sProxy$, sName$
Dim rCurrentPrice#, rDayLow#, rDayHigh#, rDayOpen#, rDayClose#
Dim objStock As cStock


    ' Check to see if there's anything to do
    On Error Resume Next
    If objSymsToLookup.Count > 0 Then
    
        ' Check we have a key
        sAPIKey = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_TIINGO_KEY)
        If sAPIKey <> "" Then
            
            sProxy = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_PROXY)
            For Each sSymbol In objSymsToLookup
                rDayOpen = 0
                rDayHigh = 0
                rDayLow = 0
                rCurrentPrice = 0
        
                DoEvents
                sName = sSymbol
                If Not sSymbol Like "*.L" Then
                    PSGEN_Log "Getting stock price from Tiingo for " + sSymbol, LogEventTypes.LogInformation, EventIdTypes.StockPrices
                    Call PSINET_GetHTTPFile("https://api.tiingo.com/iex/?tickers=" + Replace(sSymbol, "^", ".") + "&token=" + sAPIKey, sData, sProxyName:=sProxy, lConnectionTimeout:=1000, lReadTimeout:=1000, iRetries:=2)
        
                    ' Put the stock values into the lookup
                    If Trim(sData) <> "" Then
                        Err.Clear
                        PSGEN_Log "Got stock price from Tiingo successfully for " + sSymbol, LogEventTypes.LogInformation, EventIdTypes.StockPrices
                        Set bag = New JsonBag
                        bag.JSON = sData
                        Set bag = bag.Item(1)
                        If Err = 0 Then
                            rDayClose = CDbl(bag.Item("prevClose"))
                            rDayOpen = CDbl(bag.Item("open"))
                            rDayHigh = CDbl(bag.Item("high"))
                            rDayLow = CDbl(bag.Item("low"))
                            rCurrentPrice = CDbl(bag.Item("last"))
                            
                            Set objStock = New cStock
                            objStock.Code = sSymbol
                            objStock.CurrentPrice = rCurrentPrice
                            objStock.DayStart = rDayLow
                            objStock.DayHigh = rDayHigh
                            objStock.DayChange = rCurrentPrice - rDayClose
                            objStock.Source = "Tiingo"
                            objSymbolsWithData.Add objStock, sSymbol
                            objSymsToLookup.Remove sSymbol
                        End If
                    Else
                        PSGEN_Log "Failed to get stock price from Tiingo for " + sSymbol + " - " + Err.Description, LogEventTypes.LogError, EventIdTypes.StockPrices
                    End If
                End If
            Next
        End If
    End If
End Sub

'
' objSymsToLookup    - Collection of ticket symbols to lookup
' objSymbolsWithData - Collection to poulate with cStock prices
'
' Reads the current stock prioces from the exchange and populates the collection with those values
'
Private Sub Z_GetPricesFromYahoo(objSymsToLookup As Collection, objSymbolsWithData As Collection)

Dim bag As JsonBag
Dim sSymbol As Variant
Dim sData$, sProxy$, sName$
Dim rCurrentPrice#, rDayLow#, rDayHigh#, rDayOpen#, rDayClose#
Dim objStock As cStock


    ' Check to see if there's anything to do
    On Error Resume Next
    If objSymsToLookup.Count > 0 Then
    
        sProxy = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_PROXY)
        For Each sSymbol In objSymsToLookup
            rDayOpen = 0
            rDayHigh = 0
            rDayLow = 0
            rCurrentPrice = 0
    
            DoEvents
            PSGEN_Log "Getting stock price from Yahoo for " + sSymbol, LogEventTypes.LogInformation, EventIdTypes.StockPrices
            sName = sSymbol
            Call PSINET_GetHTTPFile("https://query2.finance.yahoo.com/ws/fundamentals-timeseries/v6/finance/quoteSummary/" + Replace(sName, "^", ".") + "?modules=price", sData, sProxyName:=sProxy, lConnectionTimeout:=2000, lReadTimeout:=2000)
                
            ' Put the stock values into the lookup
            If Trim(sData) <> "" Then
                PSGEN_Log "Got stock price from Yahoo successfully for " + sSymbol, LogEventTypes.LogInformation, EventIdTypes.StockPrices
                Set bag = New JsonBag
                bag.JSON = sData
                Set bag = bag.Item("quoteSummary").Item("result")(1).Item("price")
                rDayClose = CDbl(bag.Item("regularMarketPreviousClose").Item("fmt"))
                rDayHigh = CDbl(bag.Item("regularMarketDayHigh").Item("fmt"))
                rDayLow = CDbl(bag.Item("regularMarketDayLow").Item("fmt"))
                rCurrentPrice = CDbl(bag.Item("regularMarketPrice").Item("fmt"))
                If rCurrentPrice <> 0 Then
                    Set objStock = New cStock
                    objStock.Code = sSymbol
                    objStock.CurrentPrice = rCurrentPrice
                    objStock.DayStart = rDayLow
                    objStock.DayHigh = rDayHigh
                    objStock.DayChange = IIf(rDayOpen = 0, 0, rCurrentPrice - rDayClose)
                    objStock.Source = "Yahoo"
                    objSymbolsWithData.Add objStock, sSymbol
                    objSymsToLookup.Remove sSymbol
                Else
                    PSGEN_Log "Zero value returned from Yahoo for " + sSymbol, LogEventTypes.LogWarning, EventIdTypes.StockPrices
                End If
            Else
                PSGEN_Log "Failed to get stock prices from Yahoo for " + sSymbol + " - " + Err.Description, LogEventTypes.LogError, EventIdTypes.StockPrices
            End If
        Next
    End If
End Sub

'
' objSymsToLookup    - Collection of ticket symbols to lookup
' objSymbolsWithData - Collection to poulate with cStock prices
'
' Reads the current stock prioces from the exchange and populates the collection with those values
'
Private Sub Z_GetPricesFromReuters(objSymsToLookup As Collection, objSymbolsWithData As Collection)

Dim sSymbol As Variant
Dim sCSV$, sAPIKey, sProxy$
Dim rCurrentPrice#, rDayLow#, rDayHigh#, rDayOpen#, rDayClose#
Dim objStock As cStock


    ' Check to see if there's anything to do
    On Error Resume Next
    If objSymsToLookup.Count > 0 Then
    
        sProxy = mobjReg.GetSetting(App.Title, REG_SETTINGS, REG_PROXY)
        For Each sSymbol In objSymsToLookup
            rDayOpen = 0
            rDayHigh = 0
            rDayLow = 0
            rCurrentPrice = 0
    
            DoEvents
            PSGEN_Log "Getting stock price from Reuters for " + sSymbol, LogEventTypes.LogInformation, EventIdTypes.StockPrices
            Call PSINET_GetHTTPFile("https://in.reuters.com/companies/" + Replace(sSymbol, "^", "."), sCSV, sProxyName:=sProxy, lConnectionTimeout:=2000, lReadTimeout:=2000)

            ' Put the stock values into the lookup
            If InStr(sCSV, "<span>Open</span>") > 0 And Trim(sCSV) <> "" Then
                PSGEN_Log "Got stock price from Reurters successfully for " + sSymbol, LogEventTypes.LogInformation, EventIdTypes.StockPrices
                rDayOpen = CDbl(Split(Split(Split(sCSV, "<span>Prev Close</span>")(1), "<span>")(1), "<")(0))
                rDayHigh = CDbl(Split(Split(Split(Split(sCSV, "Today's High", 2)(1), "QuoteRibbon-digits-", 2)(1), ">", 2)(1), "<", 2)(0))
                rDayLow = CDbl(Split(Split(sCSV, "sectionQuoteDetailLow"">")(1), "<")(0))
                rCurrentPrice = CDbl(Split(Split(Split(sCSV, "QuoteRibbon-digits-", 2)(1), ">", 2)(1), "<", 2)(0))

                Set objStock = New cStock
                objStock.Code = sSymbol
                objStock.CurrentPrice = rCurrentPrice
                objStock.DayStart = rDayLow
                objStock.DayHigh = rDayHigh
                objStock.DayChange = rDayHigh - rDayLow
                objStock.Source = "Reuters"
                objSymbolsWithData.Add objStock, sSymbol
                objSymsToLookup.Remove sSymbol
            Else
                PSGEN_Log "Failed to get stock price from Reuters for " + sSymbol + " - " + Err.Description, LogEventTypes.LogError, EventIdTypes.StockPrices
            End If
        Next
    End If

End Sub






