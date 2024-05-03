VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Copyright (c) 2024, Pivotal Solutions and/or its affiliates. All rights reserved.
' Pivotal Solutions PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
'
' Contains the runtime value of a stock and provides calculation methods for displaying those values
' This class copes with the idea of 'aggregated transactions' i.e. multiple transactions for
' the same stock but at different prices and quantities
'
Option Explicit

    Public Code$
    Public DisplayName$
    Public NumberOfShares#
    Public TotalCost#
    Public TotalValue#
    Public CurrentPrice#
    Public CurrencySymbol$
    Public CurrencyName$
    Public DayStart#
    Public DayChange#
    Public Position As New cRegion
    
    
Public Sub AddStock(ByVal rNumbeOfShares#, ByVal rPrice#)
'
' Adds the specified amount of sahes by aggregating the total cost
'
' lNumbeOfShares& - Number of shares bought
' rPrice#         - Price paid for these shares
'
    
    NumberOfShares = NumberOfShares + rNumbeOfShares
    TotalCost = TotalCost + (rNumbeOfShares * rPrice)
    TotalValue = rNumbeOfShares * CurrentPrice
    
End Sub

Public Function AverageCost#()
'
' Returns the average price of the shares
'
' RETURN Total cost divided by the total shares
'
    AverageCost = TotalCost / IIf(NumberOfShares = 0, 1, NumberOfShares)

End Function

Public Function FormattedValue$()
'
' Returns the value of the stock formatted using the currency of the stock
'
' RETURN Formatted value
'
    FormattedValue = FormatCurrencyValueWithSymbol(CurrencySymbol, CurrencyName, TotalValue)

End Function

Public Function FormattedCost$()
'
' Returns the cost of the stock formatted using the currency of the stock
'
' RETURN Formatted cost
'
    FormattedCost = FormatCurrencyValueWithSymbol(CurrencySymbol, CurrencyName, TotalCost)

End Function

Public Function FormattedPrice$()
'
' Returns the current price of the stock formatted using the currency of the stock
'
' RETURN Formatted price
'

    FormattedPrice = FormatCurrencyValue(CurrencySymbol, CurrentPrice)

End Function

Public Function FormattedLoss$()
'
' Returns the current loss on the stock formatted using the currency of the stock
'
' RETURN Formatted loss/gain
'
    FormattedLoss = FormatCurrencyValueWithSymbol(CurrencySymbol, CurrencyName, TotalValue - TotalCost)

End Function

Public Function FormattedLossPercent$()
'
' Returns the current loss/gain of the stock formatted as a perecentage of the original cost
'
' RETURN Formatted loss/gain percentage
'

    If TotalCost > 0 Then FormattedLossPercent = Format((TotalValue - TotalCost) / TotalCost, "0.00%")

End Function

Public Function LossPercent#()
'
' Returns the current loss/gain of the stock as a perecentage of the original cost
'
' RETURN Loss/gain as percent
'

    If TotalCost > 0 Then LossPercent = (TotalValue - TotalCost) * 100 / TotalCost

End Function

Public Function FormattedLossAdjusted$(ByVal rValue#)
'
' Returns the current loss/gain of the stock formatted using the currency of the stock after adjusting for and initial investment
'
' rValue - Initial investment
'
' RETURN Loss/gain as percent
'

    FormattedLossAdjusted = FormatCurrencyValueWithSymbol(CurrencySymbol, CurrencyName, TotalValue - rValue)

End Function

Public Function FormattedLossAdjustedPercent$(ByVal rValue#)
'
' Returns the current loss/gain of the stock formatted as a perecentage of the original cost after adjusting for and initial investment
'
' rValue - Initial investment
'
' RETURN Formatted loss/gain percentage
'

    If rValue > 0 Then FormattedLossAdjustedPercent = Format((TotalValue - rValue) / rValue, "0.00%")

End Function

Public Function FormattedAverageCost$()
'
' Returns the average cost of the stock formatted using the currency of the stock
'
' RETURN Formatted average cost
'

    FormattedAverageCost = FormatCurrencyValue(CurrencySymbol, AverageCost)

End Function

Public Function FormattedBasePercent$()
'
' Returns the percentage differece between the average cost (base) and the current price
'
' RETURN Formatted price difference
'

    If AverageCost > 0 Then FormattedBasePercent = Format((CurrentPrice - AverageCost) / AverageCost, "0.00%")

End Function

Public Function FormattedTotalValue$()
'
' Returns the total value of the stock formatted using the currency of the stock and converted to the local currency
'
' RETURN Formatted total value in local currency
'
    
    FormattedTotalValue = FormatCurrencyValue(CurrencySymbol, ConvertCurrency(Me, CurrentPrice) * NumberOfShares)

End Function



