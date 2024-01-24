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
Option Explicit

    Public Code$
    Public DisplayName$
    Public NumberOfShares&
    Public TotalCost#
    Public TotalValue#
    Public CurrentPrice#
    Public CurrencySymbol$
    Public CurrencyName$
    Public DayStart#
    Public DayChange#
    Public Position As New cRegion
    
    
Public Sub AddStock(ByVal lNumbeOfShares&, ByVal rPrice#)
    
    NumberOfShares = NumberOfShares + lNumbeOfShares
    TotalCost = TotalCost + (lNumbeOfShares * rPrice)
    TotalValue = lNumbeOfShares * CurrentPrice
    
End Sub

Public Function AverageCost#()
    
    AverageCost = TotalCost / IIf(NumberOfShares = 0, 1, NumberOfShares)

End Function

Public Function FormattedValue$()

    FormattedValue = FormatCurrencyValueWithSymbol(CurrencySymbol, CurrencyName, TotalValue)

End Function

Public Function FormattedCost$()

    FormattedCost = FormatCurrencyValueWithSymbol(CurrencySymbol, CurrencyName, TotalCost)

End Function

Public Function FormattedPrice$()

    FormattedPrice = FormatCurrencyValue(CurrencySymbol, CurrentPrice)

End Function

Public Function FormattedLoss$()

    FormattedLoss = FormatCurrencyValueWithSymbol(CurrencySymbol, CurrencyName, TotalValue - TotalCost)

End Function

Public Function FormattedLossPercent$()

    If TotalCost > 0 Then FormattedLossPercent = Format((TotalValue - TotalCost) / TotalCost, "0.00%")

End Function

Public Function LossPercent#()

    If TotalCost > 0 Then LossPercent = (TotalValue - TotalCost) * 100 / TotalCost

End Function

Public Function FormattedLossAdjusted$(ByVal rValue#)

    FormattedLossAdjusted = FormatCurrencyValueWithSymbol(CurrencySymbol, CurrencyName, TotalValue - rValue)

End Function

Public Function FormattedLossAdjustedPercent$(ByVal rValue#)

    If rValue > 0 Then FormattedLossAdjustedPercent = Format((TotalValue - rValue) / rValue, "0.00%")

End Function

Public Function FormattedAverageCost$()

    FormattedAverageCost = FormatCurrencyValue(CurrencySymbol, AverageCost)

End Function

Public Function FormattedBasePercent$()

    If AverageCost > 0 Then FormattedBasePercent = Format((CurrentPrice - AverageCost) / AverageCost, "0.00%")

End Function

Public Function FormattedTotalValue$()

    FormattedTotalValue = FormatCurrencyValue(CurrencySymbol, ConvertCurrency(Me, CurrentPrice) * NumberOfShares)

End Function



