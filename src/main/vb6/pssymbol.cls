VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSymbol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    Dim mobjReg As New cRegistry

    Public RegKey$
    Public Code$
    Public Alias$
    Public Price#
    Public CurrencyName$
    Public CurrencySymbol$
    Public Shares#
    Public ShowPrice As Boolean
    Public ShowChange As Boolean
    Public ShowChangePercent As Boolean
    Public ShowChangeUpDown As Boolean
    Public ShowProfitLoss As Boolean
    Public ShowDayChange As Boolean
    Public ShowDayChangePercent As Boolean
    Public ShowDayChangeUpDown As Boolean
    Public ExcludeFromSummary As Boolean
    Public Disabled As Boolean
    Public CurrentPrice#
    Public DayStart#
    Public DayChange#
    Public DayHigh#
    Public DayLow#
    Public ErrorDescription$
    Public Position As New cRegion
    Public LowAlarmEnabled As Boolean
    Public LowAlarmValue#
    Public LowAlarmIsPercent As Boolean
    Public LowAlarmSoundEnabled As Boolean
    Public HighAlarmEnabled As Boolean
    Public HighAlarmValue#
    Public HighAlarmIsPercent As Boolean
    Public HighAlarmSoundEnabled As Boolean
    Public AlarmShowing As Boolean
    Public FromNasdaqRealTime As Boolean
    Public LastUpdate As Date


Public Sub Init(ByVal sRegKey)

    On Error Resume Next
    RegKey = sRegKey
    Code = Trim(UCase(mobjReg.GetSetting(App.Title, REG_SYMBOLS + "\" + sRegKey, REG_SYMBOL, "")))
    Alias = Trim(mobjReg.GetSetting(App.Title, REG_SYMBOLS + "\" + sRegKey, REG_ALIAS, ""))
    If PSGEN_IsCommaLocale Then
        Price = CDbl(Replace(mobjReg.GetSetting(App.Title, REG_SYMBOLS + "\" + sRegKey, REG_PRICE, Format(0, "0.0")), ".", ","))
    Else
        Price = CDbl(Replace(mobjReg.GetSetting(App.Title, REG_SYMBOLS + "\" + sRegKey, REG_PRICE, Format(0, "0.0")), ",", "."))
    End If
    CurrencyName = UCase(Trim(mobjReg.GetSetting(App.Title, REG_SYMBOLS + "\" + sRegKey, REG_CURRENCY, "")))
    CurrencySymbol = Trim(mobjReg.GetSetting(App.Title, REG_SYMBOLS + "\" + sRegKey, REG_CURRENCY_SYMBOL, ""))
    Shares = CDbl(mobjReg.GetSetting(App.Title, REG_SYMBOLS + "\" + sRegKey, REG_SHARES, Format(0, "0.0")))
    ShowPrice = CBool(mobjReg.GetSetting(App.Title, REG_SYMBOLS + "\" + sRegKey, REG_SHOW_PRICE, "-1"))
    ShowChange = CBool(mobjReg.GetSetting(App.Title, REG_SYMBOLS + "\" + sRegKey, REG_SHOW_CHANGE, "0"))
    ShowChangePercent = CBool(mobjReg.GetSetting(App.Title, REG_SYMBOLS + "\" + sRegKey, REG_SHOW_CHANGE_PERCENT, "0"))
    ShowChangeUpDown = CBool(mobjReg.GetSetting(App.Title, REG_SYMBOLS + "\" + sRegKey, REG_SHOW_CHANGE_INDICATOR, "0"))
    ShowDayChange = CBool(mobjReg.GetSetting(App.Title, REG_SYMBOLS + "\" + sRegKey, REG_SHOW_DAY_CHANGE, "0"))
    ShowDayChangePercent = CBool(mobjReg.GetSetting(App.Title, REG_SYMBOLS + "\" + sRegKey, REG_SHOW_DAY_CHANGE_PERCENT, "0"))
    ShowDayChangeUpDown = CBool(mobjReg.GetSetting(App.Title, REG_SYMBOLS + "\" + sRegKey, REG_SHOW_DAY_CHANGE_INDICATOR, "0"))
    ShowProfitLoss = CBool(mobjReg.GetSetting(App.Title, REG_SYMBOLS + "\" + sRegKey, REG_SHOW_PROFIT_LOSS, "0"))
    ExcludeFromSummary = CBool(mobjReg.GetSetting(App.Title, REG_SYMBOLS + "\" + sRegKey, REG_EXCLUDE_FROM_SUMMARY, "0"))
    Disabled = CBool(mobjReg.GetSetting(App.Title, REG_SYMBOLS + "\" + sRegKey, REG_DISABLED, "0"))
    
    LowAlarmEnabled = CBool(mobjReg.GetSetting(App.Title, REG_SYMBOLS + "\" + sRegKey, REG_LOW_ALARM_ENABLED, "0"))
    LowAlarmIsPercent = CBool(mobjReg.GetSetting(App.Title, REG_SYMBOLS + "\" + sRegKey, REG_LOW_ALARM_AS_PERCENT, "0"))
    LowAlarmSoundEnabled = CBool(mobjReg.GetSetting(App.Title, REG_SYMBOLS + "\" + sRegKey, REG_LOW_ALARM_SOUND, "0"))
    If PSGEN_IsCommaLocale Then
        LowAlarmValue = CDbl(Replace(mobjReg.GetSetting(App.Title, REG_SYMBOLS + "\" + sRegKey, REG_LOW_ALARM_VALUE, Format(0, "0.0")), ".", ","))
    Else
        LowAlarmValue = CDbl(Replace(mobjReg.GetSetting(App.Title, REG_SYMBOLS + "\" + sRegKey, REG_LOW_ALARM_VALUE, Format(0, "0.0")), ",", "."))
    End If
    
    HighAlarmEnabled = CBool(mobjReg.GetSetting(App.Title, REG_SYMBOLS + "\" + sRegKey, REG_HIGH_ALARM_ENABLED, "0"))
    HighAlarmIsPercent = CBool(mobjReg.GetSetting(App.Title, REG_SYMBOLS + "\" + sRegKey, REG_HIGH_ALARM_AS_PERCENT, "0"))
    HighAlarmSoundEnabled = CBool(mobjReg.GetSetting(App.Title, REG_SYMBOLS + "\" + sRegKey, REG_HIGH_ALARM_SOUND, "0"))
    If PSGEN_IsCommaLocale Then
        HighAlarmValue = CDbl(Replace(mobjReg.GetSetting(App.Title, REG_SYMBOLS + "\" + sRegKey, REG_HIGH_ALARM_VALUE, Format(0, "0.0")), ".", ","))
    Else
        HighAlarmValue = CDbl(Replace(mobjReg.GetSetting(App.Title, REG_SYMBOLS + "\" + sRegKey, REG_HIGH_ALARM_VALUE, Format(0, "0.0")), ",", "."))
    End If
    
End Sub

Public Sub Save()

    On Error Resume Next
    mobjReg.SaveSetting App.Title, REG_SYMBOLS + "\" + RegKey, REG_SYMBOL, UCase(Trim(Code))
    mobjReg.SaveSetting App.Title, REG_SYMBOLS + "\" + RegKey, REG_ALIAS, Trim(Alias)
    mobjReg.SaveSetting App.Title, REG_SYMBOLS + "\" + RegKey, REG_PRICE, Format(Price)
    mobjReg.SaveSetting App.Title, REG_SYMBOLS + "\" + RegKey, REG_CURRENCY, UCase(Trim(CurrencyName))
    mobjReg.SaveSetting App.Title, REG_SYMBOLS + "\" + RegKey, REG_CURRENCY_SYMBOL, Trim(CurrencySymbol)
    mobjReg.SaveSetting App.Title, REG_SYMBOLS + "\" + RegKey, REG_SHARES, Format(Shares)
    mobjReg.SaveSetting App.Title, REG_SYMBOLS + "\" + RegKey, REG_SHOW_PRICE, Format(ShowPrice)
    mobjReg.SaveSetting App.Title, REG_SYMBOLS + "\" + RegKey, REG_SHOW_CHANGE, Format(ShowChange)
    mobjReg.SaveSetting App.Title, REG_SYMBOLS + "\" + RegKey, REG_SHOW_CHANGE_PERCENT, Format(ShowChangePercent)
    mobjReg.SaveSetting App.Title, REG_SYMBOLS + "\" + RegKey, REG_SHOW_CHANGE_INDICATOR, Format(ShowChangeUpDown)
    mobjReg.SaveSetting App.Title, REG_SYMBOLS + "\" + RegKey, REG_SHOW_DAY_CHANGE, Format(ShowDayChange)
    mobjReg.SaveSetting App.Title, REG_SYMBOLS + "\" + RegKey, REG_SHOW_DAY_CHANGE_PERCENT, Format(ShowDayChangePercent)
    mobjReg.SaveSetting App.Title, REG_SYMBOLS + "\" + RegKey, REG_SHOW_DAY_CHANGE_INDICATOR, Format(ShowDayChangeUpDown)
    mobjReg.SaveSetting App.Title, REG_SYMBOLS + "\" + RegKey, REG_SHOW_PROFIT_LOSS, Format(ShowProfitLoss)
    mobjReg.SaveSetting App.Title, REG_SYMBOLS + "\" + RegKey, REG_EXCLUDE_FROM_SUMMARY, Format(ExcludeFromSummary)
    mobjReg.SaveSetting App.Title, REG_SYMBOLS + "\" + RegKey, REG_DISABLED, Format(Disabled)
    
    mobjReg.SaveSetting App.Title, REG_SYMBOLS + "\" + RegKey, REG_LOW_ALARM_ENABLED, Format(LowAlarmEnabled)
    mobjReg.SaveSetting App.Title, REG_SYMBOLS + "\" + RegKey, REG_LOW_ALARM_AS_PERCENT, Format(LowAlarmIsPercent)
    mobjReg.SaveSetting App.Title, REG_SYMBOLS + "\" + RegKey, REG_LOW_ALARM_SOUND, Format(LowAlarmSoundEnabled)
    mobjReg.SaveSetting App.Title, REG_SYMBOLS + "\" + RegKey, REG_LOW_ALARM_VALUE, Format(LowAlarmValue)
    
    mobjReg.SaveSetting App.Title, REG_SYMBOLS + "\" + RegKey, REG_HIGH_ALARM_ENABLED, Format(HighAlarmEnabled)
    mobjReg.SaveSetting App.Title, REG_SYMBOLS + "\" + RegKey, REG_HIGH_ALARM_AS_PERCENT, Format(HighAlarmIsPercent)
    mobjReg.SaveSetting App.Title, REG_SYMBOLS + "\" + RegKey, REG_HIGH_ALARM_SOUND, Format(HighAlarmSoundEnabled)
    mobjReg.SaveSetting App.Title, REG_SYMBOLS + "\" + RegKey, REG_HIGH_ALARM_VALUE, Format(HighAlarmValue)
    
End Sub

Public Function DisplayName$()

    DisplayName = IIf(Alias = "", Code, Alias)

End Function

Public Function FormattedValue$()

    FormattedValue = FormatCurrencyValue(CurrencySymbol, CurrentPrice)

End Function

Public Function FormattedTotalValue$()

    FormattedTotalValue = FormatCurrencyValue(CurrencySymbol, CurrentPrice * Shares)

End Function

Public Function FormattedCost$()

    FormattedCost = FormatCurrencyValue(CurrencySymbol, Price)

End Function

Public Function FormattedTotalCost$()

    FormattedTotalCost = FormatCurrencyValue(CurrencySymbol, Price * Shares)

End Function

Public Function SortKey$()

    SortKey = Code & Space(10 - Len(Code)) & RegKey & Space(20 - Len(RegKey))

End Function


