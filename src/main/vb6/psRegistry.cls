VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRegistry"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************
'
'   Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                LANGUAGE : Microsoft Visual Basic V6
'
'              MODULENAME : Pivotal_Registry
'
'             MODULE TYPE : BASIC Class
'
'                FILENAME : nlRegistry.bas
'
'    MODIFICATION HISTORY : Darren Adams    24 August 1999  First created for Pivotal Object Library
'
'                 PURPOSE : This class modules provides services that allow applications
'                           manipulate the Windows registry (note: All registry keys are
'                           derived from REGISTRY_ROOT\SOFTWARE\Pivotal).
'
'****************************************************************************
'
Option Explicit

'********************************
' WINDOWS DECLARATIONS
'********************************

'**********
' Windows API Types
'**********
Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

'**********
' Windows API Constants
'**********
Public Enum RegistryClasses
    RegRoot = &H80000000
    RegCurrentUser = &H80000001
    RegLocalMachine = &H80000002
    RegPerformanceData = &H80000004
    RegUsers = &H80000003
End Enum
    
Private REGISTRY_ROOT As RegistryClasses
    
Private Const ERROR_SUCCESS = 0&                    ' Appropriate error codes.
Private Const ERROR_INSUFFICIENT_BUFFER = 122
Private Const ERROR_NO_MORE_ITEMS = 259&
Private Const ERROR_MORE_DATA = 234

Private Const REG_CREATED_NEW_KEY = &H1             ' Values "returned" by the
Private Const REG_OPENED_EXISTING_KEY = &H2         ' RegCreateKeyEx API functions.

Const REG_BINARY = 3                        ' Data types of the information retrieved
Const REG_DWORD = 4                         ' from the registry.
Const REG_SZ = 1
Const REG_NONE = 0
Const REG_MULTI_SZ = 7

Private Const REG_OPTION_BACKUP_RESTORE = 4
Private Const REG_OPTION_CREATE_LINK = 2
Private Const REG_OPTION_NON_VOLATILE = 0
Private Const REG_OPTION_RESERVED = 0
Private Const REG_OPTION_VOLATILE = 1

Private Const KEY_QUERY_VALUE = &H1                 ' Registry key access constants
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const KEY_READ = ((KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

'**********
' Windows API Functions
'**********
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegEnumValueXP Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteTree Lib "advapi32.dll" Alias "RegDeleteTreeA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDisableReflectionKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

'********************************
' CLASS DECLARATIONS
'********************************

'**********
' Constants
'**********
Private Const REGISTRY_KEY_BASE As String = "SOFTWARE\"
Private Const DEF_COMPANY_NAME As String = "Pivotal"
Public CompanyName$


Private Enum PlatformIDConstants
    pi_Windows31 = 0
    pi_Windows9x = 1
    pi_WindowsNT = 2
End Enum

Private Enum MajorVersionConstants
    mjv_Windows3x = 3
    mjv_Windows9xNT = 4
    mjv_Windows2KXP = 5
End Enum

Private Enum MinorVersionConstants
    mnv_Windows2K = 0
    mnv_WindowsXP = 1
End Enum

'**********
' Types
'**********

'
' A private type that allows us to store registry
' key values (only strings at the moment)
'
Private Type REGISTRYKEY
    sName As String
    sValue As String
End Type

'
' Static Operating System Version Information
'
Private mbRetrievedVersion As Boolean
Private mstVersionInformation As OSVERSIONINFO

Public Sub DeleteSetting(ByVal sAppName$, Optional ByVal sSection$ = "", Optional ByVal sKey$ = "")
Attribute DeleteSetting.VB_Description = "Deletes a section or key setting from an application's entry in the Windows registry"
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME: Sub DeleteSetting
'
'                     sAppName$          - String expression containing the name of the application or project to which the section or key setting applies.
'                     sSection$          - String expression containing the name of the section where the key setting is being deleted.
'                     sKey$              - String expression containing the name of the key setting being deleted.
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Darren Adams    25 August 1999   First created for Pivotal Object Library
'
'                  PURPOSE: Deletes a section or key setting from an application's entry in the
'                           Windows registry.
'
'****************************************************************************
'
'

'
' The required registry variables
'
Dim lRegistryKey&, lReturnCode&

    '
    ' If we have a key and section, then dlete the key value
    '
    On Error Resume Next
    If sKey <> "" Then
        lReturnCode = RegOpenKeyEx(REGISTRY_ROOT, REGISTRY_KEY_BASE & CompanyName & "\" & sAppName & IIf(sSection <> "", "\" & sSection, ""), 0&, KEY_ALL_ACCESS, lRegistryKey)
        If lReturnCode = ERROR_SUCCESS Then
            lReturnCode = RegDeleteValue(lRegistryKey, sKey)
            Call RegCloseKey(lRegistryKey)
        End If
    Else
        lReturnCode = RegDeleteTree(REGISTRY_ROOT, REGISTRY_KEY_BASE & CompanyName & "\" & sAppName & IIf(sSection <> "", "\" & sSection, ""))
    End If

End Sub

Public Sub DeleteSettingEx(ByVal lKeyBase As RegistryClasses, ByVal sSection$, Optional ByVal sKey$ = "")
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME: Sub DeleteSettingEx
'
'                     lKeyBase           - The ID of the main key
'                     sSection$          - String expression containing the name of the section where the key setting is being deleted.
'                     sKey$              - String expression containing the name of the key setting being deleted.
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    19th November 2001   First created for Pivotal DDL Editor
'
'                  PURPOSE: Deletes a section or key setting from base entry in the
'                           Windows registry.
'
'****************************************************************************
'
'
'
' The required registry variables
'
Dim lRegistryKey&, lReturnCode&

    '
    ' If we have a key and section, then dlete the key value
    '
    On Error Resume Next
    If sKey <> "" Then
        lReturnCode = RegOpenKeyEx(lKeyBase, sSection, 0&, KEY_ALL_ACCESS, lRegistryKey)
        If lReturnCode = ERROR_SUCCESS Then
            lReturnCode = RegDeleteValue(lRegistryKey, sKey)
            Call RegCloseKey(lRegistryKey)
        End If
    Else
        lReturnCode = RegDeleteTree(lKeyBase, sSection)
    End If

End Sub

Public Sub SaveSetting(ByVal sAppName$, ByVal sSection$, ByVal sKey$, ByVal vSetting As Variant)
Attribute SaveSetting.VB_Description = "Saves or creates an application entry in the application's entry in the Windows registry."
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME: Sub gobjReg.SaveSetting
'
'                     sAppName$          - String expression containing the name of the application or project to which the setting applies
'                     sSection$          - String expression containing the name of the section where the key setting is being saved
'                     sKey$              - String expression containing the name of the key setting being saved
'                     vSetting           - Expression containing the value they 'key' is being set to.
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Darren Adams    25 August 1999   First created for Pivotal Object Library
'
'                  PURPOSE: Saves or creates an application entry in the
'
'****************************************************************************
'
'

'
' Required registry variables
'
Dim lReturnCode&, lRegistryKey&, sRegValue$

'
' Used to store a value that indicates how a
' registry key was created
'
Dim lHowKeyCreated&

'
' Security attributes required to create (or open) a
' registry key.
'
Dim saAttr As SECURITY_ATTRIBUTES

    '
    ' Initialise error vector
    '
    On Error Resume Next
    
    '
    ' Set up the Security Attributes structure, then open (
    ' or create, it doesn't matter) the specified registry key, then
    ' save the value
    '
    saAttr.nLength = 50                 ' I don't know whether these values are correct, they were taken
    saAttr.lpSecurityDescriptor = 0     ' from the Visual Basic Screensaver project. The MSDN help can
    saAttr.bInheritHandle = True        ' probably give you advice on what these values can (or should) be.
    
    lReturnCode = RegCreateKeyEx(REGISTRY_ROOT, REGISTRY_KEY_BASE & CompanyName & "\" & sAppName & IIf(sAppName = "", "", "\") & sSection, _
        0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, saAttr, lRegistryKey, lHowKeyCreated)

    '
    ' If the registry key can be created, then convert the
    ' "vSetting" argument to a string, then write the
    ' registry key value.
    '
    If lReturnCode = ERROR_SUCCESS Then
        sRegValue = CStr(vSetting) + vbNullChar
        
        lReturnCode = RegSetValueEx(lRegistryKey, sKey, 0&, REG_SZ, ByVal sRegValue, Len(sRegValue))
        If lReturnCode <> ERROR_SUCCESS Then
            '
            ' Something has happened. Make a note of the return value
            ' "lReturnCode". You can find out what the error code
            ' means by looking for the Windows Base Services in the
            ' Platform SDK section in the MSDN help.
            '
        End If
    Else
        '
        ' The Visual Basic help files says that an error is
        ' raised if the registry key could not be saved, though
        ' I have never seen this...
        '
    End If
    
    '
    ' Close the registry key anyway
    '
    Call RegCloseKey(lRegistryKey)
End Sub

Public Sub SaveSettingEx(ByVal lKeyBase As RegistryClasses, ByVal sSection$, ByVal sKey$, ByVal vSetting As Variant)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME: Sub gobjReg.SaveSetting
'
'                     lKeyBase           - The ID of the main key
'                     sSection$          - String expression containing the name of the section where the key setting is being saved
'                     sKey$              - String expression containing the name of the key setting being saved
'                     vSetting           - Expression containing the value they 'key' is being set to.
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    19th November 2001   First created for Pivotal DDL Editor
'
'                  PURPOSE: Saves or creates an application entry in the
'
'****************************************************************************
'
'

Dim lReturnCode&, lRegistryKey&, sRegValue$
Dim lHowKeyCreated&
Dim saAttr As SECURITY_ATTRIBUTES

    '
    ' Set up the Security Attributes structure, then open (
    ' or create, it doesn't matter) the specified registry key, then
    ' save the value
    '
    On Error Resume Next
    saAttr.nLength = 50                 ' I don't know whether these values are correct, they were taken
    saAttr.lpSecurityDescriptor = 0      ' from the Visual Basic Screensaver project. The MSDN help can
    saAttr.bInheritHandle = True        ' probably give you advice on what these values can (or should) be.
    
    lReturnCode = RegCreateKeyEx(lKeyBase, sSection, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, saAttr, lRegistryKey, lHowKeyCreated)

    '
    ' If the registry key can be created, then convert the
    ' "vSetting" argument to a string, then write the
    ' registry key value.
    '
    If lReturnCode = ERROR_SUCCESS Then
        sRegValue = CStr(vSetting) + vbNullChar
        
        lReturnCode = RegSetValueEx(lRegistryKey, sKey, 0&, REG_SZ, ByVal sRegValue, Len(sRegValue))
        If lReturnCode <> ERROR_SUCCESS Then
            '
            ' Something has happened. Make a note of the return value
            ' "lReturnCode". You can find out what the error code
            ' means by looking for the Windows Base Services in the
            ' Platform SDK section in the MSDN help.
            '
        End If
    Else
        '
        ' The Visual Basic help files says that an error is
        ' raised if the registry key could not be saved, though
        ' I have never seen this...
        '
    End If
    
    '
    ' Close the registry key anyway
    '
    Call RegCloseKey(lRegistryKey)
End Sub

Public Function GetAllSettings(ByVal sAppName$, ByVal sSection$) As Variant
Attribute GetAllSettings.VB_Description = "Retrieves all settings from a specific registry key"
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME: Function GetAllSettings
'
'                     sAppName$          - The name of the application whose settings are requested
'                     sSection$          - The name of the section whose key settings are requested
'
'                          ) As Integer
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Darren Adams    25 August 1999   First created for Pivotal Object Library
'
'                  PURPOSE:
'
'****************************************************************************
'
'

'
' A variant array that will be used to return a one registry keys'
' worth of values
'
Dim vRegistryKeyValues() As Variant
Dim rkKeys() As REGISTRYKEY

'
' The usual registry manipulation variables
'
Dim lReturnValue&, lRegistryKey&, lRegistryKeyLength&
Dim lNextValue&, sKeyName$, sKeyValue$, lKeyNameSize&, lKeyValueSize&
Dim ftFileTime As FILETIME, abKeyValue() As Byte

    '
    ' Initialise error vector
    '
    On Error Resume Next
    
    '
    ' Find the registry key
    '
    lReturnValue = RegOpenKeyEx(REGISTRY_ROOT, REGISTRY_KEY_BASE & CompanyName & "\" & sAppName & "\" & sSection, 0, KEY_READ, lRegistryKey)
        
    '
    ' Found the registry key? Then get the values, otherwise, return
    ' an unitialised variant (Empty)
    '
    If lReturnValue = ERROR_SUCCESS Then

        If (Z_IsWindowsXP()) Then
            GetAllSettings = Z_GetAllSettingsInXP(lRegistryKey)
        Else
            '
            ' Start finding registry key names and values
            '
            lNextValue = 0&
            lKeyNameSize = 32000: sKeyName = String$(lKeyNameSize + 1, vbNullChar)
            lKeyValueSize = 32000: ReDim abKeyValue(1 To lKeyValueSize + 1)
            lReturnValue = RegEnumValue(lRegistryKey, lNextValue, sKeyName, lKeyNameSize, 0&, REG_SZ, abKeyValue(1), lKeyValueSize)
            
            Do While lReturnValue = ERROR_SUCCESS
                sKeyName = Left$(sKeyName, lKeyNameSize)
                sKeyValue = StrConv(abKeyValue, vbUnicode)
                sKeyValue = Left$(sKeyValue, lKeyValueSize - 1)
                
                '
                ' Make the array just a little bit bigger to store these new values
                '
                ReDim Preserve rkKeys(lNextValue)
                rkKeys(lNextValue).sName = sKeyName
                rkKeys(lNextValue).sValue = sKeyValue
                lNextValue = lNextValue + 1
                lKeyNameSize = 32000: sKeyName = String$(lKeyNameSize + 1, vbNullChar)
                lKeyValueSize = 32000: ReDim abKeyValue(1 To lKeyValueSize + 1)
                
                lReturnValue = RegEnumValue(lRegistryKey, lNextValue, sKeyName, lKeyNameSize, 0&, REG_SZ, abKeyValue(1), lKeyValueSize)
                DoEvents
            Loop
            
            '
            ' Now that the registry keys have been retrieved, store them into the
            ' variant array
            '
            If lNextValue > 0 Then
                ReDim Preserve vRegistryKeyValues(UBound(rkKeys), 1)
                For lNextValue = 0 To UBound(rkKeys)
                    vRegistryKeyValues(lNextValue, 0) = rkKeys(lNextValue).sName
                    vRegistryKeyValues(lNextValue, 1) = rkKeys(lNextValue).sValue
                Next
                GetAllSettings = vRegistryKeyValues
            Else
                '
                ' No registry keys; return an unitialised (Empty) variant
                '
                GetAllSettings = Empty
            End If
        End If
    Else
        GetAllSettings = Empty
    End If

    '
    ' Close the registry key anyway
    '
    Call RegCloseKey(lRegistryKey)
End Function


Public Function GetSubkeys(ByVal sAppName$, ByVal sSection$) As Variant
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME: Function GetSubkeys
'
'                     sAppName$          - The name of the application whose settings are requested
'                     sSection$          - The name of the section whose key settings are requested
'
'                          ) As Integer
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Darren Adams    25 August 1999   First created for Pivotal Object Library
'
'                  PURPOSE:
'
'****************************************************************************
'
'

'
' A variant array that will be used to return a one registry keys'
' worth of values
'
Dim rkKeys$()

'
' The usual registry manipulation variables
'
Dim lReturnValue&, lRegistryKey&, lRegistryKeyLength&
Dim lNextValue&, sKeyName$, lKeyNameSize&
Dim ftFileTime As FILETIME

    '
    ' Initialise error vector
    '
    On Error Resume Next
    
    '
    ' Find the registry key
    '
    lReturnValue = RegOpenKeyEx(REGISTRY_ROOT, REGISTRY_KEY_BASE & CompanyName & "\" & sAppName & "\" & sSection, 0, KEY_READ, lRegistryKey)
        
    '
    ' Found the registry key? Then get the values, otherwise, return
    ' an unitialised variant (Empty)
    '
    If lReturnValue = ERROR_SUCCESS Then

        '
        ' Start finding registry key names and values
        '
        lNextValue = 0&
        lKeyNameSize = 32000: sKeyName = String$(lKeyNameSize + 1, vbNullChar)
        lReturnValue = RegEnumKeyEx(lRegistryKey, lNextValue, sKeyName, lKeyNameSize, 0&, vbNullString, 0&, ftFileTime)
        If lReturnValue = ERROR_SUCCESS Then
            Do While lReturnValue = ERROR_SUCCESS
                sKeyName = Left$(sKeyName, lKeyNameSize)
                
                '
                ' Make the array just a little bit bigger to store these new values
                '
                ReDim Preserve rkKeys(lNextValue)
                rkKeys(lNextValue) = sKeyName
                lNextValue = lNextValue + 1
                lKeyNameSize = 32000: sKeyName = String$(lKeyNameSize + 1, vbNullChar)
                lReturnValue = RegEnumKeyEx(lRegistryKey, lNextValue, sKeyName, lKeyNameSize, 0&, vbNullString, 0&, ftFileTime)
                DoEvents
            Loop
            GetSubkeys = rkKeys
        Else
            GetSubkeys = Empty
        End If
    Else
        GetSubkeys = Empty
    End If

    '
    ' Close the registry key anyway
    '
    Call RegCloseKey(lRegistryKey)

End Function



Public Function GetAllSettingsEx(ByVal lKeyBase As RegistryClasses, ByVal sSection$) As Variant
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME: Function GetAllSettingsEx
'
'                     lKeyBase           - The ID of the main key
'                     sSection$          - The name of the section whose key settings are requested
'
'                          ) As Integer
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    19th November 2001   First created for Pivotal DDL Editor
'
'                  PURPOSE: Returns an array of settings from the registry
'
'****************************************************************************
'
'

Dim vRegistryKeyValues() As Variant
Dim rkKeys() As REGISTRYKEY
Dim lReturnValue&, lRegistryKey&, lRegistryKeyLength&
Dim lNextValue&, sKeyName$, sKeyValue$, lKeyNameSize&, lKeyValueSize&
Dim ftFileTime As FILETIME, abKeyValue() As Byte

    '
    ' Find the registry key
    '
    On Error Resume Next
    lReturnValue = RegOpenKeyEx(lKeyBase, sSection, 0, KEY_READ, lRegistryKey)
        
    '
    ' Found the registry key? Then get the values, otherwise, return
    ' an unitialised variant (Empty)
    '
    If lReturnValue = ERROR_SUCCESS Then
        If (Z_IsWindowsXP()) Then
            GetAllSettingsEx = Z_GetAllSettingsInXP(lRegistryKey)
        Else
            
            '
            ' Start finding registry key names and values
            '
            lNextValue = 0&
            lKeyNameSize = 32000: sKeyName = String$(lKeyNameSize + 1, vbNullChar)
            lKeyValueSize = 32000: ReDim abKeyValue(1 To lKeyValueSize + 1)
            
            lReturnValue = RegEnumValue(lRegistryKey, lNextValue, sKeyName, lKeyNameSize, 0&, REG_SZ, abKeyValue(1), lKeyValueSize)
            Do While lReturnValue = ERROR_SUCCESS
                sKeyName = Left$(sKeyName, lKeyNameSize)
                sKeyValue = StrConv(abKeyValue, vbUnicode)
                sKeyValue = Left$(sKeyValue, lKeyValueSize - 1)
                
                '
                ' Make the array just a little bit bigger to store these new values
                '
                ReDim Preserve rkKeys(lNextValue)
                rkKeys(lNextValue).sName = sKeyName
                rkKeys(lNextValue).sValue = sKeyValue
                lNextValue = lNextValue + 1
                lKeyNameSize = 32000: sKeyName = String$(lKeyNameSize + 1, vbNullChar)
                lKeyValueSize = 32000: ReDim abKeyValue(1 To lKeyValueSize + 1)
                
                lReturnValue = RegEnumValue(lRegistryKey, lNextValue, sKeyName, lKeyNameSize, 0&, REG_SZ, abKeyValue(1), lKeyValueSize)
                DoEvents
            Loop
            
            '
            ' Now that the registry keys have been retrieved, store them into the
            ' variant array
            '
            If lNextValue > 0 Then
                ReDim Preserve vRegistryKeyValues(UBound(rkKeys), 1)
                For lNextValue = 0 To UBound(rkKeys)
                    vRegistryKeyValues(lNextValue, 0) = rkKeys(lNextValue).sName
                    vRegistryKeyValues(lNextValue, 1) = rkKeys(lNextValue).sValue
                Next
                GetAllSettingsEx = vRegistryKeyValues
            Else
                '
                ' No registry keys; return an unitialised (Empty) variant
                '
                GetAllSettingsEx = Empty
            End If
        End If
    Else
        GetAllSettingsEx = Empty
    End If
    
    '
    ' Close the registry key anyway
    '
    Call RegCloseKey(lRegistryKey)

End Function

Private Function Z_GetAllSettingsInXP(hKeyRegistry&)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME: GetAllSettingsInXP
'                           (
'                               hKeyRegistry&   -   The open key into the registry
'                           ) as Variant :  An array of entries
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Mark Travis     25 February, 2002
'
'                  PURPOSE: Returns an array of settings from the registry
'
'****************************************************************************
'
'
Dim vSettings As Variant: vSettings = Empty
Dim avSettings() As Variant

Dim astSettings() As REGISTRYKEY
Dim lNextValue&: lNextValue = 0
Dim lNextSetting&: lNextSetting = 0
Dim lIndex&

Dim sName$, sValue$

    '
    ' Initialise the error vector
    '
    On Error Resume Next

    '
    ' Loop through all the entries for the key
    '
    While (ERROR_NO_MORE_ITEMS <> Z_GetEnumValueInXP(hKeyRegistry, lNextValue, sName, sValue))
        '
        ' Don't store the default value for the sent key
        '
        If (sName <> vbNullString) Then
            '
            ' Increase the size of the return data array
            '
            ReDim Preserve astSettings(lNextSetting)
            
            '
            ' Store the returned name and value in the array
            '
            astSettings(lNextSetting).sName = sName
            astSettings(lNextSetting).sValue = sValue
            
            '
            ' Increment the next setting index
            '
            lNextSetting = lNextSetting + 1
        End If
        
        '
        ' Increment the next index value
        '
        lNextValue = lNextValue + 1
    Wend
    
    If (lNextValue > 0) Then
        '
        ' Create a 2D variant array to store the name-value pairs in
        '
        ReDim avSettings(LBound(astSettings) To UBound(astSettings), 0 To 1)
        
        '
        ' Copy all the settings into the variant array
        '
        For lIndex = LBound(astSettings) To UBound(astSettings)
            avSettings(lIndex, 0) = astSettings(lIndex).sName
            avSettings(lIndex, 1) = astSettings(lIndex).sName
        Next lIndex
        
        '
        ' Store the 2D variant array into a variant to return
        '
        vSettings = avSettings
    End If

    Z_GetAllSettingsInXP = vSettings

End Function

Private Function Z_GetEnumValueInXP&(ByVal hKeyRegistry&, ByVal lValueIndex&, sName$, sValue$)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME: GetAllSettingsInXP
'                           (
'                               hKeyRegistry&   -   The open key into the registry
'                               lValueIndex&    -   The enumerated index of the value to retrieve
'                               sName$          -   The returned enum name
'                               sValue$         -   The returned enum value
'                           ) as Long : The return value from the calls to the registry api
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Mark Travis     25 February, 2002
'
'                  PURPOSE: Returns the name and value of the next enumerated entry
'                           in the registry
'
'****************************************************************************
'
'
Dim lResult&

Dim lName&: lName = 4000
Dim lValue&: lValue = 4000
Dim lType&

    '
    ' Initialise the error vector
    '
    On Error Resume Next

    '
    ' Initialise the buffer sizes for the key and the data
    '
    sName = String$(lName + 1, vbNullChar)
    sValue = String$(lValue + 1, vbNullChar)

    '
    ' Get the name and type of the value to be retrieved
    '
    lResult = RegEnumValueXP(hKeyRegistry, lValueIndex, ByVal 0, 0, 0&, lType, 0, 0)

    '
    ' Get the value depending on the type
    ' And then format/truncate the value into a variable length string
    '
    Select Case (lType)
        Case REG_DWORD ', REG_DWORD_LITTLE_ENDIAN
            lResult = RegEnumValueXP(hKeyRegistry, lValueIndex, sName, lName, 0&, lType, _
                                        lValue, 4)
            sValue = Format$(lValue)

        '**
        ' Other types to be handled here later
        '**

        Case Else
            lResult = RegEnumValueXP(hKeyRegistry, lValueIndex, sName, lName, 0&, lType, _
                                        ByVal sValue, lValue)
            If (lValue > 0) Then
                sValue = Left$(sValue, lValue - 1)
            Else
                sValue = vbNullString
            End If

    End Select

    '
    ' Truncate the name value into a variable length string
    '
    If (lName > 0) Then
        sName = Left$(sName, lName)
    Else
        sName = vbNullString
    End If

    Z_GetEnumValueInXP = lResult

End Function

Public Function GetSetting$(ByVal sAppName$, ByVal sSection$, ByVal sKey$, Optional ByVal vDefault As Variant = "")
Attribute GetSetting.VB_Description = "Returns a key setting value from an application's entry in the Windows registry"
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME: Function ggobjReg.gobjReg.GetSetting
'
'                     sAppName$          - A string expression containing the name of the application or project whose key setting is requested
'                     sSection$          - String expression containing the name of the section where the key setting is found.
'                     sKey$              - String expression containing the name of the key setting to return.
'                     vDefault           - Expression containing the value to return if no value is set in the key setting.
'
'                          ) As String
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Darren Adams    24 August 1999   First created for Pivotal Object Library
'
'                  PURPOSE: Returns a key setting value from an application's
'
'****************************************************************************
'
'
Dim sReturn$

'
' Other variables required by this function
'
Dim lReturnValue&, lRegistryKey&, lStringLength&, sKeyValue$

    '
    ' Initialise error vector
    '
    On Error Resume Next
    
    '
    ' Attempt to open the specific registry key.
    '
    lReturnValue = RegOpenKeyEx(REGISTRY_ROOT, REGISTRY_KEY_BASE & CompanyName & "\" & sAppName & IIf(sAppName = "", "", "\") & sSection, ByVal 0&, KEY_READ, lRegistryKey)
    
    '
    ' If the registry key exists, then attempt to retrieve the value, otherwise
    ' return the argument specified by the optional sDefault argument.
    '
    If lReturnValue = ERROR_SUCCESS Then
        
        '
        ' Set an initial length of the string variable that will be
        ' used to store the retrieved registry value
        '
        lStringLength = 32000
        sKeyValue = String$(lStringLength, vbNullChar)
        lReturnValue = RegQueryValueEx(lRegistryKey, sKey, 0&, REG_SZ, ByVal sKeyValue, lStringLength)
        
        If lReturnValue = ERROR_SUCCESS Then
            '
            ' Return the registry key value found (note: the variable "lStringLength" contains
            ' the lenght of the string plus the null-terminating character)
            '
            If lStringLength = 0 Then
                sReturn = CStr(vDefault)
            Else
                sKeyValue = Left$(sKeyValue, lStringLength - 1)
                sReturn = sKeyValue
            End If
        Else
            sReturn = CStr(vDefault)
        End If
    Else
        sReturn = CStr(vDefault)
    End If
    
    '
    ' Close the open registry key anyway
    '
    Call RegCloseKey(lRegistryKey)
    
    '
    ' Return value to caller
    '
    GetSetting = sReturn
    
End Function

Private Sub Class_Initialize()

    '
    ' Set the default name
    '
    CompanyName = DEF_COMPANY_NAME
    REGISTRY_ROOT = RegistryClasses.RegLocalMachine

End Sub
Public Function GetMimeType$(ByVal sType$)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME: Function GetMimeType
'
'                     sType$          - Filename of file extension to check
'
'                          ) As String
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    24 August 1999   First created for Newleaf Object Library
'
'                  PURPOSE: Returns the content type for the given file extension
'                           or if the type contains a dot then it extracts the
'                           extension portion
'
'****************************************************************************
'
'
Const CONTENT_TYPE = "Content Type"
Const DEFAULT_TYPE = "text/html"

Dim iCnt%
Dim sReturn$, sKeyValue$
Dim lReturnValue&, lRegistryKey&, lStringLength&



    '
    ' Determine the file extension
    '
    On Error Resume Next
    sReturn = DEFAULT_TYPE
    iCnt = InStrRev(sType, ".")
    If iCnt > 0 Then sType = Right(sType, Len(sType) - iCnt)
    sType = "." + Trim(sType)
    
    '
    ' Attempt to open the specific registry key.
    '
    lReturnValue = RegOpenKeyEx(RegistryClasses.RegRoot, sType, ByVal 0&, KEY_READ, lRegistryKey)
    
    '
    ' If the registry key exists, then attempt to retrieve the content type
    '
    If lReturnValue = ERROR_SUCCESS Then
        lStringLength = 4000
        sKeyValue = String$(lStringLength, vbNullChar)
        lReturnValue = RegQueryValueEx(lRegistryKey, CONTENT_TYPE, 0&, REG_SZ, ByVal sKeyValue, lStringLength)
        If lReturnValue = ERROR_SUCCESS Then
            
            '
            ' Return the registry key value found (note: the variable "lStringLength" contains
            ' the length of the string plus the null-terminating character)
            '
            If lStringLength > 0 Then sReturn = Left$(sKeyValue, lStringLength - 1)
        End If
    End If
    
    '
    ' Close the open registry key anyway
    '
    Call RegCloseKey(lRegistryKey)
    
    '
    ' Return value to caller
    '
    GetMimeType = sReturn
    
End Function



Public Function GetDocType$(ByVal sType$)
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME: Function GetDocType
'
'                     sType$          - Filename of file extension to check
'
'                          ) As String
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    24 August 1999   First created for Newleaf Object Library
'
'                  PURPOSE: Returns the document type for the given file extension
'                           or if the type contains a dot then it extracts the
'                           extension portion
'
'****************************************************************************
'
'
Dim iCnt%
Dim sReturn$, sKeyValue$
Dim lReturnValue&, lRegistryKey&, lStringLength&



    '
    ' Determine the file extension
    '
    On Error Resume Next
    iCnt = InStrRev(sType, ".")
    If iCnt > 0 Then sType = Right(sType, Len(sType) - iCnt)
    sReturn = sType
    sType = "." + Trim(sType)
    
    '
    ' Attempt to open the specific registry key.
    '
    lReturnValue = RegOpenKeyEx(RegistryClasses.RegRoot, sType, ByVal 0&, KEY_READ, lRegistryKey)
    
    '
    ' If the registry key exists, then attempt to retrieve the content type
    '
    If lReturnValue = ERROR_SUCCESS Then
        lStringLength = 4000
        sKeyValue = String$(lStringLength, vbNullChar)
        lReturnValue = RegQueryValueEx(lRegistryKey, "", 0&, REG_SZ, ByVal sKeyValue, lStringLength)
        If lReturnValue = ERROR_SUCCESS Then
            
            '
            ' Return the registry key value found (note: the variable "lStringLength" contains
            ' the length of the string plus the null-terminating character)
            '
            If lStringLength > 0 Then sReturn = Left$(sKeyValue, lStringLength - 1)
        End If
    End If
    
    '
    ' Close the open registry key anyway
    '
    Call RegCloseKey(lRegistryKey)
    
    '
    ' Return value to caller
    '
    GetDocType = Replace(sReturn, ".", " ")
    
End Function



Public Function GetSettingEx$(ByVal lKeyBase As RegistryClasses, ByVal sSection$, ByVal sKey$, Optional ByVal vDefault As Variant = "")
'****************************************************************************
'
'     Pivotal Solutions Ltd © 2000
'
'****************************************************************************
'
'                     NAME: Function GetSettingEx
'
'                     lKeyBase           - The ID of the main key
'                     sSection$          - String expression containing the name of the section where the key setting is found.
'                     sKey$              - String expression containing the name of the key setting to return.
'                     vDefault           - Expression containing the value to return if no value is set in the key setting.
'
'                          ) As String
'
'             DEPENDENCIES: NONE
'
'     MODIFICATION HISTORY: Steve O'Hara    19th November 2001   First created for Pivotal DDL Editor
'
'                  PURPOSE: Returns a key setting value from the key specified
'
'****************************************************************************
'
'
Dim sReturn$
Dim lReturnValue&, lRegistryKey&, lStringLength&, sKeyValue$

    '
    ' Attempt to open the specific registry key.
    '
    On Error Resume Next
    lReturnValue = RegOpenKeyEx(lKeyBase, sSection, ByVal 0&, KEY_READ, lRegistryKey)
    
    '
    ' If the registry key exists, then attempt to retrieve the value, otherwise
    ' return the argument specified by the optional sDefault argument.
    '
    If lReturnValue = ERROR_SUCCESS Then
        
        '
        ' Set an initial length of the string variable that will be
        ' used to store the retrieved registry value
        '
        lStringLength = 32000
        sKeyValue = String$(lStringLength, vbNullChar)
        lReturnValue = RegQueryValueEx(lRegistryKey, sKey, 0&, REG_SZ, ByVal sKeyValue, lStringLength)
        
        If lReturnValue = ERROR_SUCCESS Then
            '
            ' Return the registry key value found (note: the variable "lStringLength" contains
            ' the lenght of the string plus the null-terminating character)
            '
            If lStringLength = 0 Then
                sReturn = CStr(vDefault)
            Else
                sKeyValue = Left$(sKeyValue, lStringLength - 1)
                sReturn = sKeyValue
            End If
        Else
            sReturn = CStr(vDefault)
        End If
    Else
        sReturn = CStr(vDefault)
    End If
    
    '
    ' Close the open registry key anyway
    '
    Call RegCloseKey(lRegistryKey)
    
    '
    ' Return value to caller
    '
    GetSettingEx = sReturn
    
End Function

Private Function Z_IsWindowsXP() As Boolean
'********************************************************************************
'
'               Function :  Z_IsWindowsXP
'                           (
'                           )   -   Boolean :   The OS is XP
'
'   Modification History :  Mark Travis     25 February, 2002        Created
'
'            Description :  This function returns whether the current operating system
'                           is Windows XP
'
'********************************************************************************
'
Dim bReturn As Boolean: bReturn = False

    '
    ' Initialise the error vector
    '
    On Error Resume Next
    
    '
    ' Get the windows version
    '
    If (Not mbRetrievedVersion) Then _
        mbRetrievedVersion = Z_GetWindowsVersionInformation()
    
    '
    ' Check the returned values for XP flags
    '
    If (mbRetrievedVersion) Then _
        If (mstVersionInformation.dwPlatformId = pi_WindowsNT) Then _
            If (mstVersionInformation.dwMajorVersion = mjv_Windows2KXP) Then _
                bReturn = (mstVersionInformation.dwMinorVersion = mnv_WindowsXP)
    
    Z_IsWindowsXP = bReturn

End Function

Private Function Z_IsWindows2K() As Boolean
'********************************************************************************
'
'               Function :  Z_IsWindows2K
'                           (
'                           )   -   Boolean :   The OS is 2K
'
'   Modification History :  Mark Travis     25 February, 2002        Created
'
'            Description :  This function returns whether the current operating system
'                           is Windows 2K
'
'********************************************************************************
'
Dim bReturn As Boolean: bReturn = False

    '
    ' Initialise the error vector
    '
    On Error Resume Next
    
    '
    ' Get the windows version
    '
    If (Not mbRetrievedVersion) Then _
        mbRetrievedVersion = Z_GetWindowsVersionInformation()

    '
    ' Check the returned values for 2K flags
    '
    If (mbRetrievedVersion) Then _
        If (mstVersionInformation.dwPlatformId = pi_WindowsNT) Then _
            If (mstVersionInformation.dwMajorVersion = mjv_Windows2KXP) Then _
                bReturn = (mstVersionInformation.dwMinorVersion = mnv_Windows2K)
    
    Z_IsWindows2K = bReturn

End Function

Private Function Z_IsWindowsNT4x() As Boolean
'********************************************************************************
'
'               Function :  Z_IsWindowsNT4x
'                           (
'                           )   -   Boolean :   The OS is NT4x
'
'   Modification History :  Mark Travis     25 February, 2002        Created
'
'            Description :  This function returns whether the current operating system
'                           is Windows NT4x
'
'********************************************************************************
'
Dim bReturn As Boolean: bReturn = False

    '
    ' Initialise the error vector
    '
    On Error Resume Next
    
    '
    ' Get the windows version
    '
    If (Not mbRetrievedVersion) Then _
        mbRetrievedVersion = Z_GetWindowsVersionInformation()

    '
    ' Check the returned values for NT4x flags
    '
    If (mbRetrievedVersion) Then _
        If (mstVersionInformation.dwPlatformId = pi_WindowsNT) Then _
            bReturn = (mstVersionInformation.dwMajorVersion = mjv_Windows9xNT)
    
    Z_IsWindowsNT4x = bReturn

End Function

Private Function Z_IsWindows9x() As Boolean
'********************************************************************************
'
'               Function :  Z_IsWindows9x
'                           (
'                           )   -   Boolean :   The OS is 9x
'
'   Modification History :  Mark Travis     25 February, 2002        Created
'
'            Description :  This function returns whether the current operating system
'                           is Windows 9x
'
'********************************************************************************
'
Dim bReturn As Boolean: bReturn = False

    '
    ' Initialise the error vector
    '
    On Error Resume Next
    
    '
    ' Get the windows version
    '
    If (Not mbRetrievedVersion) Then _
        mbRetrievedVersion = Z_GetWindowsVersionInformation()
    
    '
    ' check the returned values for 9x flags
    '
    If (mbRetrievedVersion) Then _
        bReturn = (mstVersionInformation.dwPlatformId = pi_Windows9x)
    
    Z_IsWindows9x = bReturn

End Function

Private Function Z_GetWindowsVersionInformation() As Boolean
'********************************************************************************
'
'               Function :  Z_GetWindowsVersionInformation
'                           (
'                           )   -   Boolean :   Success
'
'   Modification History :  Mark Travis     25 February, 2002        Created
'
'            Description :  This function returns the operating system version
'
'********************************************************************************
'
Dim bReturn As Boolean: bReturn = False

    '
    ' Initialise the error vector
    '
    On Error Resume Next

    '
    ' Assign the size of the structure
    '
    mstVersionInformation.dwOSVersionInfoSize = Len(mstVersionInformation)

    '
    ' Get the version of the operating system
    '
    If (GetVersionEx(mstVersionInformation) <> 0) Then _
        bReturn = True
    
    Z_GetWindowsVersionInformation = bReturn

End Function


Public Function GetRegistryRoot&()
'********************************************************************************
'
'               Function :  GetRegistryRoot
'                           (
'                           )   -   Long :   Key ID of the registry root
'
'   Modification History :  Steve O'Hara     4 December, 2023        Created
'
'            Description :  This function returns the ID of the root key
'
'********************************************************************************
'
    GetRegistryRoot = REGISTRY_ROOT
End Function


Public Function GetRegistryKeyName$()
'********************************************************************************
'
'               Function :  GetRegistryKeyName
'                           (
'                           )   -   String :   Key name of the registry root
'
'   Modification History :  Steve O'Hara     4 December, 2023        Created
'
'            Description :  This function returns the name of the root key
'
'********************************************************************************
'

    If REGISTRY_ROOT = RegRoot Then
        GetRegistryKeyName = "HKEY_CLASSES_ROOT"
    ElseIf REGISTRY_ROOT = RegCurrentUser Then
        GetRegistryKeyName = "HKEY_CURRENT_USER"
    ElseIf REGISTRY_ROOT = RegLocalMachine Then
        GetRegistryKeyName = "HKEY_LOCAL_MACHINE"
    ElseIf REGISTRY_ROOT = RegPerformanceData Then
        GetRegistryKeyName = "HKEY_PERFORMANCE_DATA"
    ElseIf REGISTRY_ROOT = RegUsers Then
        GetRegistryKeyName = "HKEY_USERS"
    Else
        GetRegistryKeyName = ""
    End If

End Function

Public Function SaveToFile(ByVal sFilename$, ByVal sAppName$)
'********************************************************************************
'
'               Function :  SaveToFile
'                           (
'                               sFilename$  - File to save the settings to
'                               sAppName$   - Name of the application
'                           )   -   String :   Any errors from the save process
'
'   Modification History :  Steve O'Hara     4 December, 2023        Created
'
'            Description :  This function saves all the settings to the specified
'                           file and will return any errors
'
'********************************************************************************
'
Dim iFile%, iCnt%
Dim asSymbols As Variant

    '
    ' Attempt to ope the file for output
    '
    SaveToFile = ""
    On Error Resume Next
    iFile = FreeFile
    Open sFilename For Output As iFile
    If Err = 0 Then
    
        '
        ' Start with the easy stuff
        '
        Print #iFile, "Windows Registry Editor Version 5.00" & vbCrLf
        Z_WriteSectionToFile iFile, sAppName
        
        '
        ' Now the Last Good Rates
        '
        Z_WriteSectionToFile iFile, sAppName, REG_LAST_GOOD_RATES
        
        '
        ' Now the Last Good Values
        '
        Z_WriteSectionToFile iFile, sAppName, REG_LAST_GOOD_VALUES
    
        '
        ' Now the Settings
        '
        Z_WriteSectionToFile iFile, sAppName, REG_SETTINGS
    
        '
        ' Loop round all the symbols
        '
        asSymbols = GetSubkeys(App.Title, REG_SYMBOLS)
        If Not IsEmpty(asSymbols) Then
            asSymbols = PSGEN_SortArraySimple(asSymbols)
            For iCnt = 0 To UBound(asSymbols)
                Z_WriteSectionToFile iFile, sAppName, REG_SYMBOLS & "\" & asSymbols(iCnt)
            Next
        End If
    
        '
        ' Close the file
        '
        Close iFile
    Else
        SaveToFile = "Error trying to open file [" & sFilename & "] for output" & vbCrLf & Err.Description
        PSGEN_Log SaveToFile, LogError
    End If

End Function


Public Function LoadFromFile(ByVal sFilename$, ByVal sAppName$)
'********************************************************************************
'
'               Function :  LoadFromFile
'                           (
'                               sFilename$  - File to save the settings to
'                               sAppName$   - Name of the application
'                           )   -   String :   Any errors from the save process
'
'   Modification History :  Steve O'Hara     4 December, 2023        Created
'
'            Description :  This function loads all the settings from the specified
'                           file and will return any errors
'
'********************************************************************************
'
Dim iFile%
Dim sLine$, sSection$, sKey$, sValue$
Dim bGotFirstSection As Boolean

    '
    ' Delete the app settings first
    '
    LoadFromFile = ""
    On Error Resume Next
    DeleteSetting sAppName

    '
    ' Attempt to open the file for output
    '
    iFile = FreeFile
    Open sFilename For Input As iFile
    If Err = 0 Then
    
        Do While Not EOF(iFile)
            Line Input #iFile, sLine
            
            '
            ' Check for a section definition
            '
            If sLine Like "[[]*]" Then
                sSection = Split(Mid(sLine, 2, Len(sLine) - 2), sAppName + "\", 2)(1)
                
            '
            ' If not empty and we have a section, then must be a value
            '
            ElseIf sSection <> "" And sLine <> "" Then
                sKey = Replace(Split(sLine, "=", 2)(0), """", "")
                sValue = Replace(Replace(Replace(Split(sLine, "=", 2)(1), "\""", "~#~"), """", ""), "~#~", """")
                If sKey <> "" Then
                    SaveSetting sAppName, sSection, sKey, sValue
                End If
            End If
        Loop
        
        '
        ' Close the file
        '
        Close iFile
    Else
        LoadFromFile = "Error trying to open file [" & sFilename & "] for output" & vbCrLf & Err.Description
        PSGEN_Log LoadFromFile, LogError
    End If

End Function


Private Sub Z_WriteSectionToFile(ByVal iFile%, ByVal sAppName$, Optional ByVal sSection$ = "")
'********************************************************************************
'
'               Function :  Z_WriteSectionToFile
'                           (
'                               iFile%      - File to write to
'                               sAppName$   - Name of the application
'                               sSection$   - Name of th section to wite with all of its values
'                           )
'
'   Modification History :  Steve O'Hara     4 December, 2023        Created
'
'            Description :  This subroutine writes the section to the output
'                           file along with the registry values for it. It throws
'                           an error if something fails
'
'********************************************************************************
'

Dim asValues As Variant
Dim iCnt%

    Print #iFile, "[" & GetRegistryKeyName() & "\" & REGISTRY_KEY_BASE & CompanyName & "\" & sAppName & IIf(sSection = "", "", "\" & sSection) & "]"
    If sSection <> "" Then
        asValues = GetAllSettings(sAppName, sSection)
        If Not IsEmpty(asValues) Then
            For iCnt = 0 To UBound(asValues)
                If asValues(iCnt, 1) <> "" Then
                    Print #iFile, """" & asValues(iCnt, 0) & """=""" & Replace(asValues(iCnt, 1), """", "\""") & """"
                End If
            Next
        End If
    End If
    Print #iFile, ""

End Sub


