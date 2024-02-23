VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'
' Copyright (c) 2024, Pivotal Solutions and/or its affiliates. All rights reserved.
' Pivotal Solutions PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
'
' Container that associates a downloaded chart with a timestamp
'
Option Explicit

    Public Timestamp As Date
    Public Chart As Picture
