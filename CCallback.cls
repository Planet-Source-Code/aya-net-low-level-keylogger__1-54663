VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event KeyPressed(ByVal intKey As Integer, ByVal strKeyName As String, ByVal bIsFunctionKey As Boolean, ByVal strInApplication As String)

Public FireShift As Boolean
Public FireTab As Boolean
Public FireFKeys As Boolean
Public FireCapital As Boolean
Public FireArrowKeys As Boolean
Public FireSpecialKeys As Boolean
Public FireAlt As Boolean
Public FireControl As Boolean
Public FireAltGr As Boolean
Public FireScrollLock As Boolean
Public FireNumLock As Boolean
Public FireWindowsKeys As Boolean

Public Function FireEvent(ByVal intKey As Integer, ByVal strKeyName As String, ByVal bIsFunctionKey As Boolean, ByVal strInApplication As String)
   RaiseEvent KeyPressed(intKey, strKeyName, bIsFunctionKey, strInApplication)
End Function

