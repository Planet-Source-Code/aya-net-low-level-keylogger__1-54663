Attribute VB_Name = "KeyboardDispatcher"
Option Explicit

Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Type KBDLLHOOKSTRUCT
   vkCode As Long
   scanCode As Long
   flags As Long
   time As Long
   dwExtraInfo As Long
End Type

Public Const WH_KEYBOARD_LL As Long = 13
Public Const HC_ACTION As Long = 0&

Public Const WM_KEYDOWN As Long = &H100
Public Const WM_KEYUP As Long = &H101
Public Const WM_SYSKEYDOWN As Long = &H104
Public Const WM_SYSKEYUP As Long = &H105

Public Const VK_ATTN As Long = &HF6
Public Const VK_ADD As Long = &H6B
Public Const VK_BACK As Long = &H8
Public Const VK_CANCEL As Long = &H3
Public Const VK_CAPITAL As Long = &H14
Public Const VK_CLEAR As Long = &HC
Public Const VK_CONTROL As Long = &H11
Public Const VK_CRSEL As Long = &HF7
Public Const VK_DECIMAL As Long = &H6E
Public Const VK_DELETE As Long = &H2E
Public Const VK_DIVIDE As Long = &H6F
Public Const VK_DOWN As Long = &H28
Public Const VK_END As Long = &H23
Public Const VK_EREOF As Long = &HF9
Public Const VK_ESCAPE As Long = &H1B
Public Const VK_EXECUTE As Long = &H2B
Public Const VK_EXSEL As Long = &HF8
Public Const VK_F1 As Long = &H70
Public Const VK_F10 As Long = &H79
Public Const VK_F11 As Long = &H7A
Public Const VK_F12 As Long = &H7B
Public Const VK_F13 As Long = &H7C
Public Const VK_F14 As Long = &H7D
Public Const VK_F15 As Long = &H7E
Public Const VK_F16 As Long = &H7F
Public Const VK_F17 As Long = &H80
Public Const VK_F18 As Long = &H81
Public Const VK_F19 As Long = &H82
Public Const VK_F2 As Long = &H71
Public Const VK_F20 As Long = &H83
Public Const VK_F21 As Long = &H84
Public Const VK_F22 As Long = &H85
Public Const VK_F23 As Long = &H86
Public Const VK_F24 As Long = &H87
Public Const VK_F3 As Long = &H72
Public Const VK_F4 As Long = &H73
Public Const VK_F5 As Long = &H74
Public Const VK_F6 As Long = &H75
Public Const VK_F7 As Long = &H76
Public Const VK_F8 As Long = &H77
Public Const VK_F9 As Long = &H78
Public Const VK_HELP As Long = &H2F
Public Const VK_HOME As Long = &H24
Public Const VK_INSERT As Long = &H2D
Public Const VK_LBUTTON As Long = &H1
Public Const VK_LCONTROL As Long = &HA2
Public Const VK_LEFT As Long = &H25
Public Const VK_LMENU As Long = &HA4
Public Const VK_LSHIFT As Long = &HA0
Public Const VK_MBUTTON As Long = &H4             '  NOT contiguous with L RBUTTON
Public Const VK_MENU As Long = &H12
Public Const VK_MULTIPLY As Long = &H6A
Public Const VK_NEXT As Long = &H22
Public Const VK_NONAME As Long = &HFC
Public Const VK_NUMLOCK As Long = &H90
Public Const VK_NUMPAD0 As Long = &H60
Public Const VK_NUMPAD1 As Long = &H61
Public Const VK_NUMPAD2 As Long = &H62
Public Const VK_NUMPAD3 As Long = &H63
Public Const VK_NUMPAD4 As Long = &H64
Public Const VK_NUMPAD5 As Long = &H65
Public Const VK_NUMPAD6 As Long = &H66
Public Const VK_NUMPAD7 As Long = &H67
Public Const VK_NUMPAD8 As Long = &H68
Public Const VK_NUMPAD9 As Long = &H69
Public Const VK_OEM_CLEAR As Long = &HFE
Public Const VK_PA1 As Long = &HFD
Public Const VK_PAUSE As Long = &H13
Public Const VK_PLAY As Long = &HFA
Public Const VK_PRINT As Long = &H2A
Public Const VK_PRIOR As Long = &H21
Public Const VK_RBUTTON As Long = &H2
Public Const VK_RCONTROL As Long = &HA3
Public Const VK_RETURN As Long = &HD
Public Const VK_RIGHT As Long = &H27
Public Const VK_RMENU As Long = &HA5
Public Const VK_RSHIFT As Long = &HA1
Public Const VK_SCROLL As Long = &H91
Public Const VK_SELECT As Long = &H29
Public Const VK_SEPARATOR As Long = &H6C
Public Const VK_SHIFT As Long = &H10
Public Const VK_SNAPSHOT As Long = &H2C
Public Const VK_SPACE As Long = &H20
Public Const VK_SUBTRACT As Long = &H6D
Public Const VK_TAB As Long = &H9
Public Const VK_UP As Long = &H26
Public Const VK_ZOOM As Long = &HFB

Public Const LLKHF_EXTENDED = &H1&    'test the extended-key flag
Public Const LLKHF_INJECTED = &H10&   'test the event-injected flag
Public Const LLKHF_ALTDOWN = &H20&    'test the context code
Public Const LLKHF_UP = &H80&         'test the transition-state flag

Public hHook As Long
Public bHooked As Boolean

Public EventHandler As CCallback

Public Sub DISPATCHER_START()
   hHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf KeyboardDispatcher.LowLevelKeyboardProc, App.hInstance, ByVal 0&)
   bHooked = True
End Sub

Public Sub DISPATCHER_STOPP()
   Dim res&
   res& = UnhookWindowsHookEx(hHook)
   bHooked = False
End Sub

Public Function LowLevelKeyboardProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   Dim msg As KBDLLHOOKSTRUCT
   Dim hwnd As Long
   Dim str$
   Dim f&
   Static blnShift As Boolean       ' Determines whether the shift key is hold
   Static blnCapital As Boolean     ' Determines whether the capital key is active.
   Static blnNumLock As Boolean     ' Determines whether the num lock key is active.
   Static blnScrollLock As Boolean  ' Determines whether the scroll lock key is active.
   Static blnAlt As Boolean         ' Determines whether the Alt key is hold.
   Static blnCtrl As Boolean        ' Determines whether the Ctrl key is hold.
   Static blnAltGr As Boolean       ' Determines whether the Alt+Gr key is hold.
   Dim bKeys(0 To 255) As Byte
   
   ' If nCode is less than zero, the hook procedure must pass the
   ' message to the CallNextHookEx function without further processing
   ' and should return the value returned by CallNextHookEx.
   If nCode < 0 Then
      LowLevelKeyboardProc = CallNextHookEx(hHook, nCode, wParam, lParam)
      Exit Function
   End If
   
   Call GetKeyboardState(bKeys(0))
      
   ' Retrieve the state of the combination keys such as
   ' num lock, lshift, rshift, lalt, ralt, lctrl, rctrl, capital, scroll lock
   
   If CBool(GetAsyncKeyState(VK_LSHIFT) And -32768) Or CBool(GetAsyncKeyState(VK_RSHIFT) And -32768) Then
      blnShift = True
   Else
      blnShift = False
   End If
   If CBool(GetAsyncKeyState(VK_RMENU) And -32768) Then
      blnAltGr = True
   Else
      blnAltGr = False
   End If
   If CBool(GetAsyncKeyState(VK_LMENU) And -32768) Then
      blnAlt = True
   Else
      blnAlt = False
   End If
   If CBool(GetAsyncKeyState(VK_LCONTROL) And -32768) Or CBool(GetAsyncKeyState(VK_RCONTROL) And -32768) Then
      blnCtrl = True
   Else
      blnCtrl = False
   End If
   If CBool(GetKeyState(VK_CAPITAL) And 1) Then
      blnCapital = True
   Else
      blnCapital = False
   End If
   If CBool(GetKeyState(VK_NUMLOCK) And 1) Then
      blnNumLock = True
   Else
      blnNumLock = False
   End If
   If CBool(GetKeyState(VK_SCROLL) And 1) Then
      blnScrollLock = True
   Else
      blnScrollLock = False
   End If
   
'   If CBool(bKeys(VK_RSHIFT) And 128) Or CBool(bKeys(VK_LSHIFT) And 128) Then
'      blnShift = True
'   Else
'      blnShift = False
'   End If
'   ' AltGr does not have the same meaning as Alt
'   If CBool(bKeys(VK_RMENU) And 128) Then
'      blnAltGr = True
'   Else
'      blnAltGr = False
'   End If
'   If CBool(bKeys(VK_LMENU) And 128) Then
'      blnAlt = True
'   Else
'      blnAlt = False
'   End If
'   ' Not important, which control key is pressed, the effect is the same.
'   If CBool(bKeys(VK_LCONTROL) And 128) Or CBool(bKeys(VK_RCONTROL) And 128) Then
'      blnCtrl = True
'   Else
'      blnCtrl = False
'   End If
'   If CBool(bKeys(VK_SCROLL) And 1) Then
'      blnScrollLock = True
'   Else
'      blnScrollLock = False
'   End If
'   ' Keys on the numeric keypad have a different behavior depending on the numlock key state.
'   If CBool(bKeys(VK_NUMLOCK) And 1) Then
'      blnNumLock = True
'   Else
'      blnNumLock = False
'   End If
'   ' Capital inverts the state of shift
'   If CBool(bKeys(VK_CAPITAL) And 1) Then
'      blnCapital = True
'   Else
'      blnCapital = False
'   End If
   
   ' Get window name for this keyboard event
   hwnd = GetForegroundWindow()
   str$ = Space(255)
   f& = GetWindowText(hwnd, ByVal str$, 255)
   str$ = Left(str$, f&)
   
   ' Retrieve event information
   Call CopyMemory(ByVal VarPtr(msg), ByVal lParam, LenB(msg))
   
   ' We just want to receive WM_KEYDOWN or WM_SYSKEYDOWN message, otherwise
   ' every key is catched 2 times.
   If wParam = WM_KEYUP Or wParam = WM_SYSKEYUP Then
      LowLevelKeyboardProc = CallNextHookEx(hHook, nCode, wParam, lParam)
      Exit Function
   End If
   
   Debug.Print "Capital: " & GetAsyncKeyState(VK_CAPITAL)
   Debug.Print "Capital: " & GetKeyState(VK_CAPITAL)
   
   ' Check for extended key, for example some special keys on logitech keyboards
   ' or a key on the numeric keypad.
   If CBool(msg.flags And LLKHF_EXTENDED) Then
      Select Case msg.vkCode
         Case VK_DELETE    ' ENTF
            Debug.Print "{Delete}"
            If EventHandler.FireSpecialKeys Then
               EventHandler.FireEvent vbKeyDelete, "{Delete}", True, str$
            End If
         Case VK_INSERT    ' INSERT
            Debug.Print "{Insert}"
            If EventHandler.FireSpecialKeys Then
               EventHandler.FireEvent vbKeyInsert, "{Insert}", True, str$
            End If
         Case VK_HOME      ' POS1
            Debug.Print "{Pos1}"
            If EventHandler.FireSpecialKeys Then
               EventHandler.FireEvent vbKeyHome, "{Home}", True, str$
            End If
         Case VK_PRIOR     ' PgUp
            Debug.Print "{PgUp}"
            If EventHandler.FireSpecialKeys Then
               EventHandler.FireEvent vbKeyPageUp, "{PgUp}", True, str$
            End If
         Case VK_END       ' End
            Debug.Print "{End}"
            If EventHandler.FireSpecialKeys Then
               EventHandler.FireEvent vbKeyEnd, "{End}", True, str$
            End If
         Case VK_NEXT      ' PgDown
            Debug.Print "{PgDown}"
            If EventHandler.FireSpecialKeys Then
               EventHandler.FireEvent vbKeyPageDown, "{PgDown}", True, str$
            End If
         Case VK_SNAPSHOT  ' Print Scrn
            Debug.Print "{Print Screen}"
            If EventHandler.FireSpecialKeys Then
               EventHandler.FireEvent vbKeySnapshot, "{Snapshot}", True, str$
            End If
         Case VK_PAUSE     ' Pause
            Debug.Print "{Pause}"
            If EventHandler.FireSpecialKeys Then
               EventHandler.FireEvent vbKeyPause, "{Pause}", True, str$
            End If
         Case VK_DIVIDE    ' Numpad Divide
            Debug.Print "/"
            ' More user friendly to send 47 instead of vbKeyDivide.
            ' vbKeyDivide cannot transformed into a char using 'Chr'
            EventHandler.FireEvent 47, "{Divide}", False, str$
         Case 13           ' Numpad Enter
            Debug.Print "{NPad.Enter}"
            EventHandler.FireEvent vbKeyReturn, "{Numpad.Return}", True, str$
         Case 92           ' Right Windows Key
            Debug.Print "{RightWin}"
            If EventHandler.FireWindowsKeys Then
               EventHandler.FireEvent -1, "{WinRight}", True, str$
            End If
         Case 91           ' Left Windows Key
            Debug.Print "{LeftWin}"
            If EventHandler.FireWindowsKeys Then
               EventHandler.FireEvent -1, "{WinLeft}", True, str$
            End If
         Case 93           ' Context menu
            Debug.Print "{ContextMenu}"
            If EventHandler.FireWindowsKeys Then
               EventHandler.FireEvent -1, "{ContextMenu}", True, str$
            End If
         Case 161          ' Right Shift
            Debug.Print "{RightShift}"
            If EventHandler.FireShift Then
               EventHandler.FireEvent vbKeyShift, "{RShift}", True, str$
            End If
         Case 37           ' Left arrow
            Debug.Print "{ArrowLeft}"
            If EventHandler.FireArrowKeys Then
               EventHandler.FireEvent vbKeyLeft, "{ArrowLeft}", True, str$
            End If
         Case 38           ' Up arrow
            Debug.Print "{ArrowUp}"
            If EventHandler.FireArrowKeys Then
               EventHandler.FireEvent vbKeyUp, "{ArrowUp}", True, str$
            End If
         Case 39           ' Right arrow
            Debug.Print "{ArrowRight}"
            If EventHandler.FireArrowKeys Then
               EventHandler.FireEvent vbKeyRight, "{ArrowRight}", True, str$
            End If
         Case 40           ' Arrow down
            Debug.Print "{ArrowDown}"
            If EventHandler.FireArrowKeys Then
               EventHandler.FireEvent vbKeyDown, "{ArrowDown}", True, str$
            End If
            
         ' AltGr seems to have 2 combinations:
         ' first is EXTENDED + 165
         ' second is ALTDOWN + 162
         ' Case 165          ' Alt+Gr
            ' Debug.Print "{AltGr}"
         Case 163
            Debug.Print "{RCtrl}"
            If EventHandler.FireControl Then
               EventHandler.FireEvent vbKeyControl, "{RCtrl}", True, str$
            End If
         Case 144    ' Numlock key
            Debug.Print "{NumLock}"
            If EventHandler.FireNumLock Then
               EventHandler.FireEvent vbKeyNumlock, "{NumLock}", True, str$
            End If
      End Select
   ElseIf msg.flags And LLKHF_ALTDOWN Then
      Select Case msg.vkCode
         Case 164       ' Alt
            Debug.Print "{Alt}"
            If EventHandler.FireAlt Then
               EventHandler.FireEvent -1, "{Alt}", True, str$
            End If
         Case 162       ' AltGr
            Debug.Print "{Alt+Gr}"
            If EventHandler.FireAltGr Then
               EventHandler.FireEvent vbKeyDelete, "{Alt+Gr}", True, str$
            End If
         Case Else
            ' Alt / Alt+Gr in combination with any other key
            GoTo defaultKeys
      End Select
   Else
defaultKeys:
      Select Case msg.vkCode
         Case 65 To 90       ' A to Z
            Select Case msg.vkCode
               Case 81
                  If blnAltGr Then
                     Debug.Print "@"
                     EventHandler.FireEvent Asc("@"), "{@}", False, str$
                  Else
                     GoTo normal_process
                  End If
               Case 77
                  If blnAltGr Then
                     Debug.Print "µ"
                     EventHandler.FireEvent Asc("µ"), "{µ}", False, str$
                  Else
                     GoTo normal_process
                  End If
               Case 69
                  If blnAltGr Then
                     Debug.Print "€"
                     EventHandler.FireEvent Asc("€"), "{€}", False, str$
                  Else
                     GoTo normal_process
                  End If
               Case Else
normal_process:
               If (blnCapital And Not blnShift) Or (Not blnCapital And blnShift) Then
                  Debug.Print UCase(Chr(msg.vkCode))
                  EventHandler.FireEvent Asc(UCase(Chr(msg.vkCode))), "{" & UCase(Chr(msg.vkCode)) & "}", False, str$
               Else
                  Debug.Print LCase(Chr(msg.vkCode))
                  EventHandler.FireEvent Asc(LCase(Chr(msg.vkCode))), "{" & LCase(Chr(msg.vkCode)) & "}", False, str$
               End If
            End Select
         Case 48     ' 0
            If blnAltGr Then
               Debug.Print "}"
               EventHandler.FireEvent Asc("}"), "{}}", False, str$
            Else
               If (blnCapital And Not blnShift) Or (Not blnCapital And blnShift) Then
                  Debug.Print "="
                  EventHandler.FireEvent Asc("="), "{=}", False, str$
               Else
                  Debug.Print "0"
                  EventHandler.FireEvent Asc("0"), "{0}", False, str$
               End If
            End If
         Case 49     ' 1
            If (blnCapital And Not blnShift) Or (Not blnCapital And blnShift) Then
               Debug.Print "!"
               EventHandler.FireEvent Asc("!"), "{!}", False, str$
            Else
               Debug.Print "1"
               EventHandler.FireEvent Asc("1"), "{1}", False, str$
            End If
         Case 50     ' 2
            If blnAltGr Then
               Debug.Print "²"
               EventHandler.FireEvent Asc("²"), "{²}", False, str$
            Else
               If (blnCapital And Not blnShift) Or (Not blnCapital And blnShift) Then
                  Debug.Print Chr(34)
                  EventHandler.FireEvent 34, "{" & Chr(34) & "}", False, str$
               Else
                  Debug.Print "2"
                  EventHandler.FireEvent Asc("2"), "{2}", False, str$
               End If
            End If
         Case 51     ' 3
            If blnAltGr Then
               Debug.Print "³"
               EventHandler.FireEvent Asc("³"), "{³}", False, str$
            Else
               If (blnCapital And Not blnShift) Or (Not blnCapital And blnShift) Then
                  Debug.Print "§"
                  EventHandler.FireEvent Asc("§"), "{§}", False, str$
               Else
                  Debug.Print "3"
                  EventHandler.FireEvent Asc("3"), "{3}", False, str$
               End If
            End If
         Case 52     ' 4
            If (blnCapital And Not blnShift) Or (Not blnCapital And blnShift) Then
               Debug.Print "$"
               EventHandler.FireEvent Asc("$"), "{$}", False, str$
            Else
               Debug.Print "4"
               EventHandler.FireEvent Asc("4"), "{4}", False, str$
            End If
         Case 53     ' 5
            If (blnCapital And Not blnShift) Or (Not blnCapital And blnShift) Then
               Debug.Print "%"
               EventHandler.FireEvent Asc("%"), "{%}", False, str$
            Else
               Debug.Print "5"
               EventHandler.FireEvent Asc("5"), "{5}", False, str$
            End If
         Case 54     ' 6
            If (blnCapital And Not blnShift) Or (Not blnCapital And blnShift) Then
               Debug.Print "&"
               EventHandler.FireEvent Asc("&"), "{&}", False, str$
            Else
               Debug.Print "6"
               EventHandler.FireEvent Asc("6"), "{6}", False, str$
            End If
         Case 55     ' 7
            If blnAltGr Then
               Debug.Print "{"
               EventHandler.FireEvent Asc("{"), "{{}", False, str$
            Else
               If (blnCapital And Not blnShift) Or (Not blnCapital And blnShift) Then
                  Debug.Print "/"
                  EventHandler.FireEvent Asc("/"), "{/}", False, str$
               Else
                  Debug.Print "7"
                  EventHandler.FireEvent Asc("7"), "{7}", False, str$
               End If
            End If
         Case 56     ' 8
            If blnAltGr Then
               Debug.Print "["
               EventHandler.FireEvent Asc("["), "{[}", False, str$
            Else
               If (blnCapital And Not blnShift) Or (Not blnCapital And blnShift) Then
                  Debug.Print "("
                  EventHandler.FireEvent Asc(")"), "{)}", False, str$
               Else
                  Debug.Print "8"
                  EventHandler.FireEvent Asc("8"), "{8}", False, str$
               End If
            End If
         Case 57     ' 9
            If blnAltGr Then
               Debug.Print "]"
               EventHandler.FireEvent Asc("]"), "{]}", False, str$
            Else
               If (blnCapital And Not blnShift) Or (Not blnCapital And blnShift) Then
                  Debug.Print ")"
                  EventHandler.FireEvent Asc(")"), "{)}", False, str$
               Else
                  Debug.Print "9"
                  EventHandler.FireEvent Asc("9"), "{9}", False, str$
               End If
            End If
         Case 220     ' ^
            If (blnCapital And Not blnShift) Or (Not blnCapital And blnShift) Then
               Debug.Print "°"
               EventHandler.FireEvent Asc("°"), "{°}", False, str$
            Else
               Debug.Print "^"
               EventHandler.FireEvent Asc("^"), "{^}", False, str$
            End If
         Case 219     ' ß
            If blnAltGr Then
               Debug.Print "\"
               EventHandler.FireEvent Asc("\"), "{\}", False, str$
            Else
               If (blnCapital And Not blnShift) Or (Not blnCapital And blnShift) Then
                  Debug.Print "?"
                  EventHandler.FireEvent Asc("?"), "{?}", False, str$
               Else
                  Debug.Print "ß"
                  EventHandler.FireEvent Asc("ß"), "{ß}", False, str$
               End If
            End If
         Case 187     ' +
            If blnAltGr Then
               Debug.Print "~"
               EventHandler.FireEvent Asc("~"), "{~}", False, str$
            Else
               If (blnCapital And Not blnShift) Or (Not blnCapital And blnShift) Then
                  Debug.Print "*"
                  EventHandler.FireEvent Asc("*"), "{*}", False, str$
               Else
                  Debug.Print "+"
                  EventHandler.FireEvent Asc("+"), "{+}", False, str$
               End If
            End If
         Case 191     ' #
            If (blnCapital And Not blnShift) Or (Not blnCapital And blnShift) Then
               Debug.Print "'"
               EventHandler.FireEvent Asc("'"), "{'}", False, str$
            Else
               Debug.Print "#"
               EventHandler.FireEvent Asc("#"), "{#}", False, str$
            End If
         Case 189     ' -
            If (blnCapital And Not blnShift) Or (Not blnCapital And blnShift) Then
               Debug.Print "_"
               EventHandler.FireEvent Asc("_"), "{_}", False, str$
            Else
               Debug.Print "-"
               EventHandler.FireEvent Asc("-"), "{-}", False, str$
            End If
         Case 190     ' .
            If (blnCapital And Not blnShift) Or (Not blnCapital And blnShift) Then
               Debug.Print ":"
               EventHandler.FireEvent Asc(":"), "{:}", False, str$
            Else
               Debug.Print "."
               EventHandler.FireEvent Asc("."), "{.}", False, str$
            End If
         Case 188     ' ,
            If (blnCapital And Not blnShift) Or (Not blnCapital And blnShift) Then
               Debug.Print ";"
               EventHandler.FireEvent Asc(";"), "{;}", False, str$
            Else
               Debug.Print ","
               EventHandler.FireEvent Asc(","), "{,}", False, str$
            End If
         Case 226     ' <
            If blnAltGr Then
               Debug.Print "|"
               EventHandler.FireEvent Asc("|"), "{|}", False, str$
            Else
               If (blnCapital And Not blnShift) Or (Not blnCapital And blnShift) Then
                  Debug.Print ">"
                  EventHandler.FireEvent Asc(">"), "{>}", False, str$
               Else
                  Debug.Print "<"
                  EventHandler.FireEvent Asc("<"), "{<}", False, str$
               End If
            End If
         Case 8       ' {BACKSPACE}
            Debug.Print "{BACKSPACE}"
            EventHandler.FireEvent vbKeyBack, "{Backspace}", False, str$
         Case 13      ' {ENTER}
            Debug.Print "{ENTER}"
            EventHandler.FireEvent vbKeyReturn, "{Return}", True, str$
         Case 221     ' ´
            If (blnCapital And Not blnShift) Or (Not blnCapital And blnShift) Then
               Debug.Print "`"
               EventHandler.FireEvent Asc("`"), "{`}", False, str$
            Else
               Debug.Print "´"
               EventHandler.FireEvent Asc("´"), "{´}", False, str$
            End If
         Case 106     ' Numpad Multiply
            Debug.Print "*"
            EventHandler.FireEvent Asc("*"), "{Multiply}", False, str$
         Case 109     ' Numpad Subtract
            Debug.Print "-"
            EventHandler.FireEvent Asc("-"), "{Subtract}", False, str$
         Case 107     ' Numpad Add
            Debug.Print "+"
            EventHandler.FireEvent Asc("+"), "{Add}", False, str$
         Case 110     ' Numpad ","
            Debug.Print ","
            EventHandler.FireEvent Asc(","), "{,}", False, str$
         Case 96      ' Numpad "0"
            Debug.Print "0"
            EventHandler.FireEvent Asc("0"), "{0}", False, str$
         Case 103     ' Numpad "7"
            Debug.Print "7"
            EventHandler.FireEvent Asc("7"), "{7}", False, str$
         Case 104     ' Numpad "8"
            Debug.Print "8"
            EventHandler.FireEvent Asc("8"), "{8}", False, str$
         Case 105     ' Numpad "9"
            Debug.Print "9"
            EventHandler.FireEvent Asc("9"), "{9}", False, str$
         Case 100     ' Numpad "4"
            Debug.Print "4"
            EventHandler.FireEvent Asc("4"), "{4}", False, str$
         Case 101     ' Numpad "5"
            Debug.Print "5"
            EventHandler.FireEvent Asc("5"), "{5}", False, str$
         Case 102     ' Numpad "6"
            Debug.Print "6"
            EventHandler.FireEvent Asc("6"), "{6}", False, str$
         Case 97      ' Numpad "1"
            Debug.Print "1"
            EventHandler.FireEvent Asc("1"), "{1}", False, str$
         Case 98      ' Numpad "2"
            Debug.Print "2"
            EventHandler.FireEvent Asc("2"), "{2}", False, str$
         Case 99      ' Numpad "3"
            Debug.Print "3"
            EventHandler.FireEvent Asc("3"), "{3}", False, str$
         Case 162     ' Left Ctrl
            Debug.Print "{LCtrl}"
            If EventHandler.FireControl Then
               EventHandler.FireEvent vbKeyControl, "{LCtrl}", True, str$
            End If
         Case 9       ' Tabulator
            Debug.Print "{Tab}"
            If EventHandler.FireTab Then
               EventHandler.FireEvent vbKeyTab, "{Tab}", True, str$
            End If
         Case 27
            Debug.Print "{Esc}"
            If EventHandler.FireSpecialKeys Then
               EventHandler.FireEvent vbKeyEscape, "{Escape}", True, str$
            End If
         Case VK_F1
            Debug.Print "{F1}"
            If EventHandler.FireFKeys Then
               EventHandler.FireEvent vbKeyF1, "{F1}", True, str$
            End If
         Case VK_F2
            Debug.Print "{F2}"
            If EventHandler.FireFKeys Then
               EventHandler.FireEvent vbKeyF2, "{F2}", True, str$
            End If
         Case VK_F3
            Debug.Print "{F3}"
            If EventHandler.FireFKeys Then
               EventHandler.FireEvent vbKeyF3, "{F3}", True, str$
            End If
         Case VK_F4
            Debug.Print "{F4}"
            If EventHandler.FireFKeys Then
               EventHandler.FireEvent vbKeyF4, "{F4}", True, str$
            End If
         Case VK_F5
            Debug.Print "{F5}"
            If EventHandler.FireFKeys Then
               EventHandler.FireEvent vbKeyF5, "{F5}", True, str$
            End If
         Case VK_F6
            Debug.Print "{F6}"
            If EventHandler.FireFKeys Then
               EventHandler.FireEvent vbKeyF6, "{F6}", True, str$
            End If
         Case VK_F7
            Debug.Print "{F7}"
            If EventHandler.FireFKeys Then
               EventHandler.FireEvent vbKeyF7, "{F7}", True, str$
            End If
         Case VK_F8
            Debug.Print "{F8}"
            If EventHandler.FireFKeys Then
               EventHandler.FireEvent vbKeyF8, "{F8}", True, str$
            End If
         Case VK_F9
            Debug.Print "{F9}"
            If EventHandler.FireFKeys Then
               EventHandler.FireEvent vbKeyF9, "{F9}", True, str$
            End If
         Case VK_F10
            Debug.Print "{F10}"
            If EventHandler.FireFKeys Then
               EventHandler.FireEvent vbKeyF10, "{F10}", True, str$
            End If
         Case VK_F11
            Debug.Print "{F11}"
            If EventHandler.FireFKeys Then
               EventHandler.FireEvent vbKeyF11, "{F11}", True, str$
            End If
         Case VK_F12
            Debug.Print "{F12}"
            If EventHandler.FireFKeys Then
               EventHandler.FireEvent vbKeyF12, "{F12}", True, str$
            End If
         Case VK_LSHIFT
            Debug.Print "{LShift}"
            If EventHandler.FireShift Then
               EventHandler.FireEvent vbKeyShift, "{LShift}", True, str$
            End If
         Case VK_CAPITAL
            Debug.Print "{Capital}"
            If EventHandler.FireCapital Then
               EventHandler.FireEvent vbKeyCapital, "{Capital}", True, str$
            End If
         Case 36     ' Numpad Pos1 (7)
            Debug.Print "{Pos1}"
            If EventHandler.FireSpecialKeys Then
               EventHandler.FireEvent vbKeyHome, "{Home}", True, str$
            End If
         Case 38     ' Numpad Arrow Up (8)
            Debug.Print "{ArrowUp}"
            If EventHandler.FireArrowKeys Then
               EventHandler.FireEvent vbKeyUp, "{ArrowUp}", True, str$
            End If
         Case 33     ' Numpad Page Down (9)
            Debug.Print "{PgUp}"
            If EventHandler.FireSpecialKeys Then
               EventHandler.FireEvent vbKeyPageUp, "{PgUp}", True, str$
            End If
         Case 37     ' Numpad Arrow Left (4)
            Debug.Print "{ArrowLeft}"
            If EventHandler.FireArrowKeys Then
               EventHandler.FireEvent vbKeyLeft, "{ArrowLeft}", True, str$
            End If
         Case 12     ' Numpad key 5 with numlock, no idea what's it's sense
            Debug.Print "{Numpad.5}"
            If EventHandler.FireSpecialKeys Then
               EventHandler.FireEvent -1, "{Numpad.5}", True, str$
            End If
         Case 39     ' Numpad Arrow Right (6)
            Debug.Print "{ArrowRight}"
            If EventHandler.FireArrowKeys Then
               EventHandler.FireEvent vbKeyRight, "{ArrowRight}", True, str$
            End If
         Case 35     ' Numpad End (1)
            Debug.Print "{End}"
            If EventHandler.FireSpecialKeys Then
               EventHandler.FireEvent vbKeyEnd, "{End}", True, str$
            End If
         Case 40     ' Numpad Arrow Down (2)
            Debug.Print "{ArrowDown}"
            If EventHandler.FireArrowKeys Then
               EventHandler.FireEvent vbKeyDown, "{ArrowDown}", True, str$
            End If
         Case 34     ' Numpad Page Down (3)
            Debug.Print "{PgDown}"
            If EventHandler.FireSpecialKeys Then
               EventHandler.FireEvent vbKeyPageDown, "{PgDown}", True, str$
            End If
         Case 45     ' Numpad Insert (0)
            Debug.Print "{Insert}"
            If EventHandler.FireSpecialKeys Then
               EventHandler.FireEvent vbKeyInsert, "{Insert}", True, str$
            End If
         Case 46     ' Numpad delete (,)
            Debug.Print "{Delete}"
            If EventHandler.FireSpecialKeys Then
               EventHandler.FireEvent vbKeyClear, "{Delete}", True, str$
            End If
         Case 32     ' Spacebar
            Debug.Print "{Space}"
            EventHandler.FireEvent 32, "{Space}", False, str$
         Case 192          ' ö
            If (blnCapital And Not blnShift) Or (Not blnCapital And blnShift) Then
               Debug.Print "Ö"
               EventHandler.FireEvent Asc("Ö"), "{Ö}", False, str$
            Else
               Debug.Print "ö"
               EventHandler.FireEvent Asc("ö"), "{ö}", False, str$
            End If
         Case 222          ' ä
            If (blnCapital And Not blnShift) Or (Not blnCapital And blnShift) Then
               Debug.Print "Ä"
               EventHandler.FireEvent Asc("Ä"), "{Ä}", False, str$
            Else
               Debug.Print "ä"
               EventHandler.FireEvent Asc("ä"), "{ä}", False, str$
            End If
         Case 186
            If (blnCapital And Not blnShift) Or (Not blnCapital And blnShift) Then
               Debug.Print "Ü"
               EventHandler.FireEvent Asc("Ü"), "{Ü}", False, str$
            Else
               Debug.Print "ü"
               EventHandler.FireEvent Asc("ü"), "{ü}", False, str$
            End If
         Case VK_SCROLL    ' Scroll
            Debug.Print "{Scroll}"
            If EventHandler.FireScrollLock Then
               EventHandler.FireEvent vbKeyScrollLock, "{ScrollLock}", True, str$
            End If
         Case Else   ' Unknown virtual keycode
            Debug.Print msg.vkCode
            EventHandler.FireEvent -1, "{Unknown: " & msg.vkCode & "}", True, str$
      End Select
   End If
   
   LowLevelKeyboardProc = CallNextHookEx(hHook, nCode, wParam, lParam)
End Function
