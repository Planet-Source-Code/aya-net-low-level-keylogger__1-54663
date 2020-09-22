VERSION 5.00
Object = "*\AKeyboardLogger.vbp"
Begin VB.Form FMain 
   Caption         =   "Form1"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows-Standard
   Begin KeyboardLogger.keybd_log32 keybd_log321 
      Left            =   6120
      Top             =   3720
      _ExtentX        =   1270
      _ExtentY        =   1270
   End
   Begin VB.CheckBox Check12 
      Caption         =   "Enable F keys callback"
      Height          =   255
      Left            =   2850
      TabIndex        =   12
      Top             =   5430
      Value           =   1  'Aktiviert
      Width           =   2325
   End
   Begin VB.CheckBox Check11 
      Caption         =   "Enable scroll lock key callback"
      Height          =   255
      Left            =   2850
      TabIndex        =   11
      Top             =   5160
      Value           =   1  'Aktiviert
      Width           =   2985
   End
   Begin VB.CheckBox Check10 
      Caption         =   "Enable num lock key callback"
      Height          =   255
      Left            =   2850
      TabIndex        =   10
      Top             =   4890
      Value           =   1  'Aktiviert
      Width           =   3105
   End
   Begin VB.CheckBox Check9 
      Caption         =   "Enable capital key callback"
      Height          =   255
      Left            =   2850
      TabIndex        =   9
      Top             =   4620
      Value           =   1  'Aktiviert
      Width           =   2325
   End
   Begin VB.CheckBox Check8 
      Caption         =   "Enable special keys callback"
      Height          =   255
      Left            =   60
      TabIndex        =   8
      Top             =   6510
      Value           =   1  'Aktiviert
      Width           =   2985
   End
   Begin VB.CheckBox Check7 
      Caption         =   "Enable control keys callback"
      Height          =   255
      Left            =   60
      TabIndex        =   7
      Top             =   6240
      Value           =   1  'Aktiviert
      Width           =   2775
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Enable AltGr key callback"
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   5970
      Value           =   1  'Aktiviert
      Width           =   2325
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Enable Alt key callback"
      Height          =   255
      Left            =   60
      TabIndex        =   5
      Top             =   5700
      Value           =   1  'Aktiviert
      Width           =   2325
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Enable windows keys callback"
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   5430
      Value           =   1  'Aktiviert
      Width           =   2925
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Enable arrow keys callback"
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   5160
      Value           =   1  'Aktiviert
      Width           =   2325
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Enable tab callback"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   4890
      Value           =   1  'Aktiviert
      Width           =   2325
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Enable shift callback"
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   4620
      Value           =   1  'Aktiviert
      Width           =   2325
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4515
      Left            =   30
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   60
      Width           =   6885
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
   If Check1.Value = 1 Then
      Me.keybd_log321.FireShift = True
   Else
      Me.keybd_log321.FireShift = False
   End If
End Sub

Private Sub Check10_Click()
   If Check10.Value = 1 Then
      Me.keybd_log321.firenumlock = True
   Else
      Me.keybd_log321.firenumlock = False
   End If
End Sub

Private Sub Check11_Click()
   If Check11.Value = 1 Then
      Me.keybd_log321.firescrolllock = True
   Else
      Me.keybd_log321.firescrolllock = False
   End If
End Sub

Private Sub Check12_Click()
   If Check12.Value = 1 Then
      Me.keybd_log321.firefkeys = True
   Else
      Me.keybd_log321.firefkeys = False
   End If
End Sub

Private Sub Check2_Click()
   If Check2.Value = 1 Then
      Me.keybd_log321.FireTab = True
   Else
      Me.keybd_log321.FireTab = False
   End If
End Sub

Private Sub Check3_Click()
   If Check3.Value = 1 Then
      Me.keybd_log321.FireArrowkeys = True
   Else
      Me.keybd_log321.FireArrowkeys = False
   End If
End Sub

Private Sub Check4_Click()
   If Check4.Value = 1 Then
      Me.keybd_log321.Firewindowskeys = True
   Else
      Me.keybd_log321.Firewindowskeys = False
   End If
End Sub

Private Sub Check5_Click()
   If Check5.Value = 1 Then
      Me.keybd_log321.firealt = True
   Else
      Me.keybd_log321.firealt = False
   End If
End Sub

Private Sub Check6_Click()
   If Check6.Value = 1 Then
      Me.keybd_log321.firealtgr = True
   Else
      Me.keybd_log321.firealtgr = False
   End If
End Sub

Private Sub Check7_Click()
   If Check7.Value = 1 Then
      Me.keybd_log321.firecontrol = True
   Else
      Me.keybd_log321.firecontrol = False
   End If
End Sub

Private Sub Check8_Click()
   If Check8.Value = 1 Then
      Me.keybd_log321.firespecialkeys = True
   Else
      Me.keybd_log321.firespecialkeys = False
   End If
End Sub

Private Sub Check9_Click()
   If Check9.Value = 1 Then
      Me.keybd_log321.firecapital = True
   Else
      Me.keybd_log321.firecapital = False
   End If
End Sub

Private Sub Form_Load()
   Me.keybd_log321.EnableLogging
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Me.keybd_log321.DisabledLogging
End Sub

Private Sub keybd_log321_KeyPressed(ByVal intKey As Integer, ByVal strKeyName As String, ByVal bIsFunctionKey As Boolean, ByVal strInApplication As String, ByVal bDiffersFromLastApp As Boolean)
   If bDiffersFromLastApp Then
      If Not Me.txt.Text = "" Then
         Me.txt.Text = Me.txt.Text & vbNewLine & vbNewLine
      End If
      Me.txt.Text = Me.txt.Text & "[ " & strInApplication & " ]" & vbNewLine
      If bIsFunctionKey Then
         Me.txt.Text = Me.txt.Text & strKeyName
      Else
         Me.txt.Text = Me.txt.Text & Chr(intKey)
      End If
   Else
      If bIsFunctionKey Then
         Me.txt.Text = Me.txt.Text & strKeyName
      Else
         Me.txt.Text = Me.txt.Text & Chr(intKey)
      End If
   End If
End Sub
