VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Progress_Bar 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   420
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar PB 
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   1
      Max             =   1
      Scrolling       =   1
   End
   Begin VB.Label lbl_Percentile 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5220
      TabIndex        =   1
      Top             =   30
      Width           =   390
   End
End
Attribute VB_Name = "frm_Progress_Bar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''''''''''''''''''''''''''''
Const MAX_VALUE As Long = 2000 ' higher the value , more the smoothness
Const DELAY As Long = 5 '' default value for dealy time...
Public Delay_Time As Long  '' delay time being set from caller
'''''''''''''''''''''''''''''

''' Progress Bar BackColor and ForeColor Setting API and declaration....''
''Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_USER = &H400
Private Const CCM_FIRST = &H2000
Private Const CCM_SETBKCOLOR = (CCM_FIRST + 1)
Private Const PBM_SETBARCOLOR = (WM_USER + 9)
Private Const SB_SETBKCOLOR = CCM_SETBKCOLOR

Private Sub SetBackColor(ProgressBarHwnd As Long, RGBValue As Long)
 Call SendMessage(ProgressBarHwnd, SB_SETBKCOLOR, 0, ByVal RGBValue)
End Sub
 
Private Sub SetBarColor(ProgressBarHwnd As Long, RGBValue As Long)
 Call SendMessage(ProgressBarHwnd, PBM_SETBARCOLOR, 0, ByVal RGBValue)
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



Private Sub Form_Activate()
 Make_Progress
End Sub


Private Sub Make_Progress()
 Dim C As Double
 Dim N As Long
 Dim Time_Stamp As Double
 
 '''''''''''''''''
 PB.Value = 0
 PB.Max = MAX_VALUE
 Randomize Timer
 PB.Scrolling = Rnd()
 
 '''''''''''''''''
 If Delay_Time = 0 Then Delay_Time = DELAY
 '''''''''''''''''
 Time_Stamp = Timer
 
 While ((Timer - Time_Stamp) < Delay_Time)
  C = Timer - Time_Stamp
  N = (MAX_VALUE / Delay_Time) * C
  If N > MAX_VALUE Then N = MAX_VALUE
  PB.Value = N
  lbl_Percentile.Caption = Format((N / MAX_VALUE) * 100, "00")
  DoEvents
 Wend
 '''''''''''''''''
 Delay_Time = 0
 Unload Me
End Sub


Private Sub Initialize_PB()
 Dim BC As Long, FC As Long
 '''''''''''''''''''''''''''''''
 PB.Value = 0
 Randomize Timer
 PB.Scrolling = Rnd
 Randomize Timer
 PB.BorderStyle = Rnd
 Randomize Timer
 PB.Appearance = Rnd
 Randomize Timer
 BC = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
 FC = vbWhite - BC
 SetBackColor PB.hWnd, BC
 SetBarColor PB.hWnd, FC
 ''''''''''''
 lbl_Percentile.BackColor = BC
 lbl_Percentile.ForeColor = FC
 lbl_Percentile.BorderStyle = PB.BorderStyle
End Sub

Private Sub Form_Load()
 Initialize_PB
End Sub
