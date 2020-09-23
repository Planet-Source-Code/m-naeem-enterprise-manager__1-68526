Attribute VB_Name = "basDeclarations"
'   *********************************************************************
'   ******  Spy GUI Component!  Designed by naeem@email.com ***************
'   *********************************************************************
Option Explicit
Public Declare Function GetDC& Lib "user32" (ByVal hwnd As Long)
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

''Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public lngColorArray(7670000) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''   Blazing effect declarations       '''''''''''''''''''

Public Const Flame_Height = 30

'''''    Higher the number the shorter the flame   '''''''''
Type Pix
    r As Integer   ' Red
    g As Integer   ' Green
    B As Integer   ' Blue
    C As Boolean   ' Constant Colour
End Type

Public maxx As Integer   ' Array max x
Public maxy As Integer   ' Array max y

Public new_flame() As Pix  ' Flames buffers
Public old_flame() As Pix


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Public Declare Function ReleaseDC& Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long)


Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, _
    ByVal yPoint As Long) As Long
    
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" _
    (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
    (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''               some structures     ''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Type POINTAPI
  X As Long
  Y As Long
End Type

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''' Sound Object from Windows Media Player ...(msdxm.ocx)  ''''
Dim Sound As New MediaPlayer.MediaPlayer
Public Function PlaySound(ByVal strFile As String)
 On Error GoTo EH
  Sound.Open App.Path & "\" & strFile
  Exit Function
EH:
  Query.txtDisplay = Err.Number & Space(2) & Err.Description & vbCrLf & "PlaySound(strFile)"
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


