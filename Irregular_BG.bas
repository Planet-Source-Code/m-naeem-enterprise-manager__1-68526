Attribute VB_Name = "Irregular_BG"
Option Explicit
''''''''''''''''''''''''''''
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Const RGN_OR = 2

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''' Form Related Variables.............  '''''''''''''''''''''
Dim rgnBasic As New Region
Dim rgnExtended As New Region
Dim CurrentRgn As Long
Dim pic(0 To 1) As New StdPicture
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Public Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Public Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type

Public Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''' Form Related Events.......  '''''''''''''''''''''

Public Sub Load_Transparent_BG(F As Form)
    ' Load pictures from file
    'Set pic(0) = LoadPicture(App.Path & "\01.bmp", 0, 0, 0, 0)
    'Set pic(1) = LoadPicture(App.Path & "\02.bmp", 0, 0, 0, 0)
    Set pic(0) = F.imgLst_Transparent_BG.ListImages(1).picture
    Set pic(1) = F.imgLst_Transparent_BG.ListImages(1).picture
    
    ' Scan Shape from Green Screen Style Image
    Call rgnExtended.ScanPicture(pic(0))
    Call rgnBasic.ScanPicture(pic(1))
    ' Offset the Shape to allow for the form header.
    Call rgnBasic.OffsetHeader(F)
    Call rgnExtended.OffsetHeader(F)
    
    F.picture = pic(1) ' Set the Form Background
    Call rgnBasic.ApplyRgn(F.hWnd) ' Set the Form Shape
    CurrentRgn = rgnBasic.hndRegion ' Set the Current Shape
End Sub

Public Sub ExtendView(F As Form)
    If F.WindowState = vbMinimized Then Exit Sub
    If CurrentRgn <> rgnExtended.hndRegion Then ' If it is not already the Current Shape
        F.picture = pic(0) ' Set the Form Background
        Call rgnExtended.ApplyRgn(F.hWnd) ' Set the Form Shape
        CurrentRgn = rgnExtended.hndRegion ' Set the Current Shape
    End If
End Sub

Public Sub BasicView(F As Form)
  If F.WindowState = vbMinimized Then Exit Sub
    If CurrentRgn <> rgnBasic.hndRegion Then ' If it is not already the Current Shape
        F.picture = pic(1) ' Set the Form Background
        Call rgnBasic.ApplyRgn(F.hWnd) ' Set the Form Shape
        CurrentRgn = rgnBasic.hndRegion ' Set the Current Shape
    End If
End Sub

Public Sub Move_Form(F As Form, MouseButton As Integer)
    If MouseButton = vbLeftButton Then
        ReleaseCapture
        SendMessage F.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
End Sub

Public Sub Delete_Transparent_BG_Objects()
    Set rgnExtended = Nothing
    Set rgnBasic = Nothing
End Sub

'''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

