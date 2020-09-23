Attribute VB_Name = "RModule"
Option Base 1
'declare functions, types and constants used for region handling


Type POINTAPI
        X As Long
        Y As Long
End Type

'a user-defined type for reading the region file
Type FileHeader
    CountsN As Integer
    VerticesN As Integer
    ID As Long
End Type

'a constant for identifying region files
Public Const fileID = 3941499

Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Declare Function CreatePolyPolygonRgn Lib "gdi32" (lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long

Public Const RGN_OR = 2
Public Const WINDING = 2
Public Const ALTERNATE = 1

'declare finctions for moving the window

Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
'declare other variables

Dim PoinTArray() As POINTAPI 'dynamic array where form region coordinates should be stored
Dim PointCounts() As Long 'dynamic array where number of points of each region is saved
Public Sub SetWRegion(PointFile As String, WindowName As Object)
'This procedure reads each point belonging to the region from a file specified,
'puts them into the array and sets the window region
'----------------------------

'declarations section
Dim PointCounter As Integer
Dim FileN As Integer
Dim Tempss As FileHeader
Dim RgnH As Long

'opening the points-file
FileN = FreeFile
Open PointFile For Binary Access Read As FileN
'reading file header
Get FileN, , Tempss
'checking if the file given is a true regions file.
'if not - exit the procedure without setting regions
If Tempss.ID <> fileID Then Exit Sub
ReDim PointCounts(Tempss.CountsN + 1)
ReDim PoinTArray(Tempss.VerticesN + 1)
Get FileN, , PointCounts
Get FileN, , PoinTArray

'closing the points-file
Close FileN


'Use the CreatePolyPolygonRgn function to create
'a region with all the points, stored in the array
RgnH = CreatePolyPolygonRgn(PoinTArray(2), PointCounts(1), UBound(PointCounts) - 1, ALTERNATE)

'set the region to the main window of the application
Call SetWindowRgn(WindowName.hwnd, RgnH, True)

End Sub
Public Sub CreateEllipse(ByRef TargetObj As Object, ByVal EllipseH As Long, ByVal EllipseW As Long)
'This procedure creates the rounded-corners form
'of each textbox. It needs a reference to the textboxes
'control array and the dimensions of the ellipse for rounded corners

Dim i As Integer
Dim RegionH As Long

For i = 0 To TargetObj.UBound
    RegionH = CreateRoundRectRgn(1, 1, TargetObj(i).Width, TargetObj(i).Height, EllipseH, EllipseW)
    Call SetWindowRgn(TargetObj(i).hwnd, RegionH, True)
Next
End Sub

Public Sub SetRegions()
'This procedure sets/creates the regions of the main form and of the controls

Dim AppPath As String

'check for points-file existence
If Not Right(App.Path, 1) = "\" Then AppPath = App.Path & "\" & "frame1.rgn" Else AppPath = App.Path & "frame1.rgn"
If Dir(AppPath) = "" Then MsgBox AppPath & " not found." & vbNewLine & "Please find and copy it to the executable file location.", vbCritical, "Error": End

'set the region to the form
Call SetWRegion(AppPath, frmPower)

'Set the regions of the buttons
Call CreateButton

'set the region for each textbox in the form
'You may add as many textboxes to the TextB array as you want
'and all of them will have rounded corners.
Call CreateEllipse(frmPower.TextB, 20, 20)

End Sub

Public Sub CreateButton()
'This sub creates the regions of the buttons

Dim RectRgn As Long, RoundRectRgn As Long, DestRgn As Long

'The way of creating the form of the buttons
'is a little bit tricky. The region is a combination of two
'independent regions which are combined using the CombineRgn function:
'a rounded-rectangular region and a rectangular region.
'If the rectangular region is narrower than the rounded-rectangular one
'and overlaps two of its rounded corners, we will have the wanted shape.

With frmPower

'Create the region of the Cancel button

'create the rounded rectangular region
RoundRectRgn = CreateRoundRectRgn(2, 2, .picCancel.Width - 2, .picCancel.Height - 2, 30, 30)

'create the rectangular region
RectRgn = CreateRectRgn(.picCancel.Width - 15, 2, .picCancel.Width - 3, .picCancel.Height - 3)

'the destination region (the combination between the two regions)
'must exist before CombineRgn is called so:
DestRgn = RoundRectRgn

'now combine the two regions using the API function
Call CombineRgn(DestRgn, RoundRectRgn, RectRgn, RGN_OR)

'set the region to the Cancel button
Call SetWindowRgn(.picCancel.hwnd, DestRgn, True)

'Create the region of the OK button in the same way
RoundRectRgn = CreateRoundRectRgn(2, 2, .picOK.Width - 2, .picOK.Height - 2, 30, 30)
RectRgn = CreateRectRgn(2, 2, 15, .picOK.Height - 3)
DestRgn = RoundRectRgn
Call CombineRgn(DestRgn, RoundRectRgn, RectRgn, RGN_OR)
Call SetWindowRgn(.picOK.hwnd, DestRgn, True)

End With

End Sub
