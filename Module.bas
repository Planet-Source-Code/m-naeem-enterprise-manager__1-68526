Attribute VB_Name = "Module"
Option Explicit
'''''''''''''''''''''''''''
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_NORMAL = 1
'''''''''''''''''''''''''''''''''''''''
Public Con As New ADODB.Connection
Public strSql As String
Public strCon As String
Public strMsg As String
Public Const EM_TITLE As String = "Enterprise Manager"
Public Const EM_MAIL = " (naeem@email.com)"
''''''''''page setup ...................
Public Top_Margin As String
Public Bottom_Margin As String
Public Left_Margin As String
Public Right_Margin As String
Public Table_Border As Tristate
Public Paragraph_Alignment As Long
'''''''''''''''''''''''''''''''''''''''
Public AB_SQL As String
Public Action_Mode As String
Public Edit_ID As Long
''''''''''''''''''''''''''''''''''''''''''''
Public Const DLMT As String = "-N-"
Public arr_Windows(5) As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''   Blazing effect declarations       '''''''''''''''''''

Public Const Flame_Height = 30

'''''    Higher the number the shorter the flame   '''''''''
Type Pix
 R As Integer   ' Red
 G As Integer   ' Green
 B As Integer   ' Blue
 C As Boolean   ' Constant Colour
End Type

Public maxx As Integer   ' Array max x
Public maxy As Integer   ' Array max y

Public new_flame() As Pix  ' Flames buffers
Public old_flame() As Pix
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Sub Main()
 Dim Pth As String, str_Reg As String
 On Error GoTo EH
 ''''''''''''''''''''''''''''
 arr_Windows(0) = "Address Book"
 arr_Windows(1) = "Vocabulary Master"
 arr_Windows(2) = "Quiz Master"
 arr_Windows(3) = "Schedular"
 arr_Windows(4) = "Loan Calculator"
 arr_Windows(5) = "Exit"
 '''''''''''''''''''''''''''''''''''''''
 '''''''''''''''''''''''''''''''''''''''
 ''Con.Open "DSN=EM"    ''' if DNS is used...
 Pth = App.Path & "\EM.mdb"
 
 'strCon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Pth
 strCon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Pth & ";Jet OLEDB:Database Password=naeems"
 Con.Open strCon
 
 
 str_Reg = Get_Registry()
  
 Select Case str_Reg
  Case arr_Windows(0): frm_Address_Book.Show
  Case arr_Windows(1): frm_Vocabulary_Master.Show
  Case arr_Windows(2): frm_Quiz.Show
  Case arr_Windows(3): frm_Schedular.Show
  Case arr_Windows(4): frm_Loan_Calculator.Show
  Case Else
     frm_Address_Book.Show '' testing purpose only...
 End Select
 
 Exit Sub
EH:
 strMsg = "Some Error in initializing the Database..." & vbCrLf & "Possible reasions are..." & vbCrLf & "Database file is missing , moved , corrupted or renamed.." & vbCrLf & "Database Driver is corrupt or uninstalled... " & vbCrLf & vbCrLf & "Check the Database file..."
 MsgBox Err.Description, 0, EM_TITLE & EM_MAIL
End Sub

Public Sub Save_Registry(last_Window As String)
 Dim R1 As New ADODB.Recordset
 R1.Open "select * from reg_setting", Con, adOpenDynamic, adLockOptimistic
 If R1.EOF = True Then R1.AddNew
  R1.Fields(0) = last_Window
  R1.Update
  R1.Close
End Sub

Private Function Get_Registry() As String
 Dim R1 As New ADODB.Recordset
 Dim Ret As String
 R1.Open "select * from reg_setting", Con
 If R1.EOF = True Then
  Ret = arr_Windows(1)
 Else
  Ret = R1.Fields(0)
 End If
 Get_Registry = Ret
End Function

Public Sub MoveForm(FF As Form, XX, YY, Bt)
 Static oldX, oldY, mF
 Dim moveLeft, moveTop
    ''''''''''''''''
 moveLeft = FF.Left + XX - oldX
 moveTop = FF.Top + YY - oldY
 If Bt = vbLeftButton Then
  If mF = 0 Then
   FF.Move moveLeft, moveTop
   FF.Refresh
   mF = 1
  Else
   mF = 0
  End If
 End If
  oldX = XX
  oldY = YY
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Load_Menu(F As Form, Menu_ID As Long)
 LoadMenu F, arr_Windows, Menu_ID
End Sub



Public Sub Load_Selected_Window(F As Form, Selected As Integer)
 Select Case F.mnu_Windows(Selected).Caption
  Case arr_Windows(0): frm_Address_Book.Show  '' Address Book
  Case arr_Windows(1): frm_Vocabulary_Master.Show   '' Vocabulary Master
  Case arr_Windows(2): frm_Quiz.Show  '' Quiz Master
  Case arr_Windows(3): frm_Schedular.Show   '' Schedular
  Case arr_Windows(4): frm_Loan_Calculator.Show '' Loan Calculator
  Case arr_Windows(5): End   '' Exit
 End Select
  Unload F
End Sub


Public Sub LoadMenu(F As Form, arr_Win() As String, CHK As Long)
 Dim i As Long
     ''''''''''''''
 F.mnu_Windows(0).Caption = arr_Win(0)
 For i = 1 To UBound(arr_Win)
  Load F.mnu_Windows(i)
  F.mnu_Windows(i).Caption = arr_Win(i)
  F.mnu_Windows(i).Visible = True
 Next
  F.mnu_Windows(CHK).Enabled = False
  F.mnu_Windows(CHK).Checked = True
  Save_Registry arr_Windows(CHK)
  F.Caption = EM_TITLE & ".." & arr_Windows(CHK) & EM_MAIL
End Sub

