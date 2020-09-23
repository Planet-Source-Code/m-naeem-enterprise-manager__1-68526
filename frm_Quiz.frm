VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frm_Quiz 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4005
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7635
   Icon            =   "frm_Quiz.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   7635
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fra_Font 
      Height          =   585
      Left            =   6930
      TabIndex        =   11
      Top             =   3390
      Width           =   675
      Begin VB.CommandButton cmd_Font 
         Height          =   315
         Index           =   1
         Left            =   390
         Picture         =   "frm_Quiz.frx":382A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Decrease Font Size"
         Top             =   210
         Width           =   225
      End
      Begin VB.CommandButton cmd_Font 
         Height          =   315
         Index           =   0
         Left            =   60
         Picture         =   "frm_Quiz.frx":3974
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Increase Font Size"
         Top             =   150
         Width           =   225
      End
   End
   Begin VB.CommandButton cmd_Next 
      Caption         =   ">>"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3660
      TabIndex        =   10
      Top             =   3450
      Width           =   1275
   End
   Begin VB.CommandButton cmd_Answer 
      Caption         =   "Answer"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5190
      TabIndex        =   9
      Top             =   3450
      Width           =   1365
   End
   Begin VB.PictureBox pic_BMP_Menu 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   705
      Left            =   7830
      ScaleHeight     =   645
      ScaleWidth      =   975
      TabIndex        =   8
      Top             =   300
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txt_Display 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Height          =   525
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3450
      Width           =   3555
   End
   Begin VB.Frame fra_Quiz 
      Enabled         =   0   'False
      Height          =   3405
      Left            =   0
      TabIndex        =   0
      Top             =   -30
      Width           =   7635
      Begin VB.OptionButton opt_Guess 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Index           =   4
         Left            =   4020
         TabIndex        =   7
         Top             =   2370
         Width           =   3525
      End
      Begin VB.OptionButton opt_Guess 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Index           =   3
         Left            =   4020
         TabIndex        =   6
         Top             =   1260
         Width           =   3525
      End
      Begin VB.OptionButton opt_Guess 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Index           =   2
         Left            =   4020
         TabIndex        =   5
         Top             =   150
         Width           =   3525
      End
      Begin VB.OptionButton opt_Guess 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Index           =   1
         Left            =   60
         TabIndex        =   4
         Top             =   1530
         Width           =   3765
      End
      Begin VB.OptionButton opt_Guess 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   0
         Left            =   60
         TabIndex        =   1
         Top             =   150
         Width           =   3765
      End
      Begin VB.Label lbl_Word 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   2910
         Width           =   75
      End
   End
   Begin MSComDlg.CommonDialog cdl_File 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".htm"
      DialogTitle     =   "Convert Data"
      Filter          =   "HTML File|*.htm;*.html|Text File|*.txt|All Files|*.htm;*.html;*.txt"
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "&Windows"
      Begin VB.Menu mnu_Windows 
         Caption         =   "Graphics Menu"
         Index           =   0
      End
   End
   Begin VB.Menu mnu_Refresh 
      Caption         =   "&Refresh"
      Begin VB.Menu mnu_Load 
         Caption         =   "All"
         Index           =   0
      End
      Begin VB.Menu mnu_Load 
         Caption         =   "A"
         Index           =   1
      End
      Begin VB.Menu mnu_Load 
         Caption         =   "B"
         Index           =   2
      End
      Begin VB.Menu mnu_Load 
         Caption         =   "C"
         Index           =   3
      End
      Begin VB.Menu mnu_Load 
         Caption         =   "D"
         Index           =   4
      End
      Begin VB.Menu mnu_Load 
         Caption         =   "E"
         Index           =   5
      End
      Begin VB.Menu mnu_Load 
         Caption         =   "F"
         Index           =   6
      End
      Begin VB.Menu mnu_Load 
         Caption         =   "G"
         Index           =   7
      End
      Begin VB.Menu mnu_Load 
         Caption         =   "H"
         Index           =   8
      End
      Begin VB.Menu mnu_Load 
         Caption         =   "I"
         Index           =   9
      End
      Begin VB.Menu mnu_Load 
         Caption         =   "J"
         Index           =   10
      End
      Begin VB.Menu mnu_Load 
         Caption         =   "K"
         Index           =   11
      End
      Begin VB.Menu mnu_Load 
         Caption         =   "L"
         Index           =   12
      End
      Begin VB.Menu mnu_Load 
         Caption         =   "M"
         Index           =   13
      End
      Begin VB.Menu mnu_Load 
         Caption         =   "N"
         Index           =   14
      End
      Begin VB.Menu mnu_Load 
         Caption         =   "O"
         Index           =   15
      End
      Begin VB.Menu mnu_Load 
         Caption         =   "P"
         Index           =   16
      End
      Begin VB.Menu mnu_Load 
         Caption         =   "Q"
         Index           =   17
      End
      Begin VB.Menu mnu_Load 
         Caption         =   "R"
         Index           =   18
      End
      Begin VB.Menu mnu_Load 
         Caption         =   "S"
         Index           =   19
      End
      Begin VB.Menu mnu_Load 
         Caption         =   "T"
         Index           =   20
      End
      Begin VB.Menu mnu_Load 
         Caption         =   "U"
         Index           =   21
      End
      Begin VB.Menu mnu_Load 
         Caption         =   "V"
         Index           =   22
      End
      Begin VB.Menu mnu_Load 
         Caption         =   "W"
         Index           =   23
      End
      Begin VB.Menu mnu_Load 
         Caption         =   "X"
         Index           =   24
      End
      Begin VB.Menu mnu_Load 
         Caption         =   "Y"
         Index           =   25
      End
      Begin VB.Menu mnu_Load 
         Caption         =   "Z"
         Index           =   26
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "&Format"
      Begin VB.Menu mnu_Format 
         Caption         =   "Options Font"
         Index           =   0
      End
      Begin VB.Menu mnu_Format 
         Caption         =   "Options Font Color"
         Index           =   1
      End
      Begin VB.Menu mnu_Format 
         Caption         =   "Question Font"
         Index           =   2
      End
      Begin VB.Menu mnu_Format 
         Caption         =   "Question Font Color"
         Index           =   3
      End
      Begin VB.Menu mnu_Format 
         Caption         =   "Mistake Color"
         Index           =   4
      End
      Begin VB.Menu mnu_Format 
         Caption         =   "Correct Color"
         Index           =   5
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Save_Preference 
         Caption         =   "Save Preference"
      End
   End
   Begin VB.Menu mnu_Count 
      Caption         =   "&Count"
   End
   Begin VB.Menu mnu_Mistakes 
      Caption         =   "Mistakes"
   End
   Begin VB.Menu mnu_About 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frm_Quiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'''''''''''''''''''''''''''''''
Dim lng_Records As Long, lng_Correct As Long, lng_Wrong As Long
Dim lng_Word As Long
Dim lng_Answer As Long
Dim old_Index As Byte
Dim arr_Word_Quiz() As String
Dim arr_Meaning_Quiz() As String
Dim arr_Mst_Word_Quiz() As String
Dim arr_Mst_Meaning_Quiz() As String
Dim bol_Mistake As Boolean
Dim Mistake_Count As Integer
Dim col_Mistake As New Collection
''''''''''''''''''''''''''''''''''''''''''''''''
Dim Font_Name As String
Dim Font_Bold As String
Dim Font_Italic As String
Dim Font_Strikethru As String
Dim Font_Underline  As String
Dim Font_Size As Long
Dim Font_Color As OLE_COLOR

Dim Mistake_Color As OLE_COLOR
Dim Correct_Color As OLE_COLOR

Dim Q_Font_Name As String
Dim Q_Font_Bold As String
Dim Q_Font_Italic As String
Dim Q_Font_Underline As String
Dim Q_Font_Strikethru As String
Dim Q_Font_Size As Long
Dim Q_Font_Color As OLE_COLOR
''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub cmd_Answer_Click()
 If lng_Answer = lng_Word Then
  txt_Display.Text = "Correct"
  txt_Display.ForeColor = vbBlack
  opt_Guess(lng_Word).ForeColor = Correct_Color
  lng_Correct = lng_Correct + 1
 Else
  txt_Display.ForeColor = vbRed
  txt_Display.Text = "Wrong"
  opt_Guess(lng_Word).ForeColor = Mistake_Color
  lng_Wrong = lng_Wrong + 1
  If bol_Mistake = False Then Store_Mistakes
 End If
 fra_Quiz.Enabled = False
 cmd_Answer.Enabled = False
 mnu_Count.Caption = lng_Correct + lng_Wrong & " / " & lng_Records
End Sub

Private Sub Store_Mistakes()
On Error Resume Next
 ''''''''' make unique mistakes '''''''''''''''''''
  col_Mistake.Add lbl_Word.Caption, lbl_Word.Caption
  If Err.Number = 0 Then
    Mistake_Count = Mistake_Count + 1
    ReDim Preserve arr_Mst_Word_Quiz(Mistake_Count)
    ReDim Preserve arr_Mst_Meaning_Quiz(Mistake_Count)
    arr_Mst_Word_Quiz(Mistake_Count - 1) = lbl_Word.Caption
    arr_Mst_Meaning_Quiz(Mistake_Count - 1) = opt_Guess(lng_Word).Caption
  Else
     Err.Clear
  End If
End Sub

Private Sub Clear_Display()
Dim i As Integer
 For i = 0 To 4
  opt_Guess(i).Caption = vbNullString
  opt_Guess(i).Value = False
 Next
  lbl_Word.Caption = vbNullString
  cmd_Answer.Enabled = False
  txt_Display.ForeColor = vbBlack
End Sub

Private Sub cmd_Font_Click(Index As Integer)
 Dim FontSize As Integer
 Dim i As Integer
 '''''''''''''''''''''
 txt_Display.SetFocus
 '''''''''''''''''''''
 FontSize = opt_Guess(0).FontSize
 Select Case Index
  Case 0:
   If FontSize >= 16 Then Exit Sub
   FontSize = FontSize + 1
  Case 1:
   If FontSize <= 6 Then Exit Sub
   FontSize = FontSize - 1
  End Select
 '''''''''''''''
 For i = 0 To opt_Guess.Count - 1
  opt_Guess(i).FontSize = FontSize
 Next
 txt_Display.Text = "Font Size: " & Space(4) & FontSize
End Sub

Private Sub cmd_Font_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 cmd_Font(Index).BackColor = vbYellow
End Sub

Private Sub cmd_Font_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 cmd_Font(Index).BackColor = &H8000000F
End Sub

Private Sub cmd_Next_Click()
 Dim i As Integer
 Dim lng_Guess(5) As Long
 Dim lng_Dummy As Long
 Dim lngRnd As Long, colRnd As New Collection
'On Error GoTo EH
''''''''''''''''''''
  txt_Display.ForeColor = vbBlack
  txt_Display.Text = vbNullString
  opt_Guess(lng_Word).ForeColor = Font_Color
  opt_Guess(lng_Answer).Value = False
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 For i = 0 To 4
  Randomize Timer
  If bol_Mistake = True Then
   lngRnd = Int(Rnd() * UBound(arr_Mst_Word_Quiz))
  Else
   lngRnd = Int(Rnd() * UBound(arr_Word_Quiz))
  End If
     ''''''process to make unique vocabulary word meaning ''''
     On Error Resume Next
  colRnd.Add CStr(lngRnd), CStr(lngRnd)
  If Err.Number = 0 Then
     lng_Guess(i) = lngRnd
  Else
     i = i - 1
     Err.Clear
  End If
 Next
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  On Error GoTo EH
   Randomize Timer
   lng_Word = Int(Rnd() * 5)
 
 For i = 0 To 4
  If bol_Mistake = True Then
   opt_Guess(i).Caption = arr_Mst_Meaning_Quiz(lng_Guess(i))
  Else
   opt_Guess(i).Caption = arr_Meaning_Quiz(lng_Guess(i))
  End If
 Next
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If bol_Mistake = True Then
  lbl_Word.Caption = arr_Mst_Word_Quiz(lng_Guess(lng_Word))
 Else
  lbl_Word.Caption = arr_Word_Quiz(lng_Guess(lng_Word))
 End If
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 fra_Quiz.Enabled = True
 Exit Sub
EH:
 txt_Display.Text = Err.Description
End Sub

Private Sub Form_Load()
 strMsg = vbNullString ''' must be at the top statement in load event
 Retrieve_Preference arr_Windows(2)
 Load_Menu Me, 2
 Format_Skin
End Sub

Private Sub Form_Unload(Cancel As Integer)
  bol_Mistake = False
End Sub


Private Sub mnu_Count_Click()
 MsgBox "Total Vocabulary:-  " & lng_Records & vbCrLf & vbCrLf & "Attempted:-  " & lng_Correct + lng_Wrong & vbCrLf & vbCrLf & "Correct:-  " & lng_Correct & vbCrLf & vbCrLf & "Wrong:-  " & lng_Wrong, 0, EM_TITLE
End Sub


Private Sub mnu_Load_Click(Index As Integer)
 Dim i As Long
  On Error Resume Next
 
'''''''''''''''''''''''''''''''''''''''''''''''''
For i = 0 To mnu_Load.Count - 1
 mnu_Load(i).Checked = False
Next
 mnu_Load(Index).Checked = True
i = 0
'''''''''''''''''''''''''''''''''''''''''''''''
 
 If mnu_Load(Index).Caption = "All" Then
  strSql = "select * from vocabulary order by word"
 Else
  strSql = "select * from vocabulary where word like '" & mnu_Load(Index).Caption & "%' order by word"
 End If
 ''''''''''''''''
 Populate_Quiz Index
End Sub
Private Sub Populate_Quiz(IDX As Integer)
 Dim R1 As New ADODB.Recordset
 Dim i As Integer
 '''''''''''''''
 Erase arr_Word_Quiz
 Erase arr_Meaning_Quiz
   
 R1.Open strSql, Con, 3
 lng_Records = R1.RecordCount
 ReDim str_ID(lng_Records)
 ReDim arr_Word_Quiz(lng_Records)
 ReDim arr_Meaning_Quiz(lng_Records)
  
  While Not R1.EOF
   arr_Word_Quiz(i) = R1.Fields(1)
   arr_Meaning_Quiz(i) = R1.Fields(2)
   i = i + 1
   R1.MoveNext
  Wend
  R1.Close
  
  Clear_Display
  
  If lng_Records > 4 Then
   fra_Quiz.Enabled = True
   cmd_Next.Enabled = True
   strMsg = "Refreshed the Vocabulary Master"
  Else
   fra_Quiz.Enabled = False
   cmd_Next.Enabled = False
   strMsg = "No or Few Records found startting with '" & mnu_Load(IDX).Caption & "'"
  End If
   mnu_Count.Caption = "0 / " & lng_Records
   lng_Correct = 0: lng_Wrong = 0
  txt_Display.Text = Err.Number & Space(2) & strMsg
End Sub

Private Sub mnu_Mistakes_Click()
On Error GoTo EH
'''''''''''''''''''''
If UBound(arr_Mst_Word_Quiz) < 5 Then
 txt_Display.Text = "Only " & UBound(arr_Mst_Word_Quiz) & " Unique Mistake(s) found.." & vbCrLf & "Mistakes must be greate than Five...."
 Exit Sub
End If
''''''''''''''''''''
Clear_Display
''''''''''''''''''''
bol_Mistake = Not bol_Mistake
If cmd_Next.Caption = "Mistakes" Then
 cmd_Next.Caption = ">>"
 mnu_Mistakes.Caption = "Mistakes"
Else
 cmd_Next.Caption = "Mistakes"
 mnu_Mistakes.Caption = ">>"
 fra_Quiz.Enabled = True
 cmd_Next.Enabled = True
End If
Exit Sub
EH:
 'txt_Display.Text = Err.Description
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub mnu_Format_Click(Index As Integer)
 Dim i As Integer
Select Case Index
 
 Case 0
  cdl_File.Flags = cdlCFBoth
  cdl_File.ShowFont
  If Trim(cdl_File.FontName) = vbNullString Then Exit Sub
  Font_Name = cdl_File.FontName
  Font_Bold = cdl_File.FontBold
  Font_Italic = cdl_File.FontItalic
  Font_Size = cdl_File.FontSize
  Font_Strikethru = cdl_File.FontStrikethru
  Font_Underline = cdl_File.FontUnderline
  cdl_File.FontName = vbNullString
  Format_Skin
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
 Case 1:
  cdl_File.ShowColor
  If cdl_File.Color = 0 Then Exit Sub
  Font_Color = cdl_File.Color
  For i = 0 To 4
   opt_Guess(i).ForeColor = Font_Color
  Next
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
 Case 2:
  cdl_File.Flags = cdlCFBoth
  cdl_File.ShowFont
  If Trim(cdl_File.FontName) = vbNullString Then Exit Sub
  Q_Font_Name = cdl_File.FontName
  Q_Font_Bold = cdl_File.FontBold
  Q_Font_Italic = cdl_File.FontItalic
  Q_Font_Size = cdl_File.FontSize
  Q_Font_Strikethru = cdl_File.FontStrikethru
  Q_Font_Underline = cdl_File.FontUnderline
  cdl_File.FontName = vbNullString
  Format_Skin
  
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
 Case 3:
  cdl_File.ShowColor
  If cdl_File.Color = 0 Then Exit Sub
  Q_Font_Color = cdl_File.Color
  lbl_Word.ForeColor = Q_Font_Color
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
 Case 4:
  cdl_File.ShowColor
  If cdl_File.Color = 0 Then Exit Sub
  Mistake_Color = cdl_File.Color
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
 Case 5:
  cdl_File.ShowColor
  If cdl_File.Color = 0 Then Exit Sub
  Correct_Color = cdl_File.Color
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
End Select

End Sub

Private Sub Format_Skin()
 Dim i As Integer
 
 For i = 0 To 4
  opt_Guess(i).FontName = Font_Name
  opt_Guess(i).FontBold = Font_Bold
  opt_Guess(i).FontItalic = Font_Italic
  opt_Guess(i).FontUnderline = Font_Underline
  opt_Guess(i).FontStrikethru = Font_Strikethru
  opt_Guess(i).FontSize = Font_Size
  opt_Guess(i).ForeColor = Font_Color
 Next
  lbl_Word.FontName = Q_Font_Name
  lbl_Word.FontBold = Q_Font_Bold
  lbl_Word.FontItalic = Q_Font_Italic
  lbl_Word.FontUnderline = Q_Font_Underline
  lbl_Word.FontStrikethru = Q_Font_Strikethru
  lbl_Word.FontSize = Q_Font_Size
  lbl_Word.ForeColor = Q_Font_Color
End Sub

Private Sub Retrieve_Preference(arr_Win As String)
 Dim R1 As New ADODB.Recordset
 Dim lng_RecCount As Long
 Dim arr_Pref As Variant
 On Error GoTo EH
  ''''''''''''''''''''''''''''''''''
strSql = "Select * from preferences where em_category='" & arr_Win & "'"

R1.Open strSql, Con, adOpenDynamic, adLockOptimistic
If Not R1.EOF Then
  arr_Pref = Split(R1.Fields(2), DLMT)
  Font_Name = arr_Pref(0)
  Font_Bold = arr_Pref(1)
  Font_Italic = arr_Pref(2)
  Font_Underline = arr_Pref(3)
  Font_Strikethru = arr_Pref(4)
  Font_Size = arr_Pref(5)
  Font_Color = arr_Pref(6)
  Q_Font_Name = arr_Pref(7)
  Q_Font_Bold = arr_Pref(8)
  Q_Font_Italic = arr_Pref(9)
  Q_Font_Underline = arr_Pref(10)
  Q_Font_Strikethru = arr_Pref(11)
  Q_Font_Size = arr_Pref(12)
  Q_Font_Color = arr_Pref(13)
  Mistake_Color = arr_Pref(14)
  Correct_Color = arr_Pref(15)
  '' close the connection now...
  R1.Close
Else
   Font_Name = "Arial"
   Font_Bold = False
   Font_Italic = False
   Font_Underline = False
   Font_Strikethru = False
   Font_Size = 8
   Font_Color = vbBlack
   Q_Font_Name = "Arial"
   Q_Font_Bold = False
   Q_Font_Italic = False
   Q_Font_Underline = False
   Q_Font_Strikethru = False
   Q_Font_Size = 8
   Q_Font_Color = vbBlack
   Mistake_Color = vbMagenta
   Correct_Color = vbGreen
End If
 ''''''''''''''''''''''''''''''
 Exit Sub
EH:
 txt_Display.Text = Err.Description
End Sub

Private Sub mnu_Save_Preference_Click()
 Save_Prefrence arr_Windows(2)
End Sub

Private Sub Save_Prefrence(arr_Win As String)
 Dim R1 As New ADODB.Recordset
 Dim lng_RecCount As Long
 Dim str_Pref As String
 On Error GoTo EH
 ''''''''''''''''''''''''''
 str_Pref = Font_Name & DLMT & Font_Bold & DLMT & Font_Italic & DLMT & Font_Underline & DLMT & Font_Strikethru & DLMT & Font_Size & DLMT & Font_Color
 str_Pref = str_Pref & DLMT & Q_Font_Name & DLMT & Q_Font_Bold & DLMT & Q_Font_Italic & DLMT & Q_Font_Underline & DLMT & Q_Font_Strikethru & DLMT & Q_Font_Size & DLMT & Q_Font_Color
 str_Pref = str_Pref & DLMT & Mistake_Color & DLMT & Correct_Color
     
 strSql = "Select * from preferences where em_category='" & arr_Win & "'"
With R1
 .Open strSql, Con, adOpenDynamic, adLockOptimistic
 If .EOF Then .AddNew
 .Fields(1) = arr_Win
 .Fields(2) = str_Pref
 .Update
 .Close
End With
  txt_Display.Text = arr_Win & " Preference was saved successfully...."
 Exit Sub
EH:
 txt_Display.Text = Err.Description
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub mnu_About_Click()
 frm_About.Show vbModal
End Sub

Private Sub opt_Guess_Click(Index As Integer)
 lng_Answer = Index
 cmd_Answer.Enabled = True
End Sub

Private Sub txt_Display_DblClick()
 txt_Display.Text = vbNullString
End Sub

Private Sub mnu_windows_Click(Index As Integer)
 Load_Selected_Window Me, Index
End Sub

