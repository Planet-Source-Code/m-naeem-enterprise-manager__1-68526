VERSION 5.00
Begin VB.Form frm_Vocabulary_Master 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3840
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8925
   Icon            =   "frm_Vocabulary_Master.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   8925
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pic_BMP_Menu 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   705
      Left            =   570
      ScaleHeight     =   645
      ScaleWidth      =   975
      TabIndex        =   15
      Top             =   3990
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Timer tmr_Slide_Show 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   6360
      Top             =   -240
   End
   Begin VB.TextBox txt_Display 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Height          =   495
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   3330
      Width           =   8775
   End
   Begin VB.Frame fra_Game 
      Enabled         =   0   'False
      Height          =   3135
      Left            =   3030
      TabIndex        =   9
      Top             =   120
      Width           =   2805
      Begin VB.TextBox txt_M_Guess 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   1815
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CommandButton cmd_Play_Meaning 
         Caption         =   "Meaning"
         Height          =   375
         Left            =   1320
         TabIndex        =   12
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txt_W_Guess 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   120
         Width           =   2595
      End
      Begin VB.CommandButton cmd_Play 
         Caption         =   "Play"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Timer tmr_VM 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   5040
      Top             =   -120
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   30
      TabIndex        =   6
      Top             =   120
      Width           =   2895
      Begin VB.TextBox txt_Word_Meaning 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   1920
         Width           =   2655
      End
      Begin VB.ListBox lst_Words 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   2655
      End
   End
   Begin VB.Frame fra_add_word 
      Height          =   3135
      Left            =   5940
      TabIndex        =   0
      Top             =   120
      Width           =   2925
      Begin VB.CommandButton cmd_Add_Word 
         Caption         =   "Add New Word"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   2520
         Width           =   2295
      End
      Begin VB.TextBox txt_Meaning 
         ForeColor       =   &H00C00000&
         Height          =   1215
         Left            =   120
         MaxLength       =   254
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txt_Word 
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         MaxLength       =   254
         TabIndex        =   1
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label lbl_add_vocabulary 
         AutoSize        =   -1  'True
         Caption         =   "Meaning"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lbl_add_vocabulary 
         AutoSize        =   -1  'True
         Caption         =   "Word"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   390
      End
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "&Windows"
      Begin VB.Menu mnu_Windows 
         Caption         =   "Address Book"
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
   Begin VB.Menu mnu_Slide_Show 
      Caption         =   "&Start Slide Show"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnu_Configure 
      Caption         =   "&Configure"
      Begin VB.Menu mnu_Meanings 
         Caption         =   "Meanings"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnu_Separator 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Slide_Show_Time 
         Caption         =   "1 Sec"
         Index           =   1
      End
      Begin VB.Menu mnu_Slide_Show_Time 
         Caption         =   "2 Sec"
         Index           =   2
      End
      Begin VB.Menu mnu_Slide_Show_Time 
         Caption         =   "3 Sec"
         Checked         =   -1  'True
         Index           =   3
      End
      Begin VB.Menu mnu_Slide_Show_Time 
         Caption         =   "4 Sec"
         Index           =   4
      End
      Begin VB.Menu mnu_Slide_Show_Time 
         Caption         =   "5 Sec"
         Index           =   5
      End
      Begin VB.Menu mnu_Slide_Show_Time 
         Caption         =   "6 Sec"
         Index           =   6
      End
      Begin VB.Menu mnu_Slide_Show_Time 
         Caption         =   "7 Sec"
         Index           =   7
      End
      Begin VB.Menu mnu_Slide_Show_Time 
         Caption         =   "8 Sec"
         Index           =   8
      End
      Begin VB.Menu mnu_Slide_Show_Time 
         Caption         =   "9 Sec"
         Index           =   9
      End
      Begin VB.Menu mnu_Slide_Show_Time 
         Caption         =   "10 Sec"
         Index           =   10
      End
      Begin VB.Menu mnu_Slide_Show_Time 
         Caption         =   "15 Sec"
         Index           =   15
      End
      Begin VB.Menu mnu_Slide_Show_Time 
         Caption         =   "20 Sec"
         Index           =   20
      End
      Begin VB.Menu mnu_Slide_Show_Time 
         Caption         =   "30 Sec"
         Index           =   30
      End
      Begin VB.Menu mnu_Slide_Show_Time 
         Caption         =   "45 Sec"
         Index           =   45
      End
      Begin VB.Menu mnu_Slide_Show_Time 
         Caption         =   "1 Min"
         Index           =   60
      End
      Begin VB.Menu mnu_sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Count 
         Caption         =   "Count"
      End
   End
   Begin VB.Menu mnu_expand 
      Caption         =   "&Expand"
   End
   Begin VB.Menu mnu_About 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frm_Vocabulary_Master"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'''''''''''''''''''''''''''''''
Dim Num As Long, old_Index As Byte, o_Index As Byte
Dim str_Word() As String, str_Meaning() As String
Dim last_Alphabet As Integer
Const offset = 3000

Dim lng_Lst_Index As Long

Private Sub cmd_add_word_Click()
 Dim R1 As New ADODB.Recordset
 Dim SQL As String
 ''''''''''''''''''''''''
 On Error GoTo EH
 
 If Trim(txt_Word.Text) = vbNullString Or Trim(txt_Meaning.Text) = vbNullString Then
  txt_Display.Text = "Either Word or its Meaning is missing"
  Exit Sub
 End If
 '''''''''''''''''''
 SQL = "select * from vocabulary order by word"
 R1.Open SQL, Con, adOpenDynamic, adLockPessimistic
 R1.AddNew
 R1.Fields(1).Value = Trim(txt_Word.Text)
 R1.Fields(2).Value = Trim(txt_Meaning.Text)
 R1.Update
 R1.Close
 
 txt_Display.Text = txt_Word.Text & " added "
 txt_Word.Text = vbNullString
 txt_Meaning.Text = vbNullString
 
 '''''''''''''''''
 Exit Sub
EH:
 Select Case Err.Number
  Case -2147217887: txt_Display.Text = "Failed:" & vbTab & "Duplicate Value..."
  Case Else: txt_Display.Text = Err.Number & vbCrLf & Err.Description
 End Select
End Sub

Private Sub cmd_Play_Click()
 Randomize Timer
 Num = Int(Rnd() * lst_Words.ListCount)
 txt_M_Guess.Text = vbNullString
 txt_W_Guess.Text = UCase(str_Word(Num))
End Sub

Private Sub cmd_Play_Meaning_Click()
 If Trim(txt_W_Guess.Text) = "" Then Exit Sub
 txt_M_Guess.Text = str_Meaning(Num)
End Sub

Private Sub Form_Activate()
 mnu_Load_Click (last_Alphabet)
End Sub

Private Sub Form_Load()
 On Error GoTo EH
 strMsg = vbNullString ''' must be at the top statement in load event
 Me.Width = Me.Width - offset
 txt_Display.Width = txt_Display.Width - offset
 strMsg = "Welcome to Vocabulary Master by" & EM_MAIL
 old_Index = 3
 Load_Menu Me, 1
 Exit Sub
EH:
 txt_Display.Text = Err.Description
End Sub

Private Sub lst_Words_Click()
 lng_Lst_Index = lst_Words.ListIndex
  If lng_Lst_Index = -1 Then lng_Lst_Index = 0
 txt_Word_Meaning.Text = str_Meaning(lng_Lst_Index)
End Sub

Private Sub lst_Words_DblClick()
 Dim F As Form
 Set F = frm_Vocabulary_Edit
 F.txt_Word.Text = lst_Words.Text
 F.txt_Meaning.Text = txt_Word_Meaning.Text
 F.Show vbModal
End Sub

Private Sub mnu_Count_Click()
 MsgBox "Total Vocabulary Items:- " & vbCrLf & vbCrLf & Space(7) & lst_Words.ListCount, vbInformation, EM_TITLE
End Sub

Private Sub mnu_Exit_Click()
 End
End Sub

Private Sub mnu_expand_Click()
 If mnu_expand.Caption = "&Expand" Then
   mnu_expand.Caption = "&Shrink"
   Me.Width = Me.Width + offset
   txt_Display.Width = txt_Display.Width + offset
 Else
   mnu_expand.Caption = "&Expand"
   Me.Width = Me.Width - offset
   txt_Display.Width = txt_Display.Width - offset
 End If
End Sub

Private Sub mnu_Meanings_Click()
  mnu_Meanings.Checked = Not mnu_Meanings.Checked
End Sub

Private Sub mnu_Load_Click(Index As Integer)
 Dim R1 As New ADODB.Recordset
 Dim Rec_Count As Long, i As Long
 On Error GoTo EH
 '''''''''''''
 lst_Words.Clear
 txt_Word_Meaning.Text = vbNullString
 last_Alphabet = Index
 If mnu_Load(Index).Caption = "All" Then
  strSql = "select * from vocabulary order by word"
 Else
  strSql = "select * from vocabulary where word like '" & mnu_Load(Index).Caption & "%' order by word"
 End If
 
''''''''''''''''''''''''''''''''''''''''''''''''
 mnu_Load(o_Index).Checked = False
 mnu_Load(Index).Checked = True
 o_Index = Index
'''''''''''''''''''''''''''''''''''''''''''''''
 R1.Open strSql, Con, 3, 3
 Rec_Count = R1.RecordCount
 ReDim str_Word(Rec_Count)
 ReDim str_Meaning(Rec_Count)
 
 While Not R1.EOF
   str_Word(i) = R1.Fields(1)
   str_Meaning(i) = R1.Fields(2)
   lst_Words.AddItem str_Word(i)
   R1.MoveNext
   i = i + 1
 Wend
 R1.Close
  
 fra_Game.Enabled = True
 mnu_Slide_Show.Enabled = True
 If lng_Lst_Index > 0 Then lst_Words.Text = lst_Words.List(lng_Lst_Index)
 txt_Display.Text = "Refreshed the Vocabulary Master...." & " Total Records = " & Rec_Count
 Exit Sub
EH:
 txt_Display.Text = Err.Description
End Sub

Private Sub mnu_Slide_Show_Click()
 If mnu_Slide_Show.Caption = "&Start Slide Show" Then
  mnu_Slide_Show.Caption = "&Stop Slide Show"
 Else
  mnu_Slide_Show.Caption = "&Start Slide Show"
 End If
 tmr_Slide_Show.Enabled = Not tmr_Slide_Show.Enabled
End Sub

Private Sub mnu_Slide_Show_Time_Click(Index As Integer)
 tmr_Slide_Show.Interval = Index * 1000
 mnu_Slide_Show_Time(Index).Checked = True
 mnu_Slide_Show_Time(old_Index).Checked = False
 old_Index = Index
End Sub

Private Sub tmr_Slide_Show_Timer()
 cmd_Play.Value = True
 If mnu_Meanings.Checked = True Then
  cmd_Play_Meaning.Value = True
 End If
End Sub

Private Sub tmr_VM_Timer()
 txt_Display.Text = vbNullString
 tmr_VM.Enabled = False
End Sub

Private Sub txt_Meaning_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
   cmd_Add_Word.Value = True
   txt_Word.SetFocus
 End If
End Sub

Private Sub txt_Word_Meaning_DblClick()
 txt_Word_Meaning.Text = vbNullString
End Sub

Private Sub mnu_About_Click()
 frm_About.Show vbModal
End Sub

Private Sub mnu_windows_Click(Index As Integer)
 Load_Selected_Window Me, Index
End Sub



