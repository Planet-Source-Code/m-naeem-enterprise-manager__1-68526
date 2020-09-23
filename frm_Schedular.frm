VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_Schedular 
   ClientHeight    =   6465
   ClientLeft      =   165
   ClientTop       =   105
   ClientWidth     =   9750
   Icon            =   "frm_Schedular.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "frm_Schedular.frx":4DF2
   ScaleHeight     =   6465
   ScaleWidth      =   9750
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chk_Font 
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   8910
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Font Strikethru"
      Top             =   4650
      Width           =   195
   End
   Begin VB.CommandButton cmd_Font 
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   8910
      TabIndex        =   21
      ToolTipText     =   "Set Font"
      Top             =   5070
      Width           =   255
   End
   Begin VB.CommandButton cmd_BG_Color 
      Caption         =   "BG"
      Height          =   225
      Left            =   8490
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Set Back Ground Color"
      Top             =   5070
      Width           =   345
   End
   Begin VB.CheckBox chk_Font 
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Font Underline"
      Top             =   4650
      Width           =   195
   End
   Begin VB.CheckBox chk_Font 
      Caption         =   "I"
      Height          =   225
      Index           =   1
      Left            =   8370
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Font Italic"
      Top             =   4650
      Width           =   195
   End
   Begin VB.CheckBox chk_Font 
      Caption         =   "B"
      Height          =   225
      Index           =   0
      Left            =   8100
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Font Bold"
      Top             =   4650
      Width           =   195
   End
   Begin VB.CommandButton cmd_Font_Color 
      Caption         =   "FC"
      Height          =   225
      Left            =   8070
      TabIndex        =   16
      ToolTipText     =   "Set Font Color"
      Top             =   5070
      Width           =   345
   End
   Begin VB.CommandButton cmd_Font 
      DownPicture     =   "frm_Schedular.frx":D2494
      Height          =   405
      Index           =   1
      Left            =   8700
      Picture         =   "frm_Schedular.frx":D271B
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Decrease Font Size"
      Top             =   4020
      Width           =   345
   End
   Begin VB.CommandButton cmd_Font 
      DownPicture     =   "frm_Schedular.frx":D29AD
      Height          =   400
      Index           =   0
      Left            =   8100
      Picture         =   "frm_Schedular.frx":D2C33
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Increase Font Size"
      Top             =   4020
      Width           =   375
   End
   Begin VB.CommandButton cmd_Schedule 
      Caption         =   "Register Schedule"
      Height          =   315
      Left            =   1710
      TabIndex        =   12
      ToolTipText     =   "Register Schedule"
      Top             =   5910
      Width           =   1515
   End
   Begin VB.ComboBox cbo_Priority 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   2460
      Style           =   2  'Dropdown List
      TabIndex        =   10
      ToolTipText     =   "Priority"
      Top             =   4590
      Width           =   765
   End
   Begin VB.TextBox txt_Task 
      BackColor       =   &H00C0FFFF&
      Height          =   2205
      Left            =   3360
      MaxLength       =   490
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      ToolTipText     =   "Text Area..."
      Top             =   4020
      Width           =   4575
   End
   Begin VB.TextBox txt_Display 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   525
      Left            =   870
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      ToolTipText     =   "Prompt Display"
      Top             =   5100
      Width           =   2355
   End
   Begin MSComctlLib.ListView lvw_Task 
      Height          =   2415
      Left            =   840
      TabIndex        =   0
      Top             =   1500
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   4260
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12648447
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Task"
         Object.Width           =   8855
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Priority"
         Object.Width           =   1129
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
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
   Begin MSComctlLib.ImageList imgLst_Transparent_BG 
      Left            =   6960
      Top             =   -150
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   650
      ImageHeight     =   431
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Schedular.frx":D2ECF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtp_Schedule 
      Height          =   345
      Left            =   480
      TabIndex        =   11
      ToolTipText     =   "Set Date for Task"
      Top             =   4020
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   609
      _Version        =   393216
      CalendarForeColor=   -2147483647
      CustomFormat    =   "dd-mm-yyyy : h:m:s"
      Format          =   22675456
      CurrentDate     =   45283
   End
   Begin VB.Label lbl_dum 
      BackStyle       =   0  'Transparent
      Caption         =   "Priority:-"
      Height          =   195
      Index           =   4
      Left            =   1740
      TabIndex        =   25
      Top             =   4650
      Width           =   555
   End
   Begin VB.Label lbl_Exit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   1
      Left            =   8550
      TabIndex        =   24
      ToolTipText     =   "Minimize Window"
      Top             =   1020
      Width           =   120
   End
   Begin VB.Label lbl_Exit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   8790
      TabIndex        =   23
      ToolTipText     =   "Exit Window"
      Top             =   390
      Width           =   135
   End
   Begin VB.Image img_Save_Preference 
      Height          =   240
      Left            =   8130
      Picture         =   "frm_Schedular.frx":1A0581
      ToolTipText     =   "Save Preference Scheme"
      Top             =   5460
      Width           =   240
   End
   Begin VB.Label lbl_Task 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   510
      Left            =   5220
      TabIndex        =   13
      ToolTipText     =   "Number of Pending Tasks"
      Top             =   870
      Width           =   120
   End
   Begin VB.Label lbl_Menus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Word Page"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Index           =   6
      Left            =   1830
      TabIndex        =   9
      Top             =   990
      Width           =   480
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl_Menus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Convert"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   5
      Left            =   990
      TabIndex        =   8
      Top             =   1020
      Width           =   630
   End
   Begin VB.Label lbl_Menus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "History"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   4
      Left            =   2460
      TabIndex        =   7
      Top             =   900
      Width           =   585
   End
   Begin VB.Label lbl_Menus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sort"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   3
      Left            =   2550
      TabIndex        =   6
      Top             =   570
      Width           =   345
   End
   Begin VB.Label lbl_Menus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   1740
      TabIndex        =   5
      Top             =   270
      Width           =   615
   End
   Begin VB.Label lbl_Menus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   1
      Left            =   600
      TabIndex        =   4
      Top             =   750
      Width           =   525
   End
   Begin VB.Label lbl_Menus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Windows"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   600
      TabIndex        =   3
      Top             =   420
      Width           =   705
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "&Windows"
      Visible         =   0   'False
      Begin VB.Menu mnu_Windows 
         Caption         =   "Windows"
         Index           =   0
      End
   End
   Begin VB.Menu mnuDelete 
      Caption         =   "&Delete"
      Visible         =   0   'False
      Begin VB.Menu mnu_Delete 
         Caption         =   "All"
         Index           =   0
      End
      Begin VB.Menu mnu_Delete 
         Caption         =   "Individual"
         Index           =   1
      End
   End
   Begin VB.Menu mnu_Refresh 
      Caption         =   "&Refresh"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuSort 
      Caption         =   "&Sort"
      Visible         =   0   'False
      Begin VB.Menu mnu_Sort 
         Caption         =   "Ascending"
         Checked         =   -1  'True
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnu_Sort 
         Caption         =   "Descending"
         Index           =   1
      End
      Begin VB.Menu mnu_Sep_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Sort_By 
         Caption         =   "by Priority"
         Checked         =   -1  'True
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnu_Sort_By 
         Caption         =   "by Date"
         Index           =   1
      End
      Begin VB.Menu mnu_Sort_By 
         Caption         =   "by Task"
         Index           =   2
      End
   End
   Begin VB.Menu mnu_History 
      Caption         =   "History"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuConvert 
      Caption         =   "&Convert"
      Visible         =   0   'False
      Begin VB.Menu mnu_Convert 
         Caption         =   "Word"
         Index           =   0
      End
      Begin VB.Menu mnu_Convert 
         Caption         =   "WEB (HTML)"
         Index           =   1
      End
   End
   Begin VB.Menu mnu_Word_Page_Setup 
      Caption         =   "&Word Page Setup"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frm_Schedular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''''''''''''''''''''''''''''
Dim str_Date As String
Const Chk_BG_Color As Long = &HD0E0F0
Const UnChk_BG_Color As Long = &H8000000F
Dim i As Integer
Const dt_Format As String = "h:m:s :: dd-mmm-yyyy"
'''''''''''''''''''''''''''''''''''''''''
Dim Task_ID As Long
Dim Task_Name As String
Dim Task_Date As String
Dim Task_Priority As Long
'''''''''''''''''''''''''''''''''''''''''
Dim Tble As String
'''''''''''''''''''''''''''''''''''''''''
 Dim Font_Name As String
 Dim Font_Bold As String
 Dim Font_Italic As String
 Dim Font_Strikethru As String
 Dim Font_Underline  As String
 Dim Font_Size As Long
 Dim Font_Color As OLE_COLOR
 Dim Back_Color As OLE_COLOR
''''''''''''''''''''''''''''''''''''''''''
Dim Asc_Desc As String
Dim Sort_By As String



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Move_Form Me, Button
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Delete_Transparent_BG_Objects
End Sub


Private Sub Form_Load()
 strMsg = vbNullString ''' must be at the top statement in load event
 Load_Transparent_BG Me
 Load_Menu Me, 3
 Tble = "schedule"
 dtp_Schedule.Value = Date
 Load_Priority
 Asc_Desc = "asc"
 Sort_By = "priority"
 Populate_Task Tble
 Retrieve_Preference arr_Windows(3)
 Format_Skin
 Font_Size = txt_Task.FontSize
End Sub

Private Sub Load_Priority()
 Dim J As Long
 For J = 1 To 100
  cbo_Priority.AddItem J
 Next
End Sub

Private Sub cmd_Schedule_Click()
 Dim R1 As New ADODB.Recordset
 Dim Slot_Exist As Boolean
 Dim Diff As Long
 Dim Answer As VbMsgBoxResult
 '''''''''''''''''''''''''''''''''''''
 On Error GoTo EH
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''
 txt_Task.Text = Trim(txt_Task.Text)
 ''''''''''''''''''''''''''''''''''''
 If txt_Task.Text = vbNullString Then
  txt_Display.Text = "No Task ,Order or Schedule mentioned...?"
  Exit Sub
 End If
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If cbo_Priority.Text = vbNullString Then
  txt_Display.Text = "PRIORITY level Missing.....!"
  Exit Sub
 End If
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''
 Answer = MsgBox("Do you want to proceed....", vbInformation + vbYesNo, EM_TITLE)
 If Answer <> vbYes Then Exit Sub
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
 strSql = "select priority from schedule where priority >=" & Val(cbo_Priority.Text)
 With R1
  .Open strSql, Con, adOpenDynamic, adLockPessimistic
 While Not .EOF
   .Fields("priority") = (Val(.Fields("priority")) + 1)
 .Update
 .MoveNext
 Wend
 .Close
 End With
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''
 strSql = "select * from schedule"
 With R1
  .Open strSql, Con, adOpenDynamic, adLockPessimistic
  .AddNew
   .Fields("task") = Trim(txt_Task.Text)
   .Fields("date_time") = dtp_Schedule.Value + Time()
   .Fields("priority") = Val(cbo_Priority.Text)
 .Update
 .Close
 End With
 '''''''''''''''''''''''''''''
 Populate_Task Tble
 str_Date = vbNullString
 txt_Display.Text = Trim(txt_Task.Text) & " has been entered "
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 Exit Sub
EH:
 txt_Display.Text = Err.Description
End Sub


Private Sub lbl_Exit_Click(Index As Integer)
 Select Case Index
  Case 0: End
  Case 1: Me.WindowState = vbMinimized
 End Select
End Sub

Private Sub lbl_Menus_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 Select Case Index
  Case 0: Me.PopupMenu mnuWindows
  Case 1: Me.PopupMenu mnuDelete
  Case 2: mnu_Refresh_Click
  Case 3: Me.PopupMenu mnuSort
  Case 4: mnu_History_Click
  Case 5: mnu_Convert_Click (1) ''Me.PopupMenu mnuConvert
  Case 6: mnu_Word_Page_Setup_Click
 End Select
End Sub

Private Sub lvw_Task_BeforeLabelEdit(Cancel As Integer)
 Cancel = True
End Sub

Private Sub mnu_Convert_Click(Index As Integer)
 Dim File_Name As String
 ''''''''''''''''''''
 On Error GoTo EH
 If lvw_Task.ListItems.Count = 0 Then Exit Sub
 '''''''''''''''''''
 cdl_File.DialogTitle = EM_TITLE & " Save As"
 
 Select Case Index
 Case 0:
  cdl_File.Filter = "Word File|*.doc"
  cdl_File.DefaultExt = "doc"
  cdl_File.ShowSave
  File_Name = cdl_File.fileName
  If File_Name = vbNullString Then Exit Sub
  Write_Word_DOC File_Name
  
  Case 1:
  cdl_File.Filter = "Web Page File|*.htm"
  cdl_File.DefaultExt = "htm"
  cdl_File.ShowSave
  File_Name = cdl_File.fileName
  If File_Name = vbNullString Then Exit Sub
  Write_Web_Page File_Name
  
 End Select
 ''''''''''''''''''''''
 File_Name = vbNullString
 Exit Sub
EH:
 txt_Display.Text = Err.Description
End Sub

Private Sub mnu_Delete_Click(Index As Integer)
 Dim R1 As New ADODB.Recordset
 Dim Answer As VbMsgBoxResult, strQuestion As String
 Dim S As String
 Dim A() As String
 Dim V As Variant
 Dim J As Long
On Error GoTo EH
'''''''''''''''''''''''''
 If lvw_Task.ListItems.Count = 0 Then Exit Sub
'''''''''''''''''''''''''
Select Case Index
 Case 0:
  strSql = "delete from schedule"
  strQuestion = "Are you sure to ERASE All Task "
  strMsg = "All Tasks were DELETED successfully"
 Case 1:
  If Task_Name = vbNullString Then Exit Sub
  strSql = "delete from schedule where id=" & Task_ID
  strQuestion = "Are you sure to ERASE Task " & vbCrLf & vbCrLf & Task_Name & vbCrLf & Task_Date
  strMsg = Task_Name & vbCrLf & Task_Date & vbCrLf & "was DELETED successfully"
End Select
 '''''''''''''''''''''''''''''''''''
 
 Answer = MsgBox(strQuestion, vbCritical + vbYesNo, EM_TITLE)
 If Answer <> vbYes Then Exit Sub
 '''''''''''''''''''''''''''''''''''''''''
 S = Replace(strSql, "delete", "select * ")
 R1.Open S, Con, adOpenKeyset
 ReDim Preserve A(R1.RecordCount - 1)
 While Not R1.EOF
  A(J) = R1.Fields("task") & DLMT & R1.Fields("date_time") & DLMT & R1.Fields("priority")
  J = J + 1
  R1.MoveNext
 Wend
 R1.Close
 '''''''''''''''''''''''''''''''''''''
 ''  dump it into schedule_history     '''
 S = "select * from schedule_history"
 R1.Open S, Con, adOpenDynamic, adLockPessimistic
 For J = LBound(A) To UBound(A)
  V = Split(A(J), DLMT)
  R1.AddNew
  R1.Fields("task") = V(0)
  R1.Fields("date_time") = V(1)
  R1.Fields("priority") = V(2)
  R1.Update
 Next
 R1.Close
 '''''''''''''''''''''''''''''
 Con.Execute strSql
 '''''''''''''''''''''''''''''
 MsgBox strMsg, vbOKOnly + vbInformation, EM_TITLE
 Task_ID = -1
 Task_Name = vbNullString
 Task_Date = vbNullString
 frm_Progress_Bar.Delay_Time = 10
 frm_Progress_Bar.Show vbModal
 Populate_Task Tble
 ''''''''''''''''''''''''''''''''
 Exit Sub
EH:
 MsgBox Err.Number & vbCrLf & Err.Description, , EM_TITLE
End Sub

Private Sub mnu_History_Click()
 lbl_Menus(1).Visible = True
 ''''''''''''''''''''''''''''''''''
 If mnu_History.Caption = "History" Then
  mnu_History.Caption = "Current"
  lbl_Menus(4).Caption = "Current"
  Tble = "schedule_history"
  mnuDelete.Visible = False
  lbl_Menus(1).Visible = False
 Else
  mnu_History.Caption = "History"
  lbl_Menus(4).Caption = "History"
  Tble = "schedule"
 End If
  '''''''''''''''''''''''
 txt_Task.Text = vbNullString
 cmd_Schedule.Enabled = Not cmd_Schedule.Enabled
 Populate_Task Tble
End Sub

Private Sub mnu_Refresh_Click()
 Populate_Task Tble
End Sub

 
Private Sub mnu_Sort_By_Click(Index As Integer)
 Select Case Index
  Case 0: Sort_By = "priority"
  Case 1: Sort_By = "date_time"
  Case 2: Sort_By = "task"
 End Select
 For i = 0 To mnu_Sort_By.Count - 1
  mnu_Sort_By(i).Checked = False
  mnu_Sort_By(i).Enabled = True
 Next
  mnu_Sort_By(Index).Checked = True
  mnu_Sort_By(Index).Enabled = False
  '''''''''''''''
  Populate_Task Tble
End Sub

Private Sub mnu_Sort_Click(Index As Integer)
 Select Case Index
   Case 0: Asc_Desc = "asc"
   Case 1: Asc_Desc = "desc"
 End Select
 For i = 0 To mnu_Sort.Count - 1
  mnu_Sort(i).Checked = False
  mnu_Sort(i).Enabled = True
 Next
  mnu_Sort(Index).Checked = True
  mnu_Sort(Index).Enabled = False
  '''''''''''''''
  Populate_Task Tble
End Sub

Private Sub lvw_task_ItemClick(ByVal Item As MSComctlLib.ListItem)
  Task_Name = Trim(Item.Text)
  Task_Priority = Item.SubItems(1)
  Task_Date = Item.SubItems(2)
  Task_ID = Item.SubItems(3)
  txt_Task.Text = Task_Name
End Sub

Private Sub Populate_Task(TBL As String)
 Dim LI As ListItem
 Dim R1 As New ADODB.Recordset
 Dim R2 As New ADODB.Recordset
 '''''''''''''''''''''
 On Error GoTo EH
 '''''''''''''''''''''''''''''
 lvw_Task.ListItems.Clear
 Sort_Task_by_Priority
 'Delay 5
 '''''''''''''''''''''''''''''
 strSql = "select * from " & TBL & " order by " & Sort_By & " " & Asc_Desc
 R1.Open strSql, Con
 While Not R1.EOF
  '''''''''''''''''''''''''''''''''''
   Set LI = lvw_Task.ListItems.Add(, , R1.Fields("task"))
   LI.ListSubItems.Add , , R1.Fields("priority")
   LI.ListSubItems.Add , , Format(R1.Fields("date_time"), dt_Format)
   LI.ListSubItems.Add , , R1.Fields("id")
  R1.MoveNext
 Wend
 R1.Close
 '''''''''''''''''''''''''
 lbl_Task.Caption = lvw_Task.ListItems.Count
 '''''''''''''''''''''''''
 Exit Sub
EH:
 txt_Display.Text = Err.Description
End Sub


Private Sub Sort_Task_by_Priority()
Dim J As Long
Dim R1 As New ADODB.Recordset
''''''''''''
On Error GoTo EH

strSql = "select *  from schedule order by priority"
With R1
 .Open strSql, Con, adOpenKeyset, adLockOptimistic
 While Not .EOF
   J = J + 1
  .Fields("priority") = J
  '.Update
  .MoveNext
 Wend
 .Close
 End With
 ''''''''''''''''''''
 Exit Sub
EH:
 txt_Display.Text = Err.Description
End Sub

Private Sub DELAY(Tm As Long)
 Dim T1 As Double
 Dim T2 As Double
 T1 = Timer
 While (Timer - T1) < Tm
  DoEvents ''' do nothing just wait........
 Wend

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnu_Format_Click(Index As Integer)
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
  
 Case 1:
  cdl_File.ShowColor
  If cdl_File.Color = 0 Then Exit Sub
  Font_Color = cdl_File.Color
  txt_Task.ForeColor = Font_Color
 
 Case 2:
  cdl_File.ShowColor
  If cdl_File.Color = 0 Then Exit Sub
  Back_Color = cdl_File.Color
  txt_Task.BackColor = Back_Color

End Select

End Sub

Private Sub img_Save_Preference_Click()
 Save_Preference arr_Windows(3)
End Sub

Private Sub Save_Preference(arr_Win As String)

 Dim R1 As New ADODB.Recordset
 Dim lng_RecCount As Long
  On Error GoTo EH
 ''''''''''''''''''''''''''''''
 strSql = "Select * from preferences where em_category='" & arr_Win & "'"
With R1
 .Open strSql, Con, adOpenDynamic, adLockOptimistic
 If .EOF Then .AddNew
 .Fields(1) = arr_Win
 .Fields(2) = Font_Name & DLMT & Font_Bold & DLMT & Font_Italic & DLMT & Font_Underline & DLMT & Font_Strikethru & DLMT & Font_Size & DLMT & Font_Color & DLMT & Back_Color
 .Update
 .Close
End With
txt_Display.Text = "Schedular Preference was saved successfully...."
''''''''''''''
Exit Sub
EH:
 txt_Display.Text = Err.Description
End Sub



Private Sub Retrieve_Preference(arr_Win As String)
 Dim R1 As New ADODB.Recordset
 Dim lng_RecCount As Long
 Dim arr_Pref As Variant
 On Error GoTo EH
   
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
  Back_Color = arr_Pref(7)
  '''close the connection
  R1.Close
Else
   Font_Name = "Arial"
   Font_Bold = False
   Font_Italic = False
   Font_Underline = False
   Font_Strikethru = False
   Font_Size = 8
   Font_Color = vbBlack
   Back_Color = vbWhite
End If
 ''''''''''''''''
 Exit Sub
EH:
 txt_Display.Text = Err.Description
End Sub

Private Sub Format_Skin()
 txt_Task.FontName = Font_Name
 txt_Task.FontBold = Font_Bold
 txt_Task.FontItalic = Font_Italic
 txt_Task.FontUnderline = Font_Underline
 txt_Task.FontStrikethru = Font_Strikethru
 txt_Task.FontSize = Font_Size
 txt_Task.ForeColor = Font_Color
 txt_Task.BackColor = Back_Color
 cmd_BG_Color.BackColor = Back_Color
End Sub

Private Sub cmd_Font_Color_Click()
 cdl_File.ShowColor
 If cdl_File.Color = 0 Then Exit Sub
 Font_Color = cdl_File.Color
 txt_Task.ForeColor = Font_Color
End Sub

Private Sub cmd_BG_Color_Click()
  cdl_File.ShowColor
 If cdl_File.Color = 0 Then Exit Sub
 Back_Color = cdl_File.Color
 txt_Task.BackColor = Back_Color
 cmd_BG_Color.BackColor = Back_Color
End Sub


Private Sub cmd_Font_Click(Index As Integer)
 On Error GoTo EH
 Select Case Index
  Case 0:
   If Font_Size > 20 Then Exit Sub
   Font_Size = Font_Size + 1
  Case 1:
   If Font_Size < 8 Then Exit Sub
   Font_Size = Font_Size - 1
  Case 2:
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
 End Select
   txt_Task.FontSize = Font_Size
 Exit Sub
EH:
 MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Private Sub chk_Font_Click(Index As Integer)
 Select Case Index
 Case 0:
   txt_Task.FontBold = Not txt_Task.FontBold
   Font_Bold = txt_Task.FontBold
 Case 1:
   txt_Task.FontItalic = Not txt_Task.FontItalic
   Font_Italic = txt_Task.FontItalic
 Case 2:
   txt_Task.FontUnderline = Not txt_Task.FontUnderline
   Font_Underline = txt_Task.FontUnderline
 Case 3:
   txt_Task.FontStrikethru = Not txt_Task.FontStrikethru
   Font_Strikethru = txt_Task.FontStrikethru
 End Select
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Write_Word_DOC(File_Name As String)
Dim oWordApp As Word.Application
Set oWordApp = New Word.Application
Dim wTable As Table
Dim Doc As Document
Dim P As Long, Q As Long
Dim Answer As VbMsgBoxResult

Dim FRow As Integer, TRow As Integer, TCol As Integer
'''''''''''''''
On Error GoTo EH
''''''''''''''''''''''''''''''''''''
txt_Display.Text = "Wait..." & vbCrLf & File_Name & " InProcess.."
FRow = 2
TCol = lvw_Task.ColumnHeaders.Count + 1
TRow = lvw_Task.ListItems.Count
''''''''''''
With oWordApp
 ''Create a new document
 Set Doc = .Documents.Add
 Doc.PageSetup.TopMargin = Application.InchesToPoints(Val(Top_Margin))
 Doc.PageSetup.BottomMargin = Application.InchesToPoints(Val(Bottom_Margin))
 Doc.PageSetup.LeftMargin = Application.InchesToPoints(Val(Left_Margin))
 Doc.PageSetup.RightMargin = Application.InchesToPoints(Val(Right_Margin))
 Doc.PageSetup.Orientation = wdOrientLandscape
 '''''''''''''''''''''
 
 Set wTable = Doc.Tables.Add(Doc.Range, FRow, TCol - 1)
 wTable.Borders.Enable = Table_Border
 wTable.Range.ParagraphFormat.Alignment = Paragraph_Alignment
  'Add text to the document
 wTable.Select
  
 '''''' Heading information...  '''''''''''
 'wTable.Cell(1, TCol - 1).Width = wTable.Cell(1, TCol - 1).Width / 4 '' reduce the size of serial number column
 wTable.Cell(1, 1).Merge wTable.Cell(1, TCol - 1)
 wTable.Cell(1, 1).Range.Bold = True
 wTable.Cell(1, 1).Range.Font.Size = 14
 wTable.Cell(1, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
 wTable.Cell(1, 1).Range.InsertAfter "List of Task (Jobs...) " & Now
 wTable.Cell(1, 1).Range.InsertParagraphAfter
 wTable.AllowAutoFit = True
  
 ''''''''''Column heading....'''''''''''''''''''
  P = 2: Q = 1
  wTable.Cell(P, Q).Range.InsertAfter "Ser #"
  wTable.Cell(P, Q).Range.Bold = True
  wTable.Cell(P, Q).Width = wTable.Cell(P, Q).Width / 4 '' reduce the width of serial # column
  '''''''''''''''''
  Q = 2
  wTable.Cell(P, Q).Range.InsertAfter lvw_Task.ColumnHeaders(Q - 1).Text '' task
  wTable.Cell(P, Q).Range.Bold = True
  wTable.Cell(P, Q).Width = (wTable.Cell(P, Q).Width * (15 / 12)) + (wTable.Cell(P, Q).Width)  '' enlarge the width of task column
  ''''''''''''''''''
  Q = 3
  wTable.Cell(P, Q).Range.InsertAfter lvw_Task.ColumnHeaders(Q - 1).Text '' priority column
  wTable.Cell(P, Q).Range.Bold = True
  wTable.Cell(P, Q).Width = wTable.Cell(P, Q).Width / 3 '' reduce the width of priority column
  '''''''''''''''''
  For Q = 4 To lvw_Task.ColumnHeaders.Count  '' last column is id
   wTable.Cell(P, Q).Range.InsertAfter lvw_Task.ColumnHeaders(Q - 1).Text
   wTable.Cell(P, Q).Range.Bold = True
  Next
 ''''''''''''''''''''''''''''''''''
  
  For P = 1 To TRow
   wTable.Rows.Add
   wTable.Rows(P + FRow).Range.Bold = False
  Next
  '''''''''''''''''''''''''''''''''
  
 For P = 1 To TRow
  Q = 1
  wTable.Cell(FRow + P, Q).Range.InsertAfter P
  ''wTable.Cell(FRow + P, Q).Width = wTable.Cell(FRow + P, Q).Width / 4 '' reduce the width of serial number column
  wTable.Cell(FRow + P, Q + 1).Range.InsertAfter lvw_Task.ListItems(P).Text '' Task value...
  ''''''''''''
  For Q = 3 To TCol - 1 '' last column is id
   wTable.Cell(FRow + P, Q).Range.InsertAfter lvw_Task.ListItems(P).SubItems(Q - 2)
   wTable.Cell(FRow + P, Q).Range.Bold = False
  Next
 Next
 ''''''''''''''''''''''''''''''''''''''''''''''''
 
 'Save the document
.ActiveDocument.SaveAs fileName:=File_Name, _
  FileFormat:=wdFormatDocument, LockComments:=False, _
  Password:="", AddToRecentFiles:=True, WritePassword _
  :="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
  SaveNativePictureFormat:=False, SaveFormsData:=False, _
  SaveAsAOCELetter:=False
  
'Answer = MsgBox("wanna print it out...", vbOKCancel + vbInformation, ICEMS_TITLE)
'.PrintOut , , wTable.Range

End With
'''''''''''''''
oWordApp.Quit
txt_Display.Text = File_Name & " Created Successfully...."
''''''''''''''''''''''''''''''''''
Exit Sub
EH:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, EM_TITLE

End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''


Private Sub Write_Web_Page(File_Name As String)
 Dim FSO As New FileSystemObject
 Dim TS As TextStream
 Dim strData As String
 Dim P As Long, Q As Long
 Dim TRow As Integer, TCol As Integer
'''''''''''''''
On Error GoTo EH
''''''''''''''''''''''''''''''''''''
Set TS = FSO.OpenTextFile(File_Name, ForWriting, True)
strData = "<html>"
strData = strData & "<head><title>" & EM_TITLE & " Leave System" & "</title></head>"
strData = strData & "<body>"
strData = strData & "<table width='100%' border=2 bordercolor=blue>"
''''''''''''''''''''''''''''''''''''
txt_Display.Text = "Wait.." & vbCrLf & File_Name & " InProcess.."
TCol = lvw_Task.ColumnHeaders.Count - 1
TRow = lvw_Task.ListItems.Count
''''''''''  Employee Name with Designation ..'''''''''''''''''''
'''''''''''''''''''''''''''''''''
 strData = strData & "<tr align=center >"
 strData = strData & "<th colspan=" & TCol & ">"
 strData = strData & "(" & "Task(s)" & ")</th>"
 strData = strData & "</tr>"
'''''''''''''''''''''''''''''''''''''''''''''''
''''''''''Column heading....'''''''''''''''''''
 P = 2:
 strData = strData & "<tr bgcolor=teal>"
 For Q = 1 To TCol  '' last column is id
  strData = strData & "<th>" & lvw_Task.ColumnHeaders(Q).Text & "</th>"
 Next
 strData = strData & "</tr>"
 ''''''''''''''''''''''''''''''''''
 ''''''''''Data Rows ....'''''''''''''''''''
 strData = strData & "<tr>"
 For P = 1 To TRow
  strData = strData & "<td>" & lvw_Task.ListItems(P) & "</td>"
  For Q = 2 To TCol   '' last column is id
   strData = strData & "<td>" & lvw_Task.ListItems(P).SubItems(Q - 1) & "</td>"
  Next
  strData = strData & "</tr>"
 Next
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '''''''''' Date & Time Stamp RowData Rows ....'''''''''''
 strData = strData & "<tr align=center >"
 strData = strData & "<td colspan=" & TCol & ">Record generated  " & Now & "</td>"
 strData = strData & "</tr>"
 ''''''''''''''''''''''''''''''''''''''''''''''''
 strData = strData & "</table></body></html>"
   'Save the document
 TS.Write strData
txt_Display.Text = "Web Page file '" & File_Name & "' Created successfully...."
''''''''''''''''''''''''''''''''''
Set TS = Nothing
Set FSO = Nothing
Exit Sub
EH:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, EM_TITLE

End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub mnu_Word_Page_Setup_Click()
 frm_Page_Setup.Show vbModal
End Sub


Private Sub txt_Display_DblClick()
 txt_Display.Text = vbNullString
End Sub

Private Sub mnu_windows_Click(Index As Integer)
 Load_Selected_Window Me, Index
End Sub

