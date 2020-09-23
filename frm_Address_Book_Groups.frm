VERSION 5.00
Begin VB.Form frm_Address_Book_Groups 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "frm_Address_Book_Groups.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   4455
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_Display 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H00FF0000&
      Height          =   525
      Left            =   30
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   3030
      Width           =   4425
   End
   Begin VB.Frame fra_Delete 
      Caption         =   "Delete Address Book Group"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   795
      Left            =   60
      TabIndex        =   2
      Top             =   2190
      Width           =   4365
      Begin VB.CommandButton cmd_Group_Del 
         Caption         =   "Delete Group"
         Height          =   300
         Left            =   2670
         TabIndex        =   8
         Top             =   300
         Width           =   1575
      End
      Begin VB.ComboBox cbo_Group_Del 
         Height          =   315
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "List of Designation"
         Top             =   300
         Width           =   2385
      End
   End
   Begin VB.Frame fra_Edit 
      Caption         =   "Edit Address Book Group"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1185
      Left            =   60
      TabIndex        =   1
      Top             =   930
      Width           =   4365
      Begin VB.ComboBox cbo_Group_Edit 
         Height          =   315
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   9
         ToolTipText     =   "List of Designation"
         Top             =   300
         Width           =   2385
      End
      Begin VB.CommandButton cmd_Group_Update 
         Caption         =   "Update Group"
         Height          =   300
         Left            =   2670
         TabIndex        =   6
         Top             =   780
         Width           =   1575
      End
      Begin VB.TextBox txt_Group_Update 
         Height          =   300
         Left            =   150
         TabIndex        =   5
         ToolTipText     =   "Update Designation"
         Top             =   780
         Width           =   2355
      End
   End
   Begin VB.Frame fra_Add 
      Caption         =   "Add Address Book Group"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   4365
      Begin VB.CommandButton cmd_Group_Add 
         Caption         =   "Add Group"
         Height          =   300
         Left            =   2640
         TabIndex        =   4
         Top             =   300
         Width           =   1515
      End
      Begin VB.TextBox txt_Group_Add 
         Height          =   300
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "New Designation"
         Top             =   300
         Width           =   2355
      End
   End
End
Attribute VB_Name = "frm_Address_Book_Groups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arr_Index() As Integer, i As Integer
Dim Index_Del As Integer, Index_Update As Integer

Private Sub cmd_Add_Group_Click()
 On Error Resume Next
 txt_Display.Text = Err.Number & Space(2) & strMsg
End Sub

Private Sub cbo_Group_Del_Click()
 Index_Del = arr_Index(cbo_Group_Del.ListIndex)
End Sub

Private Sub cbo_Group_edit_Click()
 Index_Update = arr_Index(cbo_Group_Edit.ListIndex)
End Sub

Private Sub cmd_Group_Add_Click()
 Dim R1 As New ADODB.Recordset
 On Error GoTo EH
 ''''''''''''''''''''''''
 If Trim(txt_Group_Add.Text) = vbNullString Then
  txt_Display.Text = "You did not provide any value to add ... "
  Exit Sub
 End If
 '''''''''''''''''''''''
 strSql = "Select * from address_book_group"
 With R1
 .Open strSql, Con, adOpenDynamic, 3
 .AddNew
 .Fields(1) = txt_Group_Add.Text
 .Update
 .Close
 End With
 txt_Display.Text = txt_Group_Add.Text & " added successfully...."
 txt_Group_Add.Text = vbNullString
 Populate_Data
 '''''''''''''''''''''''''''''
 Exit Sub
EH:
 txt_Display.Text = Err.Description
End Sub

Private Sub cmd_Group_Update_Click()
 Dim R1 As New ADODB.Recordset
 On Error GoTo EH
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If Index_Update = -1 Then
  txt_Display.Text = "You did not select the GROUP to update... "
  Exit Sub
 End If
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If Trim(txt_Group_Update.Text) = vbNullString Then
  txt_Display.Text = "You did not supply the new GROUP "
  Exit Sub
 End If
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''
 strSql = "select * from address_book_group where id=" & Index_Update
 With R1
  .Open strSql, Con, adOpenDynamic, adLockPessimistic
   .Fields(1) = Trim(txt_Group_Update.Text)
 .Update
 .Close
 End With
 '''''''''''''''''''''''''''''''''
 MsgBox "Group has been updated successfully....", vbInformation, EM_TITLE
 txt_Group_Update.Text = vbNullString
 Populate_Data
 Exit Sub
EH:
 txt_Display.Text = Err.Description
End Sub

Private Sub Form_Load()
 strMsg = vbNullString ''' must be at the top statement in load event
 Me.Caption = EM_TITLE & ".." & "Address Book" & EM_MAIL
 Populate_Data
End Sub

Private Sub Populate_Data()
 Dim R1 As New ADODB.Recordset
 Dim str_Data As String
 On Error GoTo EH
 strSql = "select * from address_book_group order by name"
 R1.Open strSql, Con, 3, 3
 
 cbo_Group_Edit.Clear
 cbo_Group_Del.Clear
 i = 0
 ReDim arr_Index(R1.RecordCount)
 
 While Not R1.EOF
  str_Data = R1.Fields(1)
  cbo_Group_Edit.AddItem str_Data
  cbo_Group_Del.AddItem str_Data
  arr_Index(i) = R1.Fields(0)
  i = i + 1
  R1.MoveNext
 Wend
 R1.Close
 ''''''''''''''''''''
 Index_Del = -1
 Index_Update = -1
 Exit Sub
EH:
 txt_Display.Text = Err.Description
End Sub



Private Sub cmd_Group_Del_Click()
 Dim R1 As New ADODB.Recordset
 Dim Answer As VbMsgBoxResult
On Error GoTo EH
 ''''''''''''''''''''''''''
 If Trim(cbo_Group_Del.Text) = vbNullString Then
  txt_Display.Text = "You did not select the GROUP to delete... "
  Exit Sub
 End If
 ''''''''''''''''''''''''''
 If Index_Del = -1 Then Exit Sub
 
 strSql = "select id from address_book_group_link where group_id=" & Index_Del
 R1.Open strSql, Con, 3, 3
 If R1.RecordCount > 0 Then
  MsgBox "'" & R1.RecordCount & "'" & vbCrLf & vbCrLf & "records are attached to this Group...(" & cbo_Group_Del.Text & ")" & vbCrLf & vbCrLf & "So First change the Person's Records Please...", vbCritical, EM_TITLE
  R1.Close
  Exit Sub
 End If
  R1.Close
 ''''''''''''''''''''''
 Answer = MsgBox("Are you sure to ERASE group of " & vbCrLf & "'" & cbo_Group_Del.Text & "'", vbCritical + vbYesNo, EM_TITLE)
 If Answer <> vbYes Then Exit Sub
 ''''''''''''
 strSql = "delete from address_book_group where id=" & Index_Del
 Con.Execute strSql
 ''''''''''''''''''''''''''
 MsgBox cbo_Group_Del.Text & vbCrLf & " Group Deleted successfully....", vbOKOnly, EM_TITLE
 Index_Del = -1
 frm_Progress_Bar.Show vbModal
 Populate_Data
 Exit Sub
EH:
 MsgBox Err.Number & vbCrLf & Err.Description, , EM_TITLE
End Sub


Private Sub txt_Display_DblClick()
 txt_Display.Text = vbNullString
End Sub
